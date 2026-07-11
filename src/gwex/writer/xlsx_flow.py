"""画面遷移図（スクリーンショット+矢印+ラベル）を xlsx に生成する（privileged）。

概要設計書とは別の成果物。スクショをノードとして格子配置し（列=遷移ステップ、
行=同一画面のバリエーション）、遷移を矢印で結ぶ。トリガーボタンの赤枠・
条件ラベル・処理ボックス（矢印の途中に挟む）も描く。

openpyxl は図形を保存できないため、描画要素はすべて xlsx_zip と同じ ZIP 注入
（DrawingML）で書く。Excel はセル寸法を超えるアンカーオフセットをクランプする
ため配置は absoluteAnchor に統一し、Quick Look 等の簡易レンダラ向けに spPr の
a:xfrm（同値の絶対座標）を必ず併記する。

spec（YAML/JSON を dict 化したもの）:

    sheet: 振込で借りるフロー          # シート名（省略時 画面遷移図）
    grid:                              # レイアウト上書き（px、省略可）
      node_width: 200
      col_gap: 130
      row_gap: 60
      origin_x: 40
      origin_y: 70
      label_gap: 44
    nodes:
      - {id: N0001, name: ダッシュボード, image: caps/dashboard.png, col: 0, row: 0}
    edges:
      - from: N0001
        to: N0701
        trigger_rect: [35, 505, 190, 73]   # 元画像px。赤枠+矢印の起点（省略可）
        label: 振込で借りる                 # 条件ラベル（省略可）
        via: 振込融資の実行                 # 途中に挟む処理ボックス（省略可）
"""

from __future__ import annotations

import os
from typing import Optional
from xml.sax.saxutils import escape

from gwex.writer import xlsx_zip

_EMU = 9525
_EMU_PT = 12700

_A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
_R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

_DEFAULT_GRID = {
    "node_width": 200,  # スクショ表示幅(px)
    "col_gap": 130,     # 列間隔 = 矢印・処理ボックスの領域
    "row_gap": 60,
    "origin_x": 40,
    "origin_y": 70,
    "label_gap": 44,    # 画像上部の 画面ID/画面名 ラベル領域
}

_BOX_W, _BOX_H = 120, 36      # 処理ボックス既定寸法(px)
_LABEL_W, _LABEL_H = 160, 18  # 条件ラベル既定寸法(px)


def _E(px: float) -> int:
    return round(px * _EMU)


def _xfrm(x: float, y: float, w: float, h: float, flips: str = "") -> str:
    return (f'<a:xfrm {_A}{flips}><a:off x="{_E(x)}" y="{_E(y)}"/>'
            f'<a:ext cx="{max(1, _E(w))}" cy="{max(1, _E(h))}"/></a:xfrm>')


def _anchor(p: str, x: float, y: float, w: float, h: float, inner: str) -> str:
    return (f'<{p}absoluteAnchor><{p}pos x="{_E(x)}" y="{_E(y)}"/>'
            f'<{p}ext cx="{max(1, _E(w))}" cy="{max(1, _E(h))}"/>'
            f'{inner}<{p}clientData/></{p}absoluteAnchor>')


def _pic(p: str, pid: int, name: str, rid: str, x: float, y: float, w: float, h: float) -> str:
    inner = (
        f'<{p}pic><{p}nvPicPr><{p}cNvPr id="{pid}" name="{name}"/>'
        f'<{p}cNvPicPr><a:picLocks {_A} noChangeAspect="1"/></{p}cNvPicPr></{p}nvPicPr>'
        f'<{p}blipFill><a:blip {_A} {_R} r:embed="{rid}"/>'
        f'<a:stretch {_A}><a:fillRect/></a:stretch></{p}blipFill>'
        f'<{p}spPr>{_xfrm(x, y, w, h)}'
        f'<a:prstGeom {_A} prst="rect"><a:avLst/></a:prstGeom></{p}spPr></{p}pic>'
    )
    return _anchor(p, x, y, w, h, inner)


def _text(p: str, pid: int, name: str, x: float, y: float, w: float, h: float,
          lines: list[str], *, size: int = 1000, bold: bool = False,
          color: str = "000000", align: str = "l", v_anchor: str = "t") -> str:
    """透明のテキスト図形（ラベル用）。size は 1/100pt。"""
    b = ' b="1"' if bold else ""
    paras = "".join(
        f'<a:p {_A}><a:pPr algn="{align}"/><a:r><a:rPr lang="ja-JP" sz="{size}"{b}>'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:rPr>'
        f'<a:t>{escape(t)}</a:t></a:r></a:p>'
        for t in lines
    )
    inner = (
        f'<{p}sp><{p}nvSpPr><{p}cNvPr id="{pid}" name="{name}"/><{p}cNvSpPr/></{p}nvSpPr>'
        f'<{p}spPr>{_xfrm(x, y, w, h)}<a:prstGeom {_A} prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:noFill {_A}/><a:ln {_A}><a:noFill/></a:ln></{p}spPr>'
        f'<{p}txBody><a:bodyPr {_A} anchor="{v_anchor}" lIns="0" tIns="0" rIns="0" bIns="0" wrap="none"/>'
        f'<a:lstStyle {_A}/>{paras}</{p}txBody></{p}sp>'
    )
    return _anchor(p, x, y, w, h, inner)


def _frame(p: str, pid: int, name: str, x: float, y: float, w: float, h: float,
           *, color: str = "FF0000", line_pt: float = 1.75) -> str:
    """塗りなしの枠矩形（トリガーボタンの赤枠用）。"""
    inner = (
        f'<{p}sp><{p}nvSpPr><{p}cNvPr id="{pid}" name="{name}"/><{p}cNvSpPr/></{p}nvSpPr>'
        f'<{p}spPr>{_xfrm(x, y, w, h)}<a:prstGeom {_A} prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:noFill {_A}/><a:ln {_A} w="{round(line_pt * _EMU_PT)}">'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln></{p}spPr></{p}sp>'
    )
    return _anchor(p, x, y, w, h, inner)


def _box(p: str, pid: int, name: str, x: float, y: float, w: float, h: float, text: str,
         *, fill: str = "9DC3E6", border: str = "2E74B5", size: int = 1000) -> str:
    """塗りつぶし+中央テキストの処理ボックス。"""
    inner = (
        f'<{p}sp><{p}nvSpPr><{p}cNvPr id="{pid}" name="{name}"/><{p}cNvSpPr/></{p}nvSpPr>'
        f'<{p}spPr>{_xfrm(x, y, w, h)}<a:prstGeom {_A} prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill {_A}><a:srgbClr val="{fill}"/></a:solidFill>'
        f'<a:ln {_A}><a:solidFill><a:srgbClr val="{border}"/></a:solidFill></a:ln></{p}spPr>'
        f'<{p}txBody><a:bodyPr {_A} anchor="ctr" lIns="0" tIns="0" rIns="0" bIns="0"/>'
        f'<a:lstStyle {_A}/><a:p {_A}><a:pPr algn="ctr"/><a:r><a:rPr lang="ja-JP" sz="{size}">'
        f'<a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:rPr>'
        f'<a:t>{escape(text)}</a:t></a:r></a:p></{p}txBody></{p}sp>'
    )
    return _anchor(p, x, y, w, h, inner)


def _connector(p: str, pid: int, name: str,
               x1: float, y1: float, x2: float, y2: float,
               *, color: str = "404040", line_pt: float = 1.5,
               dash: bool = False) -> str:
    """(x1,y1)→(x2,y2) の矢印。水平/垂直なら直線、それ以外はカギ線（bentConnector3）。

    dash=True で点線（モーダル表示などの補助遷移用）。
    """
    x, y = min(x1, x2), min(y1, y2)
    w, h = abs(x2 - x1), abs(y2 - y1)
    prst = "straightConnector1" if (h < 4 or w < 4) else "bentConnector3"
    flips = (' flipH="1"' if x2 < x1 else "") + (' flipV="1"' if y2 < y1 else "")
    dash_xml = '<a:prstDash val="dash"/>' if dash else ""
    inner = (
        f'<{p}cxnSp macro=""><{p}nvCxnSpPr><{p}cNvPr id="{pid}" name="{name}"/>'
        f'<{p}cNvCxnSpPr/></{p}nvCxnSpPr>'
        f'<{p}spPr>{_xfrm(x, y, w, h, flips)}<a:prstGeom {_A} prst="{prst}"><a:avLst/></a:prstGeom>'
        f'<a:ln {_A} w="{round(line_pt * _EMU_PT)}">'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>{dash_xml}'
        f'<a:tailEnd type="triangle"/></a:ln></{p}spPr></{p}cxnSp>'
    )
    return _anchor(p, x, y, w, h, inner)


def draw_flow(spec: dict, output: str, *, base_dir: Optional[str] = None) -> str:
    """spec から画面遷移図の xlsx を新規生成する（既存ファイルは上書き）。

    ノードの画像パスは base_dir（通常 spec ファイルのディレクトリ）からの相対を許容。
    生成後に openpyxl 系コマンドで再保存すると図形が消えるため、遷移図は編集せず
    成果物として扱う（作り直しは spec を変えて再実行）。
    """
    from openpyxl import Workbook
    from PIL import Image as PILImage

    nodes = spec.get("nodes") or []
    if not nodes:
        raise ValueError("spec.nodes が空です")
    edges = spec.get("edges") or []
    sheet = spec.get("sheet", "画面遷移図")
    g = {**_DEFAULT_GRID, **(spec.get("grid") or {})}

    # レイアウト: 表示寸法と絶対座標(px)を確定
    laid: dict[str, dict] = {}
    w_node = int(g["node_width"])
    max_h = 1
    for n in nodes:
        path = n["image"]
        if base_dir and not os.path.isabs(path):
            path = os.path.join(base_dir, path)
        with PILImage.open(path) as im:
            iw, ih = im.size
        h = max(1, round(w_node * ih / iw))
        max_h = max(max_h, h)
        laid[n["id"]] = {**n, "path": path, "iw": iw, "w": w_node, "h": h}
    pitch_y = max_h + g["label_gap"] + g["row_gap"]
    for nd in laid.values():
        nd["x"] = g["origin_x"] + int(nd["col"]) * (w_node + g["col_gap"])
        nd["y"] = g["origin_y"] + g["label_gap"] + int(nd["row"]) * pitch_y

    # 新規ブック（枠線非表示の専用シート）
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    ws.sheet_view.showGridLines = False
    wb.save(output)

    entries = xlsx_zip._read_zip(output)
    sheet_xml = xlsx_zip._sheet_xml_path(entries, sheet)
    draw_path, draw_rels = xlsx_zip._drawing_for_sheet(entries, sheet_xml)
    dx = entries[draw_path].decode("utf-8")
    p = "xdr:" if "<xdr:wsDr" in dx else ""

    # 画像 media の登録（同一パスは1エントリに共有）
    rid_by_path: dict[str, str] = {}
    for nd in laid.values():
        if nd["path"] in rid_by_path:
            continue
        midx = xlsx_zip._next_media_index(entries)
        ext = os.path.splitext(nd["path"])[1].lstrip(".").lower() or "png"
        media = f"xl/media/image{midx}.{ext}"
        with open(nd["path"], "rb") as f:
            entries[media] = f.read()
        xlsx_zip._ensure_png_content_type(entries)
        rid_by_path[nd["path"]] = xlsx_zip._add_rel(
            entries, draw_rels,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
            f"../media/image{midx}.{ext}")

    parts: list[str] = []
    pid = 100  # 新規 drawing なので固定開始で衝突しない

    # ノード: 画像 + 画面ID/画面名ラベル
    for nd in laid.values():
        parts.append(_pic(p, pid, f'node_{nd["id"]}', rid_by_path[nd["path"]],
                          nd["x"], nd["y"], nd["w"], nd["h"]))
        pid += 1
        lines = [f'画面ID：　{nd["id"]}', f'画面名：　{nd.get("name", "")}']
        parts.append(_text(p, pid, f'label_{nd["id"]}',
                           nd["x"], nd["y"] - g["label_gap"],
                           nd["w"] + g["col_gap"] - 10, g["label_gap"] - 4, lines, size=900))
        pid += 1

    # エッジ: 赤枠 → (処理ボックス) → 矢印 → 条件ラベル
    for i, e in enumerate(edges, start=1):
        a, b = laid[e["from"]], laid[e["to"]]
        if e.get("trigger_rect"):
            tx, ty, tw, th = e["trigger_rect"]
            s = a["w"] / a["iw"]  # 元画像px → 表示px
            fx, fy, fw, fh = a["x"] + tx * s, a["y"] + ty * s, tw * s, th * s
            parts.append(_frame(p, pid, f"trigger{i}", fx, fy, fw, fh))
            pid += 1
            sx, sy = fx + fw, fy + fh / 2  # 赤枠の右端中央から矢印を出す
        else:
            sx, sy = a["x"] + a["w"], a["y"] + a["h"] / 2
        ex, ey = b["x"], b["y"] + b["h"] / 2

        # detour（モーダル等への寄り道）は慣例に合わせて既定で点線。dash で明示上書き可
        dash = bool(e.get("dash", bool(e.get("detour"))))
        if e.get("via"):
            bx = (sx + ex) / 2 - _BOX_W / 2
            by = (sy + ey) / 2 - _BOX_H / 2
            parts.append(_box(p, pid, f"process{i}", bx, by, _BOX_W, _BOX_H, e["via"]))
            pid += 1
            parts.append(_connector(p, pid, f"edge{i}a", sx, sy, bx, by + _BOX_H / 2, dash=dash))
            pid += 1
            parts.append(_connector(p, pid, f"edge{i}b", bx + _BOX_W, by + _BOX_H / 2, ex, ey, dash=dash))
            pid += 1
            label_y = by - _LABEL_H - 4
        else:
            parts.append(_connector(p, pid, f"edge{i}", sx, sy, ex, ey, dash=dash))
            pid += 1
            label_y = min(sy, ey) - _LABEL_H - 4

        if e.get("label"):
            lx = (sx + ex) / 2 - _LABEL_W / 2
            parts.append(_text(p, pid, f"edgelabel{i}", lx, label_y, _LABEL_W, _LABEL_H,
                               [e["label"]], size=900, align="ctr", v_anchor="b"))
            pid += 1

    close = f"</{p}wsDr>"
    entries[draw_path] = dx.replace(close, "".join(parts) + close).encode("utf-8")
    xlsx_zip._write_zip(entries, output)
    return output
