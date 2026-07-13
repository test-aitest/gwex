"""巨大 .xlsm/.xlsx の1シート構造を typed JSON で返す（pure）。

47MB 超のマクロ付き画面設計書を「開かずに見る」ための目。ZIP エントリの
bytes dict を受け取り、対象シートのセクション（B列の ■ マーカー）・値セル・
埋め込み画像・図形・コネクタを列挙する。ファイル I/O は行わない
（読み込みは呼び出し側が xlsx_zip._read_zip 等で行う）。

設計メモ:
- セル文字列は sharedStrings の ふりがなラン <rPh>/<phoneticPr> を除外して
  <t> を連結し、NFC 正規化して返す（NFD/NFC 罠の回避）。
- images/shapes の from は [row, col]（どちらも 0 始まり、アンカー XML の値）。
- ext_pt は表示寸法を EMU→pt（/12700）換算した概算。
- cells_mode="sparse"（既定）は「B列のセル or 結合セルの左上」だけを返し、
  巨大シートでも JSON が数十KB に収まるようにする。"full" は全値セル。
"""

from __future__ import annotations

import html
import re
import unicodedata
from typing import Optional

EMU_PER_PT = 12700

_MARKER = "■"


def _nfc(s: str) -> str:
    return unicodedata.normalize("NFC", s)


def _unescape(s: str) -> str:
    return html.unescape(s)


def col_to_index(letters: str) -> int:
    """列文字 → 1始まりの列番号（A=1）。"""
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def sheet_xml_path(entries: dict[str, bytes], sheet_name: str) -> str:
    """workbook.xml → workbook.xml.rels からシート XML のエントリパスを引く。"""
    wb = entries["xl/workbook.xml"].decode("utf-8")
    m = re.search(r'<sheet[^>]*name="' + re.escape(sheet_name) + r'"[^>]*r:id="(rId\d+)"', wb)
    if not m:
        raise ValueError(f"シートが見つかりません: {sheet_name}")
    rid = m.group(1)
    rels = entries["xl/_rels/workbook.xml.rels"].decode("utf-8")
    rel = re.search(r'<Relationship\b[^>]*\bId="' + re.escape(rid) + r'"[^>]*/>', rels)
    if not rel:
        raise ValueError(f"worksheet rel が見つかりません: {rid}")
    mt = re.search(r'Target="([^"]+)"', rel.group(0))
    if not mt:
        raise ValueError(f"worksheet rel に Target がありません: {rid}")
    target = mt.group(1).lstrip("/")
    return target if target.startswith("xl/") else "xl/" + target


def drawing_path_for_sheet(entries: dict[str, bytes], sheet_xml: str) -> Optional[tuple[str, str]]:
    """シートに配線済みの (drawing パス, drawing rels パス)。無ければ None。"""
    base = sheet_xml.split("/")[-1]
    rels_path = f"xl/worksheets/_rels/{base}.rels"
    rels = entries.get(rels_path, b"").decode("utf-8")
    m = re.search(r'Target="([^"]*drawings/drawing\d+\.xml)"', rels) if rels else None
    if not m:
        return None
    draw = m.group(1).split("/")[-1]
    return f"xl/drawings/{draw}", f"xl/drawings/_rels/{draw}.rels"


def strip_phonetic(inner_xml: str) -> str:
    """si / is の中身から <rPh>（ふりがな）と <phoneticPr> を除いて <t> を連結する。"""
    s = re.sub(r"<rPh\b.*?</rPh>", "", inner_xml, flags=re.DOTALL)
    s = re.sub(r"<phoneticPr\b[^>]*/>", "", s)
    return _unescape("".join(re.findall(r"<t(?:\s[^>]*)?>(.*?)</t>", s, re.DOTALL)))


def shared_strings(entries: dict[str, bytes]) -> list[str]:
    """sharedStrings.xml → インデックス順の文字列リスト（rPh 除去・NFC 正規化済み）。"""
    raw = entries.get("xl/sharedStrings.xml")
    if raw is None:
        return []
    text = raw.decode("utf-8")
    return [_nfc(strip_phonetic(si)) for si in re.findall(r"<si>(.*?)</si>", text, re.DOTALL)]


def _num(v: str):
    try:
        return int(v)
    except ValueError:
        try:
            return float(v)
        except ValueError:
            return v


def cell_value(attrs: str, inner: str, sst: list[str]):
    """<c> の属性文字列と中身から値を返す（文字列は NFC・数値はそのまま数値）。"""
    tm = re.search(r'\bt="(\w+)"', attrs)
    t = tm.group(1) if tm else None
    if t == "s":
        vm = re.search(r"<v>(\d+)</v>", inner)
        if vm is None:
            return None
        i = int(vm.group(1))
        return sst[i] if i < len(sst) else None
    if t == "inlineStr":
        im = re.search(r"<is>(.*)</is>", inner, re.DOTALL)
        return _nfc(strip_phonetic(im.group(1))) if im else None
    if t == "str":
        vm = re.search(r"<v(?:\s[^>]*)?>(.*?)</v>", inner, re.DOTALL)
        return _nfc(_unescape(vm.group(1))) if vm else None
    if t == "b":
        vm = re.search(r"<v>([01])</v>", inner)
        return bool(int(vm.group(1))) if vm else None
    if t == "e":
        vm = re.search(r"<v>([^<]*)</v>", inner)
        return vm.group(1) if vm else None
    vm = re.search(r"<v>([^<]*)</v>", inner)
    return _num(vm.group(1)) if vm else None


# row 開きタグを属性値対応で取る（値内の > や自己終了 /> を正しく区別する）。
# <row\b([^>]*)(?:/>|>(.*?)</row>) の一発 regex は、自己終了 <row .../> で貪欲な
# [^>]* が末尾 / を食い「次の行の </row> まで」飲み込む（次行のセルが自己終了行に
# 誤帰属する破損クラスの真因）ため使わない。
_ROW_OPEN_RE = re.compile(r"<row\b((?:[^>\"']|\"[^\"]*\"|'[^']*')*?)(/?)>")


def iter_row_blocks(text: str):
    """(attrs, body, start, end) を document 順に返す。自己終了 <row/> は body=""。"""
    for m in _ROW_OPEN_RE.finditer(text):
        if m.group(2):
            yield m.group(1), "", m.start(), m.end()
            continue
        close = text.find("</row>", m.end())
        if close == -1:
            yield m.group(1), text[m.end():], m.start(), len(text)
        else:
            yield m.group(1), text[m.end():close], m.start(), close + len("</row>")


def iter_cells(sheet_xml_text: str):
    """(row:int, ref:str, col_letters:str, attrs:str, inner:str) を document 順に返す。"""
    for row_attrs, body, _start, _end in iter_row_blocks(sheet_xml_text):
        rr = re.search(r'\br="(\d+)"', row_attrs)
        if not rr or not body:
            continue
        row = int(rr.group(1))
        for cm in re.finditer(r"<c\b([^>]*?)(?:/>|>(.*?)</c>)", body, re.DOTALL):
            attrs = cm.group(1)
            ref_m = re.search(r'\br="([A-Z]+)(\d+)"', attrs)
            if not ref_m:
                continue
            yield row, ref_m.group(1) + ref_m.group(2), ref_m.group(1), attrs, cm.group(2) or ""


def merged_ranges(sheet_xml_text: str) -> list[str]:
    return re.findall(r'<mergeCell ref="([A-Z0-9:]+)"/>', sheet_xml_text)


_ANCHOR_RE = re.compile(
    r"<(?:xdr:)?(oneCellAnchor|twoCellAnchor|absoluteAnchor)\b[^>]*>(.*?)</(?:xdr:)?\1>",
    re.DOTALL,
)


def _anchor_from(block: str) -> Optional[list[int]]:
    m = re.search(
        r"<(?:xdr:)?from>\s*<(?:xdr:)?col>(\d+)</(?:xdr:)?col>.*?"
        r"<(?:xdr:)?row>(\d+)</(?:xdr:)?row>",
        block, re.DOTALL,
    )
    return [int(m.group(2)), int(m.group(1))] if m else None  # [row, col]


def _ext_pt(block: str) -> Optional[list[int]]:
    m = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"/>', block) or \
        re.search(r'<(?:xdr:)?ext cx="(\d+)" cy="(\d+)"/>', block)
    if not m:
        return None
    return [round(int(m.group(1)) / EMU_PER_PT), round(int(m.group(2)) / EMU_PER_PT)]


def _cnvpr(block: str) -> tuple[Optional[int], str]:
    m = re.search(r'<(?:xdr:)?cNvPr [^>]*?id="(\d+)"[^>]*?name="([^"]*)"', block)
    if not m:
        m2 = re.search(r'<(?:xdr:)?cNvPr [^>]*?name="([^"]*)"[^>]*?id="(\d+)"', block)
        if not m2:
            return None, ""
        return int(m2.group(2)), _nfc(_unescape(m2.group(1)))
    return int(m.group(1)), _nfc(_unescape(m.group(2)))


def shape_text(block: str) -> str:
    """図形ブロック内のテキスト（<a:t> 連結・NFC）。"""
    return _nfc("".join(_unescape(t) for t in re.findall(r"<a:t>(.*?)</a:t>", block, re.DOTALL)))


def _shape_labels(dx: str) -> dict[int, dict]:
    """cNvPr id → {kind, text}（グループ内の図形も含む。コネクタ接続先の解決用）。"""
    out: dict[int, dict] = {}
    for m in re.finditer(r"<(?:xdr:)?sp[ >].*?</(?:xdr:)?sp>", dx, re.DOTALL):
        sid, name = _cnvpr(m.group(0))
        if sid is not None:
            out[sid] = {"kind": "sp", "text": shape_text(m.group(0)) or name}
    for m in re.finditer(r"<(?:xdr:)?pic[ >].*?</(?:xdr:)?pic>", dx, re.DOTALL):
        pid, name = _cnvpr(m.group(0))
        if pid is not None:
            out[pid] = {"kind": "pic", "text": name}
    return out


def _cxn_endpoint(m: Optional[re.Match], labels: dict[int, dict]) -> Optional[dict]:
    if not m:
        return None
    d: dict = {"id": int(m.group(1)), "idx": int(m.group(2))}
    info = labels.get(d["id"])
    if info:
        d["kind"] = info["kind"]
        d["text"] = info["text"]
    return d


def media_for_rid(entries: dict[str, bytes], draw_rels: str, rid: str) -> Optional[str]:
    rels = entries.get(draw_rels, b"").decode("utf-8")
    rel = re.search(r'<Relationship\b[^>]*\bId="' + re.escape(rid) + r'"[^>]*/>', rels)
    if not rel:
        return None
    mt = re.search(r'Target="([^"]+)"', rel.group(0))
    if not mt:
        return None
    target = mt.group(1)
    if "../" in target:
        return "xl/" + target.split("../", 1)[1]
    return target.lstrip("/")


def scan(entries: dict[str, bytes], sheet: str, *, cells_mode: str = "sparse") -> dict:
    """1シートの構造を dict で返す（JSON 化可能）。"""
    if cells_mode not in ("sparse", "full"):
        raise ValueError(f"cells_mode は sparse | full: {cells_mode!r}")
    xml_path = sheet_xml_path(entries, sheet)
    sx = entries[xml_path].decode("utf-8")
    sst = shared_strings(entries)

    merges = merged_ranges(sx)
    merge_top: dict[str, str] = {}
    for ref in merges:
        merge_top[ref.split(":")[0]] = ref

    sections: list[dict] = []
    cells: list[dict] = []
    max_row = 0
    max_col = 0
    for row, ref, col_letters, attrs, inner in iter_cells(sx):
        val = cell_value(attrs, inner, sst)
        if val is None or val == "":
            continue
        max_row = max(max_row, row)
        max_col = max(max_col, col_to_index(col_letters))
        if col_letters == "B" and isinstance(val, str) and val.lstrip().startswith(_MARKER):
            sections.append({"marker": val.strip(), "row": row})
        if cells_mode == "sparse" and col_letters != "B" and ref not in merge_top:
            continue
        entry: dict = {"ref": ref, "text": val}
        if ref in merge_top:
            entry["merged"] = merge_top[ref]
        cells.append(entry)

    images: list[dict] = []
    shapes: list[dict] = []
    connectors: list[dict] = []
    counts = {"pic": 0, "sp": 0, "cxnSp": 0, "grpSp": 0, "merges": len(merges)}
    found = drawing_path_for_sheet(entries, xml_path)
    if found:
        draw_path, draw_rels = found
        dx = entries[draw_path].decode("utf-8")
        counts["pic"] = len(re.findall(r"<(?:xdr:)?pic[ >]", dx))
        counts["sp"] = len(re.findall(r"<(?:xdr:)?sp[ >]", dx))
        counts["cxnSp"] = len(re.findall(r"<(?:xdr:)?cxnSp[ >]", dx))
        counts["grpSp"] = len(re.findall(r"<(?:xdr:)?grpSp[ >]", dx))
        labels = _shape_labels(dx) if counts["cxnSp"] else {}
        for am in _ANCHOR_RE.finditer(dx):
            kind, block = am.group(1), am.group(2)
            frm = _anchor_from(block)
            if re.search(r"<(?:xdr:)?grpSp[ >]", block):
                continue  # グループは counts のみ（中の個々の図形は列挙しない）
            if re.search(r"<(?:xdr:)?cxnSp[ >]", block):
                sid, _name = _cnvpr(block)
                prst = re.search(r'prst="(\w+)"', block)
                st = re.search(r'<a:stCxn id="(\d+)" idx="(\d+)"', block)
                end = re.search(r'<a:endCxn id="(\d+)" idx="(\d+)"', block)
                xfm = re.search(r'<a:xfrm([^>]*)><a:off x="(-?\d+)" y="(-?\d+)"/>'
                                r'<a:ext cx="(\d+)" cy="(\d+)"/>', block)
                connectors.append({
                    "id": sid,
                    "preset": prst.group(1) if prst else None,
                    "st": _cxn_endpoint(st, labels),
                    "end": _cxn_endpoint(end, labels),
                    "xfrm": ({"off": [int(xfm.group(2)), int(xfm.group(3))],
                              "ext": [int(xfm.group(4)), int(xfm.group(5))],
                              "flipH": 'flipH="1"' in xfm.group(1),
                              "flipV": 'flipV="1"' in xfm.group(1)} if xfm else None),
                })
            elif re.search(r"<(?:xdr:)?pic[ >]", block):
                pid, name = _cnvpr(block)
                rid = re.search(r'r:embed="(rId\d+)"', block)
                images.append({
                    "id": pid,
                    "name": name,
                    "anchor": kind,
                    "from": frm,
                    "ext_pt": _ext_pt(block),
                    "media": media_for_rid(entries, draw_rels, rid.group(1)) if rid else None,
                })
            elif re.search(r"<(?:xdr:)?sp[ >]", block):
                sid, _name = _cnvpr(block)
                prst = re.search(r'prst="(\w+)"', block)
                shapes.append({
                    "id": sid,
                    "kind": "sp",
                    "preset": prst.group(1) if prst else None,
                    "text": shape_text(block),
                    "anchor": kind,
                    "from": frm,
                })

    return {
        "sheet": sheet,
        "xml_path": xml_path,
        "dims": {"max_row": max_row, "max_col": max_col},
        "sections": sections,
        "cells": cells,
        "images": images,
        "shapes": shapes,
        "connectors": connectors,
        "counts": counts,
    }


def summarize(result: dict) -> str:
    lines = [
        f"シート: {result['sheet']} ({result['xml_path']})  "
        f"最大 {result['dims']['max_row']}行 × {result['dims']['max_col']}列",
        "セクション: " + (" / ".join(f"{s['marker']}@{s['row']}" for s in result["sections"]) or "(なし)"),
        "counts: " + " ".join(f"{k}={v}" for k, v in result["counts"].items()),
        f"値セル: {len(result['cells'])}件 / 画像: {len(result['images'])}件 / "
        f"図形: {len(result['shapes'])}件 / コネクタ: {len(result['connectors'])}件",
    ]
    for im in result["images"]:
        ext = f"{im['ext_pt'][0]}x{im['ext_pt'][1]}pt" if im.get("ext_pt") else "?"
        lines.append(f"  [pic {im['id']}] {im['name']} {im['anchor']} from={im['from']} {ext} {im['media']}")
    for sp in result["shapes"][:40]:
        txt = (sp["text"][:24] + "…") if len(sp["text"]) > 24 else sp["text"]
        lines.append(f"  [sp {sp['id']}] {sp['preset']} {txt!r} from={sp['from']}")
    if len(result["shapes"]) > 40:
        lines.append(f"  … 他 {len(result['shapes']) - 40} 図形")

    def _ep(d: Optional[dict]) -> str:
        if not d:
            return "(座標)"
        t = str(d.get("text", ""))
        return f"id{d['id']}" + (f":{t[:14]!r}" if t else "")

    for cn in result["connectors"][:40]:
        lines.append(f"  [cxn {cn['id']}] {cn['preset']} {_ep(cn.get('st'))} → {_ep(cn.get('end'))}")
    if len(result["connectors"]) > 40:
        lines.append(f"  … 他 {len(result['connectors']) - 40} コネクタ")
    return "\n".join(lines)
