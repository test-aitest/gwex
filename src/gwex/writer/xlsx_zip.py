"""xlsx を ZIP として扱い、既存内容に最小タッチで over-cell 画像を注入する（privileged）。

なぜ openpyxl(load+save) でなくこれを使うか：openpyxl の round-trip は**任意の実 Excel** で
グラフ・ピボット・スライサー・フォームコントロール・VBA(xlsm) 等の高度要素をドロップ/破損
させる既知の弱点があり、また既存 media を再圧縮・再採番し sharedStrings/calcChain/xl/metadata
を書き換える。本モジュールは対象シート＋drawing 以外の ZIP エントリに一切触れず、既存の
media/書式/高度要素を**バイト単位で温存**する（対象が単純な doc なら openpyxl でも実害は無いが、
任意 doc への安全側の既定として ZIP 注入を採る）。

※「openpyxl は埋め込み画像を破壊する」という旧前提は誤り。現行 openpyxl は画像(ICC含む)を
保持する（破壊はしないが上記の書き換え/高度要素ロスは起きる）。

`set_cell_image` は xlsx_writer.set_cell_image と同じシグネチャ・同じ振る舞い
（cell_range 枠フィット＝アスペクト維持、width/height 明示、重複結合の解除、
insert_rows で枠ぶんの行挿入）を、既存画像を壊さずに実現する。
"""

from __future__ import annotations

import io
import re
import zipfile
from typing import Optional

from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries

from gwex.writer._image import downscale

_EMU = 9525  # 1px
_COL_PX = 64.0   # 既定列幅相当(px)。xlsx_writer と一致させる
_ROW_PX = 20.0   # 既定行高相当(px)


def _read_zip(path: str) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}


def _write_zip(entries: dict[str, bytes], dest: str) -> None:
    with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in entries.items():
            z.writestr(name, data)


def _sheet_xml_path(entries: dict[str, bytes], sheet_name: str) -> str:
    wb = entries["xl/workbook.xml"].decode("utf-8")
    m = re.search(r'<sheet[^>]*name="' + re.escape(sheet_name) + r'"[^>]*r:id="(rId\d+)"', wb)
    if not m:
        raise ValueError(f"シートが見つかりません: {sheet_name}")
    rid = m.group(1)
    rels = entries["xl/_rels/workbook.xml.rels"].decode("utf-8")
    # 属性順は実装依存（Target が Id より前のこともある）ので、当該 Relationship 要素を
    # 取り出してから Target を読む。
    rel = re.search(r'<Relationship\b[^>]*\bId="' + re.escape(rid) + r'"[^>]*/>', rels)
    if not rel:
        raise ValueError(f"worksheet rel が見つかりません: {rid}")
    mt = re.search(r'Target="([^"]+)"', rel.group(0))
    if not mt:
        raise ValueError(f"worksheet rel に Target がありません: {rid}")
    target = mt.group(1).lstrip("/")
    return target if target.startswith("xl/") else "xl/" + target


def _next_media_index(entries: dict[str, bytes]) -> int:
    idx = 0
    for n in entries:
        m = re.match(r"xl/media/image(\d+)\.", n)
        if m:
            idx = max(idx, int(m.group(1)))
    return idx + 1


def _ensure_png_content_type(entries: dict[str, bytes]) -> None:
    ct = entries["[Content_Types].xml"].decode("utf-8")
    if 'Extension="png"' not in ct:
        ct = ct.replace("</Types>", '<Default Extension="png" ContentType="image/png"/></Types>')
        entries["[Content_Types].xml"] = ct.encode("utf-8")


def _drawing_for_sheet(entries: dict[str, bytes], sheet_xml: str) -> tuple[str, str]:
    """シートの drawing パスと drawing rels パスを返す。無ければ新規作成して配線する。"""
    base = sheet_xml.split("/")[-1]  # sheet2.xml
    rels_path = f"xl/worksheets/_rels/{base}.rels"
    rels = entries.get(rels_path, b"").decode("utf-8")
    m = re.search(r'Target="([^"]*drawings/drawing\d+\.xml)"', rels) if rels else None
    if m:
        draw = m.group(1).split("/")[-1]
        return f"xl/drawings/{draw}", f"xl/drawings/_rels/{draw}.rels"

    # 新規 drawing を作成
    nums = [int(x) for n in entries for x in re.findall(r"xl/drawings/drawing(\d+)\.xml$", n)]
    dn = (max(nums) + 1) if nums else 1
    draw_path = f"xl/drawings/drawing{dn}.xml"
    draw_rels = f"xl/drawings/_rels/drawing{dn}.xml.rels"
    entries[draw_path] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"></xdr:wsDr>'
    ).encode("utf-8")
    entries[draw_rels] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    ).encode("utf-8")
    # シートに drawing rel と <drawing> 要素を追加
    rid = _add_rel(entries, rels_path,
                   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
                   f"../drawings/drawing{dn}.xml")
    sx = entries[sheet_xml].decode("utf-8")
    if "<drawing " not in sx:
        # <worksheet> に r: 名前空間が無い（新規 openpyxl シート等）と <drawing r:id=>
        # の r: が未束縛になりパース不能。無ければ宣言を補う。
        wm = re.match(r"(<worksheet\b[^>]*?)>", sx)
        if wm and "xmlns:r=" not in wm.group(1):
            sx = sx.replace(
                wm.group(0),
                wm.group(1)
                + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
                1,
            )
        sx = re.sub(r"</worksheet>", f'<drawing r:id="{rid}"/></worksheet>', sx)
        entries[sheet_xml] = sx.encode("utf-8")
    # [Content_Types] に drawing を宣言
    ct = entries["[Content_Types].xml"].decode("utf-8")
    if f"/xl/drawings/drawing{dn}.xml" not in ct:
        ct = ct.replace("</Types>",
                        f'<Override PartName="/xl/drawings/drawing{dn}.xml" '
                        'ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>')
        entries["[Content_Types].xml"] = ct.encode("utf-8")
    return draw_path, draw_rels


def _add_rel(entries: dict[str, bytes], rels_path: str, rtype: str, target: str) -> str:
    rels = entries.get(rels_path, b"").decode("utf-8")
    if not rels:
        rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
    existing = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels)]
    rid = f"rId{(max(existing) + 1) if existing else 1}"
    rel = f'<Relationship Id="{rid}" Type="{rtype}" Target="{target}"/>'
    rels = rels.replace("</Relationships>", rel + "</Relationships>")
    entries[rels_path] = rels.encode("utf-8")
    return rid


def _add_merge(entries: dict[str, bytes], sheet_xml: str, ref: str) -> None:
    sx = entries[sheet_xml].decode("utf-8")
    m = re.search(r'<mergeCells count="(\d+)">', sx)
    if m:
        cnt = int(m.group(1)) + 1
        sx = sx.replace(m.group(0), f'<mergeCells count="{cnt}">', 1)
        sx = sx.replace("</mergeCells>", f'<mergeCell ref="{ref}"/></mergeCells>', 1)
    else:
        # sheetData の後ろに mergeCells を新設
        sx = re.sub(r"</sheetData>",
                    f'</sheetData><mergeCells count="1"><mergeCell ref="{ref}"/></mergeCells>', sx, count=1)
    entries[sheet_xml] = sx.encode("utf-8")


def _remove_overlapping_merges(entries: dict[str, bytes], sheet_xml: str,
                               min_c: int, min_r: int, max_c: int, max_r: int) -> None:
    """指定範囲と重なる既存 mergeCell を取り除く（重複結合の解除）。"""
    sx = entries[sheet_xml].decode("utf-8")
    removed = 0

    def keep(m):
        nonlocal removed
        a, b = m.group(1), m.group(2)
        c1, r1 = coordinate_to_tuple(a)[1], coordinate_to_tuple(a)[0]
        if b:
            c2, r2 = coordinate_to_tuple(b)[1], coordinate_to_tuple(b)[0]
        else:
            c2, r2 = c1, r1
        lo_c, hi_c, lo_r, hi_r = min(c1, c2), max(c1, c2), min(r1, r2), max(r1, r2)
        overlap = not (hi_c < min_c or lo_c > max_c or hi_r < min_r or lo_r > max_r)
        if overlap:
            removed += 1
            return ""
        return m.group(0)

    sx = re.sub(r'<mergeCell ref="([A-Z]+\d+)(?::([A-Z]+\d+))?"/>', keep, sx)
    if removed:
        mc = re.search(r'<mergeCells count="(\d+)">', sx)
        if mc:
            new_cnt = int(mc.group(1)) - removed
            if new_cnt <= 0:
                sx = re.sub(r'<mergeCells count="\d+">\s*</mergeCells>', "", sx)
            else:
                sx = sx.replace(mc.group(0), f'<mergeCells count="{new_cnt}">', 1)
    entries[sheet_xml] = sx.encode("utf-8")


def _shift_rows_down(entries: dict[str, bytes], sheet_xml: str, draw_path: str,
                     min_r: int, nrows: int) -> None:
    """min_r 以降の行を nrows 行ぶん下へずらす（既存内容を押し下げ、空枠を作る）。

    sheetData の <row r=> と セル <c r=>、mergeCell の参照、既存 drawing アンカーの
    <xdr:row>（0始まり）をまとめてシフトする。
    """
    sx = entries[sheet_xml].decode("utf-8")

    # 1) 行タグ <row ... r="N" ...>
    def shift_rowtag(m):
        n = int(m.group(2))
        return f'{m.group(1)}{n + nrows if n >= min_r else n}{m.group(3)}'
    sx = re.sub(r'(<row\b[^>]*?\br=")(\d+)(")', shift_rowtag, sx)

    # 2) セル参照 <c r="B7" ...>（列文字＋行番号）
    def shift_cellref(m):
        col, n = m.group(1), int(m.group(2))
        return f'r="{col}{n + nrows if n >= min_r else n}"'
    sx = re.sub(r'\br="([A-Z]+)(\d+)"', shift_cellref, sx)

    # 3) mergeCell ref="A5:C8"
    def shift_one(ref):
        cm = re.match(r'([A-Z]+)(\d+)', ref)
        col, n = cm.group(1), int(cm.group(2))
        return f'{col}{n + nrows if n >= min_r else n}'

    def shift_merge(m):
        parts = m.group(1).split(":")
        return '<mergeCell ref="' + ":".join(shift_one(p) for p in parts) + '"/>'
    sx = re.sub(r'<mergeCell ref="([A-Z0-9:]+)"/>', shift_merge, sx)
    entries[sheet_xml] = sx.encode("utf-8")

    # 4) 既存 drawing アンカー <xdr:row>N</xdr:row>（0始まり）
    if draw_path in entries:
        dx = entries[draw_path].decode("utf-8")
        thr = min_r - 1  # 0始まりしきい値

        def shift_anchor(m):
            pre, n = m.group(1), int(m.group(2))
            return f'<{pre}row>{n + nrows if n >= thr else n}</{pre}row>'
        # xdr: 接頭辞（Excel）と接頭辞なし（openpyxl）の両対応
        dx = re.sub(r'<(xdr:)?row>(\d+)</(?:xdr:)?row>', shift_anchor, dx)
        entries[draw_path] = dx.encode("utf-8")


def _sheet_defaults(sx: str) -> tuple[float, float]:
    """defaultColWidth(文字幅), defaultRowHeight(pt) を返す（無ければ None 相当）。"""
    m = re.search(r'<sheetFormatPr\b[^>]*>', sx)
    dc = dr = None
    if m:
        tag = m.group(0)
        mc = re.search(r'defaultColWidth="([\d.]+)"', tag)
        mr = re.search(r'defaultRowHeight="([\d.]+)"', tag)
        if mc:
            dc = float(mc.group(1))
        if mr:
            dr = float(mr.group(1))
    return dc, dr


def _col_widths(sx: str) -> dict[int, float]:
    """列インデックス(1始まり)→ 文字幅。<cols><col min max width> から。"""
    widths: dict[int, float] = {}
    cols_block = re.search(r'<cols>(.*?)</cols>', sx, re.DOTALL)
    if not cols_block:
        return widths
    for m in re.finditer(r'<col\b[^>]*?min="(\d+)"[^>]*?max="(\d+)"[^>]*?width="([\d.]+)"[^>]*?>', cols_block.group(1)):
        lo, hi, w = int(m.group(1)), int(m.group(2)), float(m.group(3))
        for c in range(lo, hi + 1):
            widths[c] = w
    return widths


def _row_heights(sx: str) -> dict[int, float]:
    """行番号(1始まり)→ 高さ(pt)。<row r ht> から。"""
    heights: dict[int, float] = {}
    for m in re.finditer(r'<row\b[^>]*?\br="(\d+)"[^>]*?\bht="([\d.]+)"', sx):
        heights[int(m.group(1))] = float(m.group(2))
    return heights


def _frame_px(sx: str, min_c: int, min_r: int, max_c: int, max_r: int) -> tuple[int, int]:
    """結合範囲の概算ピクセルサイズ（xlsx_writer._range_px と同じ換算）。"""
    cw = _col_widths(sx)
    rh = _row_heights(sx)
    dc, dr = _sheet_defaults(sx)
    w = 0.0
    for c in range(min_c, max_c + 1):
        width = cw.get(c, dc)
        w += width * 7.0 if width else _COL_PX
    h = 0.0
    for r in range(min_r, max_r + 1):
        height = rh.get(r, dr)
        h += height * 1.333 if height else _ROW_PX
    return int(w), int(h)


def set_cell_image(
    path: str,
    sheet: str,
    anchor: str,
    image_path: str,
    *,
    cell_range: Optional[str] = None,
    insert_rows: bool = False,
    width: Optional[int] = None,
    height: Optional[int] = None,
    max_dim: Optional[int] = None,
    col_off_px: int = 0,
    row_off_px: int = 0,
    output: Optional[str] = None,
) -> str:
    """既存内容（埋め込み画像含む）を保持して over-cell 画像を注入する。

    xlsx_writer.set_cell_image と同じ振る舞い:
    - cell_range 指定時はその範囲を結合し、画像を枠サイズへ**アスペクト維持**でフィット。
    - insert_rows=True で枠ぶんの行を先に挿入（既存は下へ）。
    - cell_range 省略時は anchor に width/height（または自然サイズ）で浮かせる。
    """
    from PIL import Image as PILImage

    src = downscale(image_path, max_dim)
    with PILImage.open(src) as im:
        iw, ih = im.size
    with open(src, "rb") as f:
        img_bytes = f.read()

    entries = _read_zip(path)
    sheet_xml = _sheet_xml_path(entries, sheet)
    draw_path, draw_rels = _drawing_for_sheet(entries, sheet_xml)

    # 配置先セルとサイズの決定
    if cell_range:
        min_c, min_r, max_c, max_r = range_boundaries(cell_range)
        nrows = max_r - min_r + 1
        if insert_rows:
            _shift_rows_down(entries, sheet_xml, draw_path, min_r, nrows)
        _remove_overlapping_merges(entries, sheet_xml, min_c, min_r, max_c, max_r)
        merge_ref = f"{get_column_letter(min_c)}{min_r}:{get_column_letter(max_c)}{max_r}"
        _add_merge(entries, sheet_xml, merge_ref)
        # 枠にアスペクト維持でフィット
        sx = entries[sheet_xml].decode("utf-8")
        fw, fh = _frame_px(sx, min_c, min_r, max_c, max_r)
        scale = min(fw / iw, fh / ih) if iw and ih else 1.0
        disp_w, disp_h = max(1, int(iw * scale)), max(1, int(ih * scale))
        col0, row0 = min_c - 1, min_r - 1
    else:
        row, col = coordinate_to_tuple(anchor)  # 返り値は (row, col)
        col0, row0 = col - 1, row - 1
        disp_w = width if width else iw
        disp_h = height if height else ih

    # media 追加
    midx = _next_media_index(entries)
    media_name = f"xl/media/image{midx}.png"
    entries[media_name] = img_bytes
    _ensure_png_content_type(entries)

    # drawing rels に画像 rel 追加
    rid = _add_rel(entries, draw_rels,
                   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                   f"../media/image{midx}.png")

    cx, cy = disp_w * _EMU, disp_h * _EMU
    pic_id = 100000 + midx
    dx = entries[draw_path].decode("utf-8")
    # 既存 drawing が xdr: 接頭辞か、デフォルト名前空間（接頭辞なし）かを検出して合わせる。
    # （Excel 由来=xdr:、openpyxl 由来=接頭辞なし。混在は不正 XML になるため必ず一致させる）
    p = "xdr:" if "<xdr:wsDr" in dx else ""
    a_ns = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    r_ns = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    anchor_xml = (
        f'<{p}oneCellAnchor>'
        f'<{p}from><{p}col>{col0}</{p}col><{p}colOff>{col_off_px * _EMU}</{p}colOff>'
        f'<{p}row>{row0}</{p}row><{p}rowOff>{row_off_px * _EMU}</{p}rowOff></{p}from>'
        f'<{p}ext cx="{cx}" cy="{cy}"/>'
        f'<{p}pic>'
        f'<{p}nvPicPr><{p}cNvPr id="{pic_id}" name="capture{midx}"/>'
        f'<{p}cNvPicPr><a:picLocks {a_ns} noChangeAspect="1"/></{p}cNvPicPr></{p}nvPicPr>'
        f'<{p}blipFill><a:blip {a_ns} {r_ns} r:embed="{rid}"/>'
        f'<a:stretch {a_ns}><a:fillRect/></a:stretch></{p}blipFill>'
        f'<{p}spPr><a:prstGeom {a_ns} prst="rect"><a:avLst/></a:prstGeom></{p}spPr>'
        f'</{p}pic><{p}clientData/></{p}oneCellAnchor>'
    )
    close = "</xdr:wsDr>" if p else "</wsDr>"
    dx = dx.replace(close, anchor_xml + close)
    entries[draw_path] = dx.encode("utf-8")

    dest = output or path
    _write_zip(entries, dest)
    return dest


def add_image_over_cells(
    path: str,
    sheet: str,
    anchor: str,
    image_path: str,
    *,
    cell_range: Optional[str] = None,
    max_dim: Optional[int] = None,
    output: Optional[str] = None,
) -> str:
    """後方互換ラッパ。oneCellAnchor で自然サイズ配置（cell_range 指定時は結合のみ追加）。"""
    return set_cell_image(
        path, sheet, anchor, image_path,
        cell_range=cell_range, insert_rows=False, max_dim=max_dim, output=output,
    )
