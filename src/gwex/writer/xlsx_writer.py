"""ローカル .xlsx への書き戻し（openpyxl）。privileged。

既定は in-place 上書き。`output` 指定時のみ別名保存。
"""

from __future__ import annotations

from typing import Optional

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_to_tuple

from gwex.domains.model import DEFAULT_MAPPING, MappingConfig
from gwex.writer._image import downscale

_DATA_ROLES = (
    "screen_name", "medium_category", "small_category",
    "test_no", "verification_content", "execution_steps",
)


def _anchor(ws, cell: str) -> str:
    """cell が結合範囲内なら左上アンカーに解決する（結合セルは左上のみ書込可）。"""
    row, col = coordinate_to_tuple(cell)
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return f"{get_column_letter(rng.min_col)}{rng.min_row}"
    return cell


def set_cell_text(
    path: str, sheet: str, cell: str, text: str, *, output: Optional[str] = None
) -> str:
    wb = load_workbook(path)
    ws = wb[sheet]
    ws[_anchor(ws, cell)] = text
    dest = output or path
    wb.save(dest)
    return dest


_COL_PX = 64.0   # 既定列幅相当(px)
_ROW_PX = 20.0   # 既定行高相当(px)


def _range_px(ws, min_c, min_r, max_c, max_r) -> tuple[int, int]:
    """結合範囲の概算ピクセルサイズ（列幅/行高の設定があれば反映）。"""
    w = 0.0
    for c in range(min_c, max_c + 1):
        cw = ws.column_dimensions[get_column_letter(c)].width
        w += cw * 7.0 if cw else _COL_PX
    h = 0.0
    for r in range(min_r, max_r + 1):
        rh = ws.row_dimensions[r].height
        h += rh * 1.333 if rh else _ROW_PX
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
    output: Optional[str] = None,
) -> str:
    """画像をセルの上にオーバーレイで乗せる。

    cell_range（例 "C7:F20"）指定時は、その範囲を画像枠として**結合**し、画像を範囲サイズに
    フィットさせて乗せる。insert_rows=True なら枠ぶんの行を先に挿入して既存内容を押し下げ、
    確実に空の枠を作る（既存レイアウトを壊さない）。cell_range 省略時は anchor に浮かせるだけ。
    """
    from openpyxl.utils.cell import range_boundaries

    wb = load_workbook(path)
    ws = wb[sheet]
    img = XLImage(downscale(image_path, max_dim))

    if cell_range:
        min_c, min_r, max_c, max_r = range_boundaries(cell_range)
        nrows = max_r - min_r + 1
        if insert_rows:
            ws.insert_rows(min_r, nrows)  # 枠ぶんの空行を挿入（既存は下へ）
        # 範囲に重なる既存結合を解除してから結合し直す
        for rng in list(ws.merged_cells.ranges):
            if not (rng.max_col < min_c or rng.min_col > max_c or rng.max_row < min_r or rng.min_row > max_r):
                ws.unmerge_cells(str(rng))
        ws.merge_cells(start_row=min_r, start_column=min_c, end_row=max_r, end_column=max_c)
        fw, fh = _range_px(ws, min_c, min_r, max_c, max_r)
        iw, ih = img.width, img.height
        scale = min(fw / iw, fh / ih) if iw and ih else 1.0
        img.width = int(iw * scale)
        img.height = int(ih * scale)
        img.anchor = f"{get_column_letter(min_c)}{min_r}"
    else:
        if width:
            img.width = width
        if height:
            img.height = height
        img.anchor = anchor

    ws.add_image(img)
    dest = output or path
    wb.save(dest)
    return dest


def write_cells(
    path: str, sheet: str, assignments: list[tuple[str, object]], *, output: Optional[str] = None
) -> str:
    """[(セル, 値), ...] を一括書込（構造書き戻し用）。"""
    wb = load_workbook(path)
    ws = wb[sheet]
    for cell, value in assignments:
        ws[_anchor(ws, cell)] = value
    dest = output or path
    wb.save(dest)
    return dest


def last_data_row(path: str, sheet: str, mapping: Optional[MappingConfig] = None) -> int:
    """testspec のデータ列に値がある最終行（1始まり）を返す。データ無しなら data_start_row-1。"""
    cfg = mapping or DEFAULT_MAPPING
    cols = [i for i in (cfg.column_index(r) for r in _DATA_ROLES) if i is not None]
    wb = load_workbook(path, read_only=True)
    ws = wb[sheet]
    last = cfg.data_start_row - 1
    for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if r_idx < cfg.data_start_row:
            continue
        if any(c < len(row) and row[c] is not None for c in cols):
            last = r_idx
    return last
