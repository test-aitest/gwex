"""Google Sheets (`spreadsheets.get(includeGridData=true)` の JSON) → IR。

pure 層。各シートを `Heading(2)` + `Table` に。`formattedValue`（表示値）ベース、
1 行目をヘッダ扱い、全体で空の列・行は畳む。リンクは `hyperlink` を反映。
"""

from __future__ import annotations

from typing import Any

from gwmd.ir import Block, Document, Heading, Link, Paragraph, Table, Text
from gwmd.parsers._grids import collapse_empty


def parse(ss: dict) -> Document:
    blocks: list[Block] = []
    title = ss.get("properties", {}).get("title")
    for sheet in ss.get("sheets", []):
        st = sheet.get("properties", {}).get("title", "")
        blocks.append(Heading(level=2, inlines=[Text(value=st)]))
        grid = collapse_empty(_raw_cells(sheet), is_empty=_empty)
        if not grid:
            continue
        ncols = len(grid[0])
        header = [_cell(grid[0][c]) for c in range(ncols)]
        rows = [
            [_cell(r[c] if c < len(r) else None) for c in range(ncols)]
            for r in grid[1:]
        ]
        blocks.append(Table(header=header, rows=rows))
    return Document(title=title, source="gsheet", blocks=blocks)


def _raw_cells(sheet: dict) -> list[list[dict | None]]:
    """シートの CellData を 2 次元配列（dict or None）で返す。"""
    data = sheet.get("data", [])
    if not data:
        return []
    rows = []
    for row in data[0].get("rowData", []):
        rows.append(list(row.get("values", [])))
    return rows


def _empty(cell: Any) -> bool:
    return cell is None or not cell.get("formattedValue")


def _cell(cell: dict | None) -> list[Block]:
    text = "" if cell is None else cell.get("formattedValue", "")
    if not text:
        return [Paragraph(inlines=[])]
    href = None if cell is None else cell.get("hyperlink")
    inlines = [Text(value=text)]
    if href:
        inlines = [Link(href=href, inlines=inlines)]
    return [Paragraph(inlines=inlines)]
