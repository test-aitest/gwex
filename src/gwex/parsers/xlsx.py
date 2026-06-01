"""Microsoft Excel (.xlsx) → IR。

pure 層: 与えられた**バイト列**を BytesIO で解析するだけ（パスを開かない）。
方針: `formattedValue` 相当（data_only のキャッシュ値）、1 行目をヘッダ扱い、
結合セルは左上に値・他は空欄（openpyxl が既にそう返すため自然に一致）。
"""

from __future__ import annotations

import datetime as _dt
from io import BytesIO
from typing import Any

import openpyxl

from gwex.ir import Document, Heading, Paragraph, Table, Text
from gwex.parsers._grids import collapse_empty


def parse(data: bytes) -> Document:
    wb = openpyxl.load_workbook(BytesIO(data), data_only=True)
    blocks: list[Any] = []
    for ws in wb.worksheets:
        blocks.append(Heading(level=2, inlines=[Text(value=ws.title)]))
        grid = _used_grid(ws)
        if not grid:
            continue
        ncols = len(grid[0])
        header = [_cell(grid[0][c]) for c in range(ncols)]
        rows = [
            [_cell(row[c] if c < len(row) else None) for c in range(ncols)]
            for row in grid[1:]
        ]
        blocks.append(Table(header=header, rows=rows))
    return Document(title=None, source="xlsx", blocks=blocks)


def _used_grid(ws) -> list[list[Any]]:
    """実データのある行・列だけを残した 2 次元配列を返す。

    結合セル由来の空列・行間スペーサ行はレイアウト情報にすぎず LLM には
    ノイズなので、**全体で空の列・行は畳む**（意味を持つ構造だけ残す原則）。
    列を落としても全行に等しく作用するためヘッダ/データの整合は保たれる。
    """
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    return collapse_empty(rows, is_empty=lambda v: v is None)


def _cell(value: Any) -> list:
    """セル 1 つを Block 列（[Paragraph]）に。空セルは空 Paragraph。

    NOTE(v1): ハイパーリンクは未対応（空列畳み込みで座標がずれるため）。後続で
    元座標を保持して対応する。
    """
    text = _fmt(value)
    if not text:
        return [Paragraph(inlines=[])]
    return [Paragraph(inlines=[Text(value=text)])]


def _fmt(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, (_dt.datetime, _dt.date, _dt.time)):
        return value.isoformat()
    return str(value)
