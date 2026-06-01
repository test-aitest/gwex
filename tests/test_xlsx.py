"""xlsx parser の単体テスト（メモリ上で小さな .xlsx を生成して検証）。"""

from __future__ import annotations

from io import BytesIO

import openpyxl

from gwex.ir import Heading, Table
from gwex.parsers import xlsx
from gwex.serializers import markdown as md


def _xlsx_bytes(rows: list[list], title: str = "Sheet1") -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = title
    for r in rows:
        ws.append(r)
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_basic_table_with_sheet_heading():
    data = _xlsx_bytes([["項目", "値"], ["A", 1], ["B", 2]], title="シート1")
    doc = xlsx.parse(data)
    assert doc.source == "xlsx"
    assert isinstance(doc.blocks[0], Heading)
    assert doc.blocks[0].inlines[0].value == "シート1"
    table = doc.blocks[1]
    assert isinstance(table, Table)
    out = md.render(doc)
    assert "## シート1" in out
    assert "| 項目 | 値 |" in out
    assert "| A | 1 |" in out
    assert "| B | 2 |" in out


def test_empty_columns_and_rows_are_collapsed():
    # 中央に空列、間に空行を挟む（レイアウト由来のノイズ）
    rows = [
        ["名前", None, None, "得点"],
        [None, None, None, None],
        ["太郎", None, None, 80],
    ]
    doc = _xlsx_bytes(rows)
    out = md.render(xlsx.parse(doc))
    # 全体で空の列(2,3列目)と空行は畳まれ、2列に圧縮される
    assert "| 名前 | 得点 |" in out
    assert "| 太郎 | 80 |" in out


def test_float_integer_formatting():
    doc = xlsx.parse(_xlsx_bytes([["x"], [3.0]]))
    out = md.render(doc)
    assert "| 3 |" in out  # 3.0 ではなく 3
