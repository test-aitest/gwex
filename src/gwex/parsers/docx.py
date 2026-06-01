"""Microsoft Word (.docx) → IR。

pure 層: 与えられたバイト列を BytesIO で解析するだけ。本文の出現順を保つため
body 直下の要素（w:p / w:tbl）を順に走査する。
"""

from __future__ import annotations

from io import BytesIO

from docx import Document as _Docx
from docx.oxml.ns import qn
from docx.table import Table as _DocxTable
from docx.text.paragraph import Paragraph as _DocxParagraph

from gwex.ir import (
    Block,
    Document,
    Emph,
    Heading,
    Inline,
    ListBlock,
    Paragraph,
    Strong,
    Table,
    Text,
)

_MAX_HEADING = 6


def parse(data: bytes) -> Document:
    docx = _Docx(BytesIO(data))
    items = list(_iter_block_items(docx))
    blocks = _walk(items)
    title = docx.core_properties.title or None
    return Document(title=title, source="docx", blocks=blocks)


def _iter_block_items(docx):
    """body 直下の段落・表を出現順に yield する。"""
    body = docx.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield _DocxParagraph(child, docx)
        elif child.tag == qn("w:tbl"):
            yield _DocxTable(child, docx)


def _walk(items: list) -> list[Block]:
    blocks: list[Block] = []
    i, n = 0, len(items)
    while i < n:
        it = items[i]
        if isinstance(it, _DocxParagraph) and _list_kind(it) is not None:
            ordered = _list_kind(it) == "ordered"
            group = []
            while i < n and isinstance(items[i], _DocxParagraph) and _list_kind(items[i]) is not None:
                group.append(items[i])
                i += 1
            blocks.append(
                ListBlock(ordered=ordered, items=[[Paragraph(inlines=_inlines(p))] for p in group])
            )
            continue
        if isinstance(it, _DocxParagraph):
            block = _paragraph(it)
            if block is not None:
                blocks.append(block)
        elif isinstance(it, _DocxTable):
            blocks.append(_table(it))
        i += 1
    return blocks


def _list_kind(p: _DocxParagraph) -> str | None:
    name = (p.style.name or "").lower() if p.style else ""
    if name.startswith("list number"):
        return "ordered"
    if name.startswith("list bullet") or name.startswith("list paragraph"):
        return "unordered"
    return None


def _paragraph(p: _DocxParagraph) -> Block | None:
    inlines = _inlines(p)
    if not inlines:
        return None
    level = _heading_level(p)
    if level is not None:
        return Heading(level=level, inlines=inlines)
    return Paragraph(inlines=inlines)


def _heading_level(p: _DocxParagraph) -> int | None:
    name = (p.style.name or "") if p.style else ""
    if name == "Title":
        return 1
    if name.startswith("Heading "):
        try:
            return min(int(name.split(" ", 1)[1]), _MAX_HEADING)
        except ValueError:
            return None
    return None


def _inlines(p: _DocxParagraph) -> list[Inline]:
    out: list[Inline] = []
    for run in p.runs:
        if run.text == "":
            continue
        node: Inline = Text(value=run.text)
        if run.italic:
            node = Emph(inlines=[node])
        if run.bold:
            node = Strong(inlines=[node])
        out.append(node)
    return out


def _table(t: _DocxTable) -> Table:
    rows: list[list[list[Block]]] = []
    for row in t.rows:
        cells = []
        for cell in row.cells:
            cell_blocks: list[Block] = [
                b for p in cell.paragraphs if (b := _paragraph(p)) is not None
            ]
            cells.append(cell_blocks or [Paragraph(inlines=[])])
        rows.append(cells)
    if not rows:
        return Table(header=None, rows=[])
    return Table(header=rows[0], rows=rows[1:])
