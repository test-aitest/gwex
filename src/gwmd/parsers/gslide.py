"""Google Slides (`presentations.get` の JSON) → IR。

pure 層。各スライドを `Divider` + `Heading(1)` "Slide N" に区切り、図形テキスト・
表・画像・発表者ノートを写像する。pageElements は読み順保証が無いため
transform（y→x）でソートしてから直列化する。
"""

from __future__ import annotations

from typing import Any

from gwmd.ir import (
    Block,
    Divider,
    Document,
    Emph,
    Heading,
    Image,
    Inline,
    Link,
    Note,
    Paragraph,
    Strong,
    Table,
    Text,
    UrlRef,
)


def parse(pres: dict) -> Document:
    blocks: list[Block] = []
    for idx, slide in enumerate(pres.get("slides", []), start=1):
        blocks.append(Divider())
        blocks.append(Heading(level=1, inlines=[Text(value=f"Slide {idx}")]))
        for pe in sorted(slide.get("pageElements", []), key=_yx):
            blocks.extend(_element(pe))
        notes = _notes(slide)
        if notes:
            blocks.append(Note(blocks=notes))
    return Document(title=pres.get("title"), source="gslide", blocks=blocks)


def _yx(pe: dict) -> tuple[float, float]:
    t = pe.get("transform", {})
    return (t.get("translateY", 0.0), t.get("translateX", 0.0))


def _element(pe: dict) -> list[Block]:
    if "shape" in pe:
        text = pe["shape"].get("text")
        return _text_paragraphs(text) if text else []
    if "table" in pe:
        return [_table(pe["table"])]
    if "image" in pe:
        url = pe["image"].get("contentUrl")
        if url:
            desc = pe.get("description") or pe.get("title")
            return [Image(src=UrlRef(url=url), alt=desc)]
    return []


def _table(table: dict) -> Table:
    rows: list[list[list[Block]]] = []
    for row in table.get("tableRows", []):
        cells = []
        for cell in row.get("tableCells", []):
            cells.append(_text_paragraphs(cell.get("text")))
        rows.append(cells)
    if not rows:
        return Table(header=None, rows=[])
    return Table(header=rows[0], rows=rows[1:])


def _text_paragraphs(text: dict | None) -> list[Block]:
    """shape.text / cell.text の textElements を Paragraph 列に。

    textElements は flat で、paragraphMarker が段落の開始を示す。
    """
    if not text:
        return []
    out: list[Block] = []
    current: list[Inline] = []

    def flush() -> None:
        if current:
            out.append(Paragraph(inlines=list(current)))
            current.clear()

    for te in text.get("textElements", []):
        if "paragraphMarker" in te:
            flush()
        elif "textRun" in te:
            run = _run(te["textRun"])
            if run is not None:
                current.append(run)
    flush()
    return out


def _run(tr: dict) -> Inline | None:
    content = tr.get("content", "").rstrip("\n")
    if content == "":
        return None
    style = tr.get("style", {})
    node: Inline = Text(value=content)
    link = style.get("link", {})
    if link.get("url"):
        node = Link(href=link["url"], inlines=[node])
    if style.get("italic"):
        node = Emph(inlines=[node])
    if style.get("bold"):
        node = Strong(inlines=[node])
    return node


def _notes(slide: dict) -> list[Block]:
    notes_page = slide.get("slideProperties", {}).get("notesPage", {})
    out: list[Block] = []
    for pe in notes_page.get("pageElements", []):
        if "shape" in pe and pe["shape"].get("text"):
            out.extend(_text_paragraphs(pe["shape"]["text"]))
    return out
