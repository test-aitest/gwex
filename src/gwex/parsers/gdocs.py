"""Google Docs (`documents.get` の JSON) → IR。

pure 層: 与えられた dict を走査するだけ（I/O しない）。
`body.content[]`（StructuralElement）を再帰的に IR へ写像する。
"""

from __future__ import annotations

from typing import Any

from gwex.ir import (
    Block,
    Document,
    Emph,
    Heading,
    Image,
    Inline,
    Link,
    ListBlock,
    Paragraph,
    Strong,
    Table,
    Text,
    UrlRef,
)

_HEADING = {
    "TITLE": 1,
    "SUBTITLE": 2,
    "HEADING_1": 1,
    "HEADING_2": 2,
    "HEADING_3": 3,
    "HEADING_4": 4,
    "HEADING_5": 5,
    "HEADING_6": 6,
}


def parse(doc: dict) -> Document:
    body = doc.get("body", {}).get("content", [])
    blocks = _walk(body, doc)
    return Document(title=doc.get("title"), source="gdocs", blocks=blocks)


def _walk(content: list[dict], doc: dict) -> list[Block]:
    """StructuralElement 列を走査。連続する箇条書き段落は 1 つの ListBlock にまとめる。"""
    blocks: list[Block] = []
    i, n = 0, len(content)
    while i < n:
        para = content[i].get("paragraph")
        if para is not None and "bullet" in para:
            group = []
            while i < n:
                p = content[i].get("paragraph")
                if p is None or "bullet" not in p:
                    break
                group.append(p)
                i += 1
            blocks.append(_build_list(group, doc))
            continue
        blocks.extend(_structural(content[i], doc))
        i += 1
    return blocks


def _structural(el: dict, doc: dict) -> list[Block]:
    if "paragraph" in el:
        para = el["paragraph"]
        style = para.get("paragraphStyle", {}).get("namedStyleType", "NORMAL_TEXT")
        inlines, images = _paragraph_inlines(para, doc)
        out: list[Block] = []
        if inlines:
            if style in _HEADING:
                out.append(Heading(level=_HEADING[style], inlines=inlines))
            else:
                out.append(Paragraph(inlines=inlines))
        out.extend(images)
        return out
    if "table" in el:
        return [_table(el["table"], doc)]
    # sectionBreak / tableOfContents 等は構造的意味を持たないので無視
    return []


def _build_list(paras: list[dict], doc: dict) -> ListBlock:
    """箇条書き段落群を nestingLevel に従ってネストした ListBlock に組む。"""
    root = ListBlock(ordered=_is_ordered(paras[0], doc), items=[])
    stack: list[tuple[int, ListBlock]] = [(0, root)]
    for p in paras:
        level = p["bullet"].get("nestingLevel", 0)
        inlines, images = _paragraph_inlines(p, doc)
        item: list[Block] = [Paragraph(inlines=inlines), *images]
        while len(stack) > 1 and stack[-1][0] > level:
            stack.pop()
        if stack[-1][0] < level:
            nested = ListBlock(ordered=_is_ordered(p, doc), items=[])
            parent = stack[-1][1]
            if parent.items:
                parent.items[-1].append(nested)
            else:
                parent.items.append([nested])
            stack.append((level, nested))
        stack[-1][1].items.append(item)
    return root


def _is_ordered(para: dict, doc: dict) -> bool:
    bullet = para.get("bullet", {})
    list_id = bullet.get("listId")
    nesting = bullet.get("nestingLevel", 0)
    try:
        nl = doc["lists"][list_id]["listProperties"]["nestingLevels"][nesting]
    except (KeyError, IndexError, TypeError):
        return False
    gt = nl.get("glyphType")
    return gt is not None and gt != "GLYPH_TYPE_UNSPECIFIED"


def _table(table: dict, doc: dict) -> Table:
    cells_by_row: list[list[list[Block]]] = []
    for row in table.get("tableRows", []):
        cells_by_row.append(
            [_walk(cell.get("content", []), doc) for cell in row.get("tableCells", [])]
        )
    if not cells_by_row:
        return Table(header=None, rows=[])
    return Table(header=cells_by_row[0], rows=cells_by_row[1:])


def _paragraph_inlines(para: dict, doc: dict) -> tuple[list[Inline], list[Image]]:
    inlines: list[Inline] = []
    images: list[Image] = []
    for e in para.get("elements", []):
        if "textRun" in e:
            run = _run(e["textRun"])
            if run is not None:
                inlines.append(run)
        elif "inlineObjectElement" in e:
            img = _inline_image(e["inlineObjectElement"], doc)
            if img is not None:
                images.append(img)
    return inlines, images


def _run(tr: dict) -> Inline | None:
    content = tr.get("content", "").rstrip("\n")
    if content == "":
        return None
    style = tr.get("textStyle", {})
    node: Inline = Text(value=content)
    link = style.get("link", {})
    if link.get("url"):
        node = Link(href=link["url"], inlines=[node])
    if style.get("italic"):
        node = Emph(inlines=[node])
    if style.get("bold"):
        node = Strong(inlines=[node])
    return node


def _inline_image(elem: dict, doc: dict) -> Image | None:
    obj_id = elem.get("inlineObjectId")
    try:
        embedded = doc["inlineObjects"][obj_id]["inlineObjectProperties"]["embeddedObject"]
    except (KeyError, TypeError):
        return None
    uri = embedded.get("imageProperties", {}).get("contentUri")
    if not uri:
        return None
    alt = embedded.get("title") or embedded.get("description")
    return Image(src=UrlRef(url=uri), alt=alt)
