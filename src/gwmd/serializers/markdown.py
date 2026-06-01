"""IR → Markdown（GitHub Flavored）。全ソースがこの 1 実装を共有する。

pure 層: I/O しない。`Document` を受け取り文字列を返すだけ。
"""

from __future__ import annotations

from gwmd.ir import (
    Block,
    CodeBlock,
    CodeSpan,
    Divider,
    Document,
    Emph,
    FetchedRef,
    Heading,
    Image,
    Inline,
    Link,
    ListBlock,
    Note,
    Paragraph,
    Quote,
    Strong,
    Table,
    Text,
    UrlRef,
)


def render(doc: Document) -> str:
    """Document を Markdown 文字列に変換する。"""
    body = _render_blocks(doc.blocks)
    return body + "\n" if body else ""


def _render_blocks(blocks: list[Block]) -> str:
    """ブロック列を空行区切りで連結する。"""
    parts = [p for b in blocks if (p := _render_block(b))]
    return "\n\n".join(parts)


def _render_block(b: Block) -> str:
    if isinstance(b, Heading):
        return f"{'#' * b.level} {_inlines(b.inlines)}"
    if isinstance(b, Paragraph):
        return _inlines(b.inlines)
    if isinstance(b, ListBlock):
        return _render_list(b)
    if isinstance(b, Table):
        return _render_table(b)
    if isinstance(b, CodeBlock):
        return f"```{b.lang or ''}\n{b.text}\n```"
    if isinstance(b, Quote):
        return _prefix_lines(_render_blocks(b.blocks), "> ")
    if isinstance(b, Note):
        inner = _prefix_lines(_render_blocks(b.blocks), "> ")
        return f"> **Note:**\n{inner}" if inner else "> **Note:**"
    if isinstance(b, Image):
        return f"![{b.alt or ''}]({_image_src(b)})"
    if isinstance(b, Divider):
        return "---"
    raise TypeError(f"unknown block: {type(b).__name__}")


# --- list -----------------------------------------------------------------


def _render_list(lst: ListBlock, depth: int = 0) -> str:
    lines: list[str] = []
    indent = "  " * depth
    for i, item in enumerate(item for item in lst.items):
        marker = f"{i + 1}." if lst.ordered else "-"
        lines.append(_render_list_item(item, marker, indent, depth))
    return "\n".join(lines)


def _render_list_item(item: list[Block], marker: str, indent: str, depth: int) -> str:
    """1 項目（Block 列）を描画。先頭にマーカー、ネストしたリストは深さで字下げ。"""
    out: list[str] = []
    cont_indent = indent + " " * (len(marker) + 1)
    for j, block in enumerate(item):
        if isinstance(block, ListBlock):
            out.append(_render_list(block, depth + 1))
            continue
        text = _render_block(block)
        if j == 0:
            out.append(f"{indent}{marker} {text}")
        else:
            out.append(_prefix_lines(text, cont_indent))
    return "\n".join(out)


# --- table ----------------------------------------------------------------


def _render_table(t: Table) -> str:
    ncols = _column_count(t)
    if ncols == 0:
        return ""
    header = t.header if t.header is not None else [[] for _ in range(ncols)]
    lines = [_table_row(header, ncols), _table_divider(ncols)]
    lines += [_table_row(row, ncols) for row in t.rows]
    return "\n".join(lines)


def _column_count(t: Table) -> int:
    counts = [len(t.header)] if t.header is not None else []
    counts += [len(r) for r in t.rows]
    return max(counts, default=0)


def _table_row(cells: list[list[Block]], ncols: int) -> str:
    rendered = [_render_cell(c) for c in cells]
    rendered += [""] * (ncols - len(rendered))
    return "| " + " | ".join(rendered) + " |"


def _table_divider(ncols: int) -> str:
    return "| " + " | ".join(["---"] * ncols) + " |"


def _render_cell(cell: list[Block]) -> str:
    """セル（Block 列）を 1 行に折りたたむ。GFM のセルは改行不可なので <br> 化。"""
    text = _render_blocks(cell)
    return text.replace("\\", "\\\\").replace("|", "\\|").replace("\n", "<br>")


# --- inline ---------------------------------------------------------------


def _inlines(inlines: list[Inline]) -> str:
    return "".join(_inline(i) for i in _coalesce(inlines))


def _coalesce(inlines: list[Inline]) -> list[Inline]:
    """隣接する同種の強調（Strong/Emph）を結合する。

    複数ソースで「同じ書式が別 run に分かれる」ことが多く、素直に出すと
    `**A****B**` のような空の強調が生じるため、`**AB**` に畳む。
    """
    out: list[Inline] = []
    for node in inlines:
        prev = out[-1] if out else None
        if isinstance(node, Strong) and isinstance(prev, Strong):
            out[-1] = Strong(inlines=[*prev.inlines, *node.inlines])
        elif isinstance(node, Emph) and isinstance(prev, Emph):
            out[-1] = Emph(inlines=[*prev.inlines, *node.inlines])
        else:
            out.append(node)
    return out


def _inline(i: Inline) -> str:
    if isinstance(i, Text):
        return _normalize(i.value)
    if isinstance(i, Strong):
        return f"**{_inlines(i.inlines)}**"
    if isinstance(i, Emph):
        return f"*{_inlines(i.inlines)}*"
    if isinstance(i, CodeSpan):
        return f"`{i.text}`"
    if isinstance(i, Link):
        return f"[{_inlines(i.inlines)}]({i.href})"
    raise TypeError(f"unknown inline: {type(i).__name__}")


# --- helpers --------------------------------------------------------------


def _image_src(img: Image) -> str:
    src = img.src
    if isinstance(src, UrlRef):
        return src.url
    if isinstance(src, FetchedRef):
        return f"attachment:{src.id}"
    raise TypeError(f"unknown image ref: {type(src).__name__}")


def _prefix_lines(text: str, prefix: str) -> str:
    return "\n".join(prefix + line for line in text.split("\n"))


def _normalize(value: str) -> str:
    """制御文字を正規化する。

    Google Docs/Slides は段落内のソフト改行に垂直タブ(\\v)を使う。不可視のまま
    出すと `**A****B**` のような壊れた強調になるため改行に変換する。BOM も除去。
    """
    return (
        value.replace("\v", "\n")
        .replace("\f", "\n")
        .replace("\r", "")
        .replace("﻿", "")
    )
