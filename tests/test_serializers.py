"""背骨（IR + Serializer）の単体テスト。手書き IR から MD/JSON が出ることを確認。"""

from __future__ import annotations

import json

from gwex.ir import (
    CodeBlock,
    Divider,
    Document,
    Emph,
    Heading,
    Image,
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
from gwex.serializers import markdown as md
from gwex.serializers import json as jsonser


def _doc(*blocks) -> Document:
    return Document(source="gdocs", title="t", blocks=list(blocks))


def test_heading_and_paragraph():
    out = md.render(_doc(Heading(level=2, inlines=[Text(value="概要")])))
    assert out == "## 概要\n"


def test_inline_styles():
    p = Paragraph(
        inlines=[
            Text(value="これは "),
            Strong(inlines=[Text(value="太字")]),
            Text(value=" と "),
            Emph(inlines=[Text(value="斜体")]),
            Text(value=" と "),
            Link(href="https://example.com", inlines=[Text(value="リンク")]),
        ]
    )
    out = md.render(_doc(p))
    assert "**太字**" in out
    assert "*斜体*" in out
    assert "[リンク](https://example.com)" in out


def test_table_gfm():
    cell = lambda s: [Paragraph(inlines=[Text(value=s)])]  # noqa: E731
    t = Table(
        header=[cell("項目"), cell("値")],
        rows=[[cell("A"), cell("1")], [cell("B"), cell("2")]],
    )
    out = md.render(_doc(t))
    lines = out.strip().split("\n")
    assert lines[0] == "| 項目 | 値 |"
    assert lines[1] == "| --- | --- |"
    assert lines[2] == "| A | 1 |"
    assert lines[3] == "| B | 2 |"


def test_table_cell_pipe_escaped():
    cell = lambda s: [Paragraph(inlines=[Text(value=s)])]  # noqa: E731
    t = Table(header=None, rows=[[cell("a|b")]])
    out = md.render(_doc(t))
    assert "a\\|b" in out


def test_nested_list():
    item = lambda s, *sub: [Paragraph(inlines=[Text(value=s)]), *sub]  # noqa: E731
    inner = ListBlock(ordered=False, items=[item("子1"), item("子2")])
    lst = ListBlock(ordered=False, items=[item("親", inner), item("親2")])
    out = md.render(_doc(lst))
    assert "- 親" in out
    assert "  - 子1" in out
    assert "  - 子2" in out
    assert "- 親2" in out


def test_ordered_list():
    item = lambda s: [Paragraph(inlines=[Text(value=s)])]  # noqa: E731
    lst = ListBlock(ordered=True, items=[item("一"), item("二")])
    out = md.render(_doc(lst))
    assert "1. 一" in out
    assert "2. 二" in out


def test_code_quote_note_image_divider():
    doc = _doc(
        CodeBlock(lang="python", text="print(1)"),
        Quote(blocks=[Paragraph(inlines=[Text(value="引用文")])]),
        Note(blocks=[Paragraph(inlines=[Text(value="ノート文")])]),
        Image(src=UrlRef(url="https://img/x.png"), alt="図"),
        Divider(),
    )
    out = md.render(doc)
    assert "```python\nprint(1)\n```" in out
    assert "> 引用文" in out
    assert "> **Note:**" in out
    assert "![図](https://img/x.png)" in out
    assert "---" in out


def test_json_roundtrip():
    doc = _doc(
        Heading(level=1, inlines=[Text(value="概要")]),
        Paragraph(inlines=[Link(href="https://x", inlines=[Text(value="こちら")])]),
    )
    s = jsonser.render(doc)
    data = json.loads(s)
    assert data["source"] == "gdocs"
    assert data["blocks"][0]["kind"] == "heading"
    assert data["blocks"][1]["inlines"][0]["kind"] == "link"
    # typed JSON はそのまま IR に戻せる
    back = Document.model_validate_json(s)
    assert md.render(back) == md.render(doc)
