"""正規化 Document IR の型定義（pydantic v2）。

設計原則: 「意味が変わる構造は保存、見た目だけの情報は捨てる」。
この Union に無い視覚情報（フォント・色・座標・ページ余白）は IR に存在できない。

`kind` フィールドを discriminator にした tagged union で、JSON へそのまま
ラウンドトリップできる（typed JSON 出力・MCP スキーマにそのまま使える）。
"""

from __future__ import annotations

from typing import Annotated, Literal, Optional, Union

from pydantic import BaseModel, Field

# --------------------------------------------------------------------------
# Inline（段落内のテキスト断片）
# --------------------------------------------------------------------------


class Text(BaseModel):
    kind: Literal["text"] = "text"
    value: str


class Strong(BaseModel):
    kind: Literal["strong"] = "strong"
    inlines: list[Inline]


class Emph(BaseModel):
    kind: Literal["emph"] = "emph"
    inlines: list[Inline]


class CodeSpan(BaseModel):
    kind: Literal["code_span"] = "code_span"
    text: str


class Link(BaseModel):
    kind: Literal["link"] = "link"
    href: str
    inlines: list[Inline]


Inline = Annotated[
    Union[Text, Strong, Emph, CodeSpan, Link],
    Field(discriminator="kind"),
]


# --------------------------------------------------------------------------
# 画像参照（既定: 取得しない＝URL を参照のまま残す）
# --------------------------------------------------------------------------


class UrlRef(BaseModel):
    kind: Literal["url"] = "url"
    url: str


class FetchedRef(BaseModel):
    kind: Literal["fetched"] = "fetched"
    id: int
    mime: str


ImageRef = Annotated[Union[UrlRef, FetchedRef], Field(discriminator="kind")]


class MediaItem(BaseModel):
    """取得済みバイナリ（画像など）の実体。

    `Image(src=FetchedRef(id=...))` から id 参照される。bytes は base64 文字列で
    持ち、sandbox の別プロセス境界（Document の JSON 化）を越えて運べるようにする。
    """

    id: int
    mime: str
    data_b64: str


# --------------------------------------------------------------------------
# Block（文書構造）
# --------------------------------------------------------------------------
#
# 表・リストのセル/項目は「Block 列」を持つ。これにより Google Docs の
# 「セル内も StructuralElement（＝再帰構造）」をそのまま表現できる。
#   - ListBlock.items : list[項目]  各項目 = list[Block]
#   - Table.rows      : list[行]    各行 = list[セル], 各セル = list[Block]


class Heading(BaseModel):
    kind: Literal["heading"] = "heading"
    level: int = Field(ge=1, le=6)
    inlines: list[Inline]


class Paragraph(BaseModel):
    kind: Literal["paragraph"] = "paragraph"
    inlines: list[Inline]


class ListBlock(BaseModel):
    kind: Literal["list"] = "list"
    ordered: bool
    items: list[list[Block]]


class Table(BaseModel):
    kind: Literal["table"] = "table"
    header: Optional[list[list[Block]]] = None
    rows: list[list[list[Block]]]


class CodeBlock(BaseModel):
    kind: Literal["code"] = "code"
    lang: Optional[str] = None
    text: str


class Quote(BaseModel):
    kind: Literal["quote"] = "quote"
    blocks: list[Block]


class Image(BaseModel):
    kind: Literal["image"] = "image"
    src: ImageRef
    alt: Optional[str] = None


class Note(BaseModel):
    """発表者ノート等、本文に対する補助ブロック。"""

    kind: Literal["note"] = "note"
    blocks: list[Block]


class Divider(BaseModel):
    """水平線。スライド境界などに用いる。"""

    kind: Literal["divider"] = "divider"


Block = Annotated[
    Union[Heading, Paragraph, ListBlock, Table, CodeBlock, Quote, Image, Note, Divider],
    Field(discriminator="kind"),
]


# --------------------------------------------------------------------------
# Document（ルート）
# --------------------------------------------------------------------------

SourceKind = Literal["gdocs", "gsheet", "gslide", "docx", "xlsx", "pptx"]


class Document(BaseModel):
    title: Optional[str] = None
    source: SourceKind
    blocks: list[Block]
    # 取得済みバイナリ（画像等）の実体。Image(FetchedRef(id)) から id 参照される。
    media: list[MediaItem] = Field(default_factory=list)


# 前方参照（list[Inline] / list[Block] 等）を解決する。
for _model in (
    Strong,
    Emph,
    Link,
    Heading,
    Paragraph,
    ListBlock,
    Table,
    Quote,
    Image,
    Note,
    Document,
):
    _model.model_rebuild()
