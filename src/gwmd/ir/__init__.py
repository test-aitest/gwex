"""gwmd の正規化 Document IR。

全ソース（Google Docs/Sheets/Slides, Microsoft docx/xlsx/pptx）はこの IR に落ち、
Serializer は IR からのみ Markdown / typed JSON を生成する。

この層は **pure**: I/O も fetcher も import しない（import-linter で強制）。
"""

from gwmd.ir.model import (
    Block,
    CodeBlock,
    CodeSpan,
    Divider,
    Document,
    Emph,
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
    FetchedRef,
    ImageRef,
)

__all__ = [
    "Document",
    "Block",
    "Heading",
    "Paragraph",
    "ListBlock",
    "Table",
    "CodeBlock",
    "Quote",
    "Image",
    "Note",
    "Divider",
    "Inline",
    "Text",
    "Strong",
    "Emph",
    "CodeSpan",
    "Link",
    "ImageRef",
    "UrlRef",
    "FetchedRef",
]
