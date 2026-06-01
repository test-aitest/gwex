"""IR → typed JSON。`kind` 付き tagged union をそのまま出力する。

pure 層: I/O しない。
"""

from __future__ import annotations

from gwmd.ir import Document


def render(doc: Document, *, indent: int | None = 2) -> str:
    """Document を typed JSON 文字列に変換する。None フィールドは省略。"""
    return doc.model_dump_json(indent=indent, exclude_none=True)
