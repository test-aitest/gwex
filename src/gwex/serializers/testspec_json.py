"""TestSpec（階層）→ typed JSON。親（画面/中項目）を1回だけ持つネスト構造。pure。"""

from __future__ import annotations

from gwex.domains.model import TestSpec


def render(spec: TestSpec, *, indent: int | None = 2) -> str:
    # source_row 等の None フィールドは省く（生成側が書く JSON とも揃う）。
    return spec.model_dump_json(indent=indent, exclude_none=True)
