"""TestSpec（階層）→ typed JSON。親（画面/中項目）を1回だけ持つネスト構造。pure。"""

from __future__ import annotations

from gwmd.domains.model import TestSpec


def render(spec: TestSpec, *, indent: int | None = 2) -> str:
    return spec.model_dump_json(indent=indent)
