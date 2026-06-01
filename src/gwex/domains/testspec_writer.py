"""TestSpec（階層）→ セル割当 [(cell, value)] への展開。pure な配置計算。

抽出の逆変換。結合を作らずに「親（画面/中項目/小項目）は先頭行に1回だけ」書き、
継続行は空欄にする（抽出側がカテゴリを前方補完して復元する）。test_no/確認内容/
実施手順は各ケース行に書く（execution_steps は改行結合で 1 セル）。
"""

from __future__ import annotations

from typing import Optional

from openpyxl.utils import get_column_letter

from gwex.domains.model import DEFAULT_MAPPING, MappingConfig, TestSpec

_ROLES = (
    "screen_name",
    "medium_category",
    "small_category",
    "test_no",
    "verification_content",
    "execution_steps",
)


def assignments(
    spec: TestSpec,
    mapping: Optional[MappingConfig] = None,
    start_row: Optional[int] = None,
) -> list[tuple[str, object]]:
    cfg = mapping or DEFAULT_MAPPING
    col = {}
    for role in _ROLES:
        idx = cfg.column_index(role)
        if idx is not None:
            col[role] = get_column_letter(idx + 1)

    out: list[tuple[str, object]] = []
    row = start_row if start_row is not None else cfg.data_start_row

    def put(role: str, value: object) -> None:
        if role in col and value not in (None, ""):
            out.append((f"{col[role]}{row}", value))

    for screen in spec.screens:
        first_screen = True
        for group in screen.groups:
            first_group = True
            for case in group.cases:
                if first_screen:
                    put("screen_name", screen.screen_name)
                    first_screen = False
                if first_group:
                    put("medium_category", group.medium_category)
                    put("small_category", group.small_category)
                    first_group = False
                put("test_no", case.test_no)
                put("verification_content", case.verification_content)
                put("execution_steps", "\n".join(case.execution_steps))
                row += 1
    return out
