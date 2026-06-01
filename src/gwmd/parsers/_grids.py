"""表系パーサ共有: 全体で空の列・行を畳むユーティリティ（pure）。

結合セル由来の空列や行間スペーサ行はレイアウト情報にすぎず LLM にはノイズ。
列を落としても全行に等しく作用するためヘッダ/データの整合は保たれる。
"""

from __future__ import annotations

from typing import Any, Callable


def collapse_empty(rows: list[list[Any]], is_empty: Callable[[Any], bool]) -> list[list[Any]]:
    if not rows:
        return []
    width = max((len(r) for r in rows), default=0)
    keep = [
        c for c in range(width)
        if any(c < len(r) and not is_empty(r[c]) for r in rows)
    ]
    return [
        [r[c] if c < len(r) else None for c in keep]
        for r in rows
        if any(not is_empty(v) for v in r)
    ]
