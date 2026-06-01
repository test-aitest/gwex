"""表系パーサ共有: 全体で空の列・行を畳むユーティリティ（pure）。

結合セル由来の空列や行間スペーサ行はレイアウト情報にすぎず LLM にはノイズ。
列を落としても全行に等しく作用するためヘッダ/データの整合は保たれる。
"""

from __future__ import annotations

from typing import Any, Callable


def collapse_empty(rows: list[list[Any]], is_empty: Callable[[Any], bool]) -> list[list[Any]]:
    """全体で空の列・行を畳む。

    旧実装は「列ごとに全行を再走査」で O(列 × 行) の入れ子だった。ここでは
    全セルを 1 度だけ走査して使用列・使用行を記録する単一パスに置き換える。
    """
    if not rows:
        return []
    width = max((len(r) for r in rows), default=0)
    col_used = [False] * width
    row_used = [False] * len(rows)
    for ri, r in enumerate(rows):
        for c, v in enumerate(r):
            if not is_empty(v):
                col_used[c] = True
                row_used[ri] = True
    keep = [c for c in range(width) if col_used[c]]
    return [
        [r[c] if c < len(r) else None for c in keep]
        for ri, r in enumerate(rows)
        if row_used[ri]
    ]


def merge_fill(ws) -> list[list]:
    """結合セルを前方補完した絶対座標グリッド（0始まり）を返す。

    汎用パーサの collapse_empty とは逆方向: 列を畳まず、結合範囲の左上値を
    範囲内の全セルへ展開する。テスト仕様抽出は列レター（絶対位置）でロールを
    解決するため、列を保持したまま結合だけ埋める。

    グリッド構築は per-cell の `ws.cell()`（openpyxl で高コスト）をやめ、
    `iter_rows(values_only=True)` の単一走査にする。結合の補完は各範囲の
    面積分のみ（総和は最大でもセル数）で、全体は O(セル数)。
    """
    grid = [list(row) for row in ws.iter_rows(values_only=True)]
    for rng in ws.merged_cells.ranges:
        top = grid[rng.min_row - 1][rng.min_col - 1]
        for r in range(rng.min_row - 1, rng.max_row):
            row = grid[r]
            for c in range(rng.min_col - 1, rng.max_col):
                row[c] = top
    return grid
