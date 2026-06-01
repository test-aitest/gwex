"""結合テスト項目シート（xlsx bytes）→ TestSpec（階層）抽出。pure 層。

結合セルを merge_fill で前方補完し、列マッピングで各ロールを解決して
画面名→(中項目,小項目)→テストケース の階層に組む。
"""

from __future__ import annotations

import re
from io import BytesIO
from typing import Any, Optional, Union

import openpyxl

from gwex.domains.model import (
    DEFAULT_MAPPING,
    Case,
    Group,
    MappingConfig,
    Screen,
    TestSpec,
)
from gwex.parsers._grids import merge_fill

_CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"
_MARKER_RE = re.compile(r"(?=[" + _CIRCLED + r"])|(?=（\s*\d+\s*）)|(?=\(\s*\d+\s*\))")


def extract(data: bytes, sheet: str, mapping: Optional[MappingConfig] = None) -> TestSpec:
    cfg = mapping or DEFAULT_MAPPING
    wb = openpyxl.load_workbook(BytesIO(data), data_only=True)
    ws = wb[sheet]
    grid = merge_fill(ws)

    header = grid[cfg.header_row - 1] if 0 <= cfg.header_row - 1 < len(grid) else None
    col = {role: cfg.column_index(role, header) for role in (
        "screen_name", "medium_category", "small_category",
        "test_no", "verification_content", "execution_steps",
    )}

    spec = TestSpec(sheet=sheet)
    cur_screen: Optional[Screen] = None
    cur_group: Optional[Group] = None
    last = {"screen_name": "", "medium_category": "", "small_category": ""}

    for r in range(cfg.data_start_row - 1, len(grid)):
        row = grid[r]
        # 画面名・中項目・小項目は前方補完（merge_fill 済みだが非結合セルの保険）。
        # ただし階層的に上位（画面→中項目→小項目）が変わったら下位はリセットする。
        # そうしないと、空の小項目を持つ新グループに直前グループの小項目が漏れ込み、
        # 書込→再抽出でキーが変わって dedup の冪等が崩れる。
        sn = _clean(_get(row, col["screen_name"]))
        mc = _clean(_get(row, col["medium_category"]))
        sc = _clean(_get(row, col["small_category"]))
        if sn and sn != last["screen_name"]:
            last["screen_name"] = sn
            last["medium_category"] = ""
            last["small_category"] = ""
        if mc and mc != last["medium_category"]:
            last["medium_category"] = mc
            last["small_category"] = ""
        if sc:
            last["small_category"] = sc
        verification = _clean1(_get(row, col["verification_content"]))
        if not verification:
            continue  # 確認内容が無い行（スペーサ）はケースにしない

        if cur_screen is None or last["screen_name"] != cur_screen.screen_name:
            cur_screen = Screen(screen_name=last["screen_name"])
            spec.screens.append(cur_screen)
            cur_group = None
        if (
            cur_group is None
            or cur_group.medium_category != last["medium_category"]
            or cur_group.small_category != last["small_category"]
        ):
            cur_group = Group(
                medium_category=last["medium_category"],
                small_category=last["small_category"],
            )
            cur_screen.groups.append(cur_group)

        cur_group.cases.append(
            Case(
                test_no=_num(_get(row, col["test_no"])),
                verification_content=verification,
                execution_steps=_split_steps(_get(row, col["execution_steps"]), cfg.steps_split),
                source_row=r + 1,  # grid は 0 始まり → シート行は +1
            )
        )
    return spec


def _get(row: list, idx: Optional[int]) -> Any:
    if idx is None or idx >= len(row):
        return None
    return row[idx]


def _clean(v: Any) -> str:
    """カテゴリ名: 前後空白除去・内部改行は除去（表示用の折返しを畳む）。"""
    if v is None:
        return ""
    return str(v).replace("\n", "").replace("\r", "").strip()


def _clean1(v: Any) -> str:
    """確認内容: 改行は空白化して 1 行に。"""
    if v is None:
        return ""
    return " ".join(str(v).split())


def _num(v: Any) -> Union[int, str]:
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, int):
        return v
    return "" if v is None else str(v).strip()


def _split_steps(v: Any, mode: str) -> list[str]:
    if v is None:
        return []
    text = str(v).replace("\r", "").strip()
    if not text:
        return []
    parts = [p.strip() for p in text.split("\n") if p.strip()]
    if len(parts) > 1 or mode == "newline":
        return parts
    pieces = [p.strip() for p in _MARKER_RE.split(text) if p.strip()]
    return pieces if len(pieces) > 1 else [text]
