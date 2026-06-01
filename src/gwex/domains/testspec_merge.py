"""既存 TestSpec と生成 TestSpec の重複検知・差分分類（pure）。

決定論的に「完全一致 / 近似候補 / 新規」へ分類する。近似候補の最終判断
（同一テストの更新か別物か）は呼び出し側エージェントの LLM が行う。
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field

from gwex.domains.model import Case, TestSpec


def _norm(s: str | None) -> str:
    """正規化: NFKC（全角/半角・記号差吸収）＋空白除去。"""
    s = unicodedata.normalize("NFKC", s or "")
    return re.sub(r"\s+", "", s)


def case_key(screen: str, medium: str, small: str, verification: str) -> tuple[str, str, str, str]:
    return (_norm(screen), _norm(medium), _norm(small), _norm(verification))


def group_key(screen: str, medium: str, small: str) -> tuple[str, str, str]:
    return (_norm(screen), _norm(medium), _norm(small))


@dataclass
class ClassifiedCase:
    screen_name: str
    medium_category: str
    small_category: str
    case: Case
    candidates: list[Case] = field(default_factory=list)  # 近似時: 同一グループの既存ケース


@dataclass
class DiffResult:
    new_cases: list[ClassifiedCase]      # 既存に無い → 追記
    exact_dups: list[ClassifiedCase]     # 完全一致 → skip（冪等）
    near_candidates: list[ClassifiedCase]  # 同一グループだが確認内容差 → LLM が更新/別物を判断


def diff(existing: TestSpec, new: TestSpec) -> DiffResult:
    ex_keys: dict[tuple, Case] = {}
    ex_groups: dict[tuple, list[Case]] = {}
    for sc in existing.screens:
        for g in sc.groups:
            for c in g.cases:
                ex_keys[case_key(sc.screen_name, g.medium_category, g.small_category, c.verification_content)] = c
                ex_groups.setdefault(group_key(sc.screen_name, g.medium_category, g.small_category), []).append(c)

    new_cases: list[ClassifiedCase] = []
    exact_dups: list[ClassifiedCase] = []
    near_candidates: list[ClassifiedCase] = []
    for sc in new.screens:
        for g in sc.groups:
            for c in g.cases:
                item = ClassifiedCase(sc.screen_name, g.medium_category, g.small_category, c)
                ck = case_key(sc.screen_name, g.medium_category, g.small_category, c.verification_content)
                gk = group_key(sc.screen_name, g.medium_category, g.small_category)
                if ck in ex_keys:
                    exact_dups.append(item)
                elif gk in ex_groups:
                    item.candidates = list(ex_groups[gk])
                    near_candidates.append(item)
                else:
                    new_cases.append(item)
    return DiffResult(new_cases, exact_dups, near_candidates)
