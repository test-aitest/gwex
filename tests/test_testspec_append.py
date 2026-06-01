"""append / dedup（冪等）/ 採番継続 / diff 分類のテスト。tmp の xlsx で非破壊。"""

from __future__ import annotations

import openpyxl

from gwex import core
from gwex.domains import testspec
from gwex.domains.model import Case, Group, Screen, TestSpec
from gwex.domains.testspec_merge import diff


def _blank(tmp_path) -> str:
    p = tmp_path / "s.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "S"
    wb.save(p)
    return str(p)


def _one(screen, medium, no, verification, steps):
    return TestSpec(sheet="S", screens=[
        Screen(screen_name=screen, groups=[
            Group(medium_category=medium, cases=[
                Case(test_no=no, verification_content=verification, execution_steps=steps)
            ])
        ])
    ])


def _extract(path):
    with open(path, "rb") as f:
        return testspec.extract(f.read(), "S")


def test_append_continues_numbering(tmp_path):
    p = _blank(tmp_path)
    core.write_testspec(_one("画面A", "中X", 1, "確認1", ["①a"]).model_dump_json(), p, "S")
    # 同一グループへ append（test_no=99 は無視され既存最大+1=2 になる）
    core.write_testspec(_one("画面A", "中X", 99, "確認2", ["①b"]).model_dump_json(), p, "S", append=True, dedup=True)

    g = _extract(p).screens[0].groups[0]
    assert [c.verification_content for c in g.cases] == ["確認1", "確認2"]
    assert [c.test_no for c in g.cases] == [1, 2]
    assert g.cases[0].source_row is not None


def test_dedup_idempotent(tmp_path):
    p = _blank(tmp_path)
    spec = _one("A", "X", 1, "確認1", ["①a"]).model_dump_json()
    core.write_testspec(spec, p, "S")                      # seed
    core.write_testspec(spec, p, "S", append=True, dedup=True)  # 同一 → 0件
    core.write_testspec(spec, p, "S", append=True, dedup=True)  # 同一 → 0件
    total = sum(len(g.cases) for s in _extract(p).screens for g in s.groups)
    assert total == 1


def test_diff_classification():
    existing = _one("A", "X", 1, "確認1", [])
    new = TestSpec(sheet="S", screens=[
        Screen(screen_name="A", groups=[Group(medium_category="X", cases=[
            Case(test_no=1, verification_content="確認1"),   # 完全一致
            Case(test_no=2, verification_content="確認2"),   # 近似（同一グループ・内容差）
        ])]),
        Screen(screen_name="B", groups=[Group(medium_category="Y", cases=[
            Case(test_no=1, verification_content="確認3"),   # 新規
        ])]),
    ])
    r = diff(existing, new)
    assert len(r.exact_dups) == 1
    assert len(r.near_candidates) == 1 and r.near_candidates[0].candidates
    assert len(r.new_cases) == 1
