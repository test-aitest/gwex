"""構造書き戻しのラウンドトリップ: TestSpec → xlsx → 再抽出 で一致するか。"""

from __future__ import annotations

import openpyxl

from gwmd.domains import testspec, testspec_writer
from gwmd.domains.model import Case, Group, Screen, TestSpec
from gwmd.writer import xlsx_writer


def _spec() -> TestSpec:
    return TestSpec(
        sheet="S",
        screens=[
            Screen(screen_name="画面A", groups=[
                Group(medium_category="中X", small_category="", cases=[
                    Case(test_no=1, verification_content="確認1", execution_steps=["①a", "②b"]),
                    Case(test_no=2, verification_content="確認2", execution_steps=["①a", "②b"]),
                ]),
                Group(medium_category="中Y", small_category="小1", cases=[
                    Case(test_no=1, verification_content="確認3", execution_steps=["①c"]),
                ]),
            ]),
            Screen(screen_name="画面B", groups=[
                Group(medium_category="中Z", small_category="", cases=[
                    Case(test_no=1, verification_content="確認4", execution_steps=["①d"]),
                ]),
            ]),
        ],
    )


def test_roundtrip(tmp_path):
    spec = _spec()
    # 空テンプレ（既定マッピングの data_start_row=10 に書く）
    p = tmp_path / "out.xlsx"
    wb = openpyxl.Workbook()
    wb.active.title = "S"
    wb.save(p)

    asgn = testspec_writer.assignments(spec)
    xlsx_writer.write_cells(str(p), "S", asgn)

    with open(p, "rb") as f:
        back = testspec.extract(f.read(), "S")

    assert [s.screen_name for s in back.screens] == ["画面A", "画面B"]
    a = back.screens[0]
    assert [g.medium_category for g in a.groups] == ["中X", "中Y"]
    assert a.groups[1].small_category == "小1"
    assert [c.test_no for c in a.groups[0].cases] == [1, 2]
    # steps は各ケース行に書かれ、再抽出で復元される
    assert a.groups[0].cases[0].execution_steps == ["①a", "②b"]
    assert a.groups[0].cases[1].execution_steps == ["①a", "②b"]
    assert back.screens[1].groups[0].cases[0].verification_content == "確認4"
