"""TestSpec 階層シリアライザ（md/json）のテスト。"""

from __future__ import annotations

import json

from gwex.domains.model import Case, Group, Screen, TestSpec
from gwex.serializers import testspec_json, testspec_md


def _spec() -> TestSpec:
    return TestSpec(
        sheet="3.結合テスト項目",
        screens=[
            Screen(
                screen_name="属性変更_登録情報確認",
                groups=[
                    Group(
                        medium_category="本人情報_氏名",
                        small_category="",
                        cases=[
                            Case(test_no=1, verification_content="…マスキング…",
                                 execution_steps=["①遷移する", "（②実行）"]),
                            Case(test_no=2, verification_content="…同じこと",
                                 execution_steps=["①遷移する"]),
                        ],
                    )
                ],
            )
        ],
    )


def test_markdown_hierarchy():
    out = testspec_md.render(_spec())
    assert "# 3.結合テスト項目" in out
    assert "### 画面: 属性変更_登録情報確認" in out
    assert "#### 中項目: 本人情報_氏名" in out
    assert "| No | 確認内容 | 実施手順 |" in out
    assert "| 1 | …マスキング… | ①遷移する<br>（②実行） |" in out
    # 画面名は1回だけ（親を繰り返さない）
    assert out.count("属性変更_登録情報確認") == 1


def test_json_nested():
    data = json.loads(testspec_json.render(_spec()))
    assert data["sheet"] == "3.結合テスト項目"
    screen = data["screens"][0]
    assert screen["screen_name"] == "属性変更_登録情報確認"
    group = screen["groups"][0]
    assert group["medium_category"] == "本人情報_氏名"
    assert group["cases"][0]["execution_steps"] == ["①遷移する", "（②実行）"]
