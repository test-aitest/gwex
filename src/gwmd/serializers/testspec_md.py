"""TestSpec（階層）→ Markdown。画面/中項目を見出しにし、配下に小表を出す。pure。"""

from __future__ import annotations

from gwmd.domains.model import Group, TestSpec


def render(spec: TestSpec) -> str:
    out: list[str] = [f"# {spec.sheet}"]
    for screen in spec.screens:
        out.append(f"### 画面: {screen.screen_name}")
        for group in screen.groups:
            out.append(_group_heading(group))
            out.append(_table(group))
    return "\n\n".join(out) + "\n"


def _group_heading(group: Group) -> str:
    label = group.medium_category
    if group.small_category:
        label += f" / {group.small_category}"
    return f"#### 中項目: {label}"


def _table(group: Group) -> str:
    lines = ["| No | 確認内容 | 実施手順 |", "| --- | --- | --- |"]
    for case in group.cases:
        steps = "<br>".join(case.execution_steps)
        content = _esc(case.verification_content)
        lines.append(f"| {case.test_no} | {content} | {_esc(steps)} |")
    return "\n".join(lines)


def _esc(s: str) -> str:
    return s.replace("\\", "\\\\").replace("|", "\\|").replace("\n", "<br>")
