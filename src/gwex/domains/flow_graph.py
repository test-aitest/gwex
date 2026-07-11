"""devpilot スクリーングラフ（YAML）から画面遷移図 spec の雛形を生成する。

graph 形式（devpilot-graph.yaml。UI クローラーが構築する）:

    screens:
      <ScreenName>:
        components: [{name, type, accessibilityId, label}, ...]
        navigators: [{targetScreen, type, triggerComponent, description, ...}, ...]

出力は `gwex draw-flow` が読む spec（nodes/edges）。画像パスと trigger_rect は
Appium での自動収集（Phase 3）または手動で埋める前提の TODO プレースホルダ。
requiredInputs（テスト認証情報等）は spec に**転記しない**。
"""

from __future__ import annotations

from collections import deque
from typing import Optional


def load_graph(path: str) -> dict:
    """graph YAML を読み、screens 辞書を返す。"""
    import yaml

    with open(path, encoding="utf-8") as f:
        data = yaml.safe_load(f)
    screens = (data or {}).get("screens")
    if not screens:
        raise ValueError(f"screens が見つかりません: {path}")
    return screens


def _navigators(screens: dict, name: str) -> list[dict]:
    return (screens.get(name) or {}).get("navigators") or []


def _bfs(screens: dict, start: str, goal: str) -> Optional[list[str]]:
    """start→goal の最短経路（画面名リスト）。無ければ None。"""
    prev: dict[str, Optional[str]] = {start: None}
    q = deque([start])
    while q:
        cur = q.popleft()
        if cur == goal:
            seg = [cur]
            while prev[cur] is not None:
                cur = prev[cur]  # type: ignore[assignment]
                seg.append(cur)
            return seg[::-1]
        for nav in _navigators(screens, cur):
            t = nav.get("targetScreen")
            if t and t in screens and t not in prev:
                prev[t] = cur
                q.append(t)
    return None


def find_path(screens: dict, start: str, goal: str,
              via: Optional[list[str]] = None) -> list[str]:
    """start →(via…)→ goal の最短経路。経路が無い/画面名が不明なら ValueError。"""
    for s in [start, goal, *(via or [])]:
        if s not in screens:
            candidates = ", ".join(sorted(screens))
            raise ValueError(f"画面がグラフにありません: {s}（候補: {candidates}）")
    waypoints = [start, *(via or []), goal]
    path = [start]
    for a, b in zip(waypoints, waypoints[1:]):
        seg = _bfs(screens, a, b)
        if seg is None:
            raise ValueError(f"経路が見つかりません: {a} → {b}")
        path += seg[1:]
    return path


def _edge_nav(screens: dict, src: str, dst: str) -> dict:
    """src→dst の navigator（複数あれば先頭）。"""
    for nav in _navigators(screens, src):
        if nav.get("targetScreen") == dst:
            return nav
    return {}


def _component_label(screens: dict, screen: str, comp: str) -> str:
    """トリガーコンポーネントの表示ラベル（無ければコンポーネント名で代用）。"""
    for c in (screens.get(screen) or {}).get("components") or []:
        if comp and (c.get("name") == comp or c.get("accessibilityId") == comp):
            return c.get("label") or comp
    return comp or ""


def _short(name: str) -> str:
    return name.removesuffix("ViewController")


def build_spec(screens: dict, path: list[str], *,
               sheet: Optional[str] = None, image_dir: str = "caps") -> dict:
    """経路から draw-flow 用 spec を組み立てる（列=経路順、行=0）。

    edges の trigger_component は Appium で trigger_rect を自動取得するためのキー
    （accessibilityId / コンポーネント名）。draw-flow は未知キーを無視するので
    spec にそのまま残してよい。
    """
    title = sheet or f"画面遷移図_{_short(path[0])}→{_short(path[-1])}"
    if len(title) > 31:  # Excel のシート名上限
        title = title[:31]
    nodes = [
        {"id": s, "name": _short(s), "image": f"{image_dir}/{s}.png", "col": i, "row": 0}
        for i, s in enumerate(path)
    ]
    edges = []
    for a, b in zip(path, path[1:]):
        nav = _edge_nav(screens, a, b)
        trig = nav.get("triggerComponent") or ""
        edges.append({
            "from": a,
            "to": b,
            "label": _component_label(screens, a, trig),
            "trigger_component": trig,
        })
    return {"sheet": title, "nodes": nodes, "edges": edges}


def _q(s: str) -> str:
    """YAML 用の最小クオート（ダブルクォート囲い）。"""
    return '"' + str(s).replace("\\", "\\\\").replace('"', '\\"') + '"'


def render_spec_yaml(spec: dict) -> str:
    """spec を、記入すべき箇所の TODO コメント付き YAML として整形する。"""
    lines = [
        "# gwex flow-spec が生成した画面遷移図 spec の雛形。",
        "# TODO: 各 node の image にスクリーンショットを配置する。",
        "# TODO: 各 edge に trigger_rect（元画像px [x, y, w, h]）を追記すると赤枠が描かれる。",
        f"sheet: {_q(spec['sheet'])}",
        "nodes:",
    ]
    for n in spec["nodes"]:
        lines.append(
            f'  - {{id: {n["id"]}, name: {_q(n["name"])}, '
            f'image: {_q(n["image"])}, col: {n["col"]}, row: {n["row"]}}}'
        )
    lines.append("edges:")
    for e in spec["edges"]:
        lines.append(f"  - from: {e['from']}")
        lines.append(f"    to: {e['to']}")
        lines.append(f"    label: {_q(e['label'])}")
        lines.append(f"    trigger_component: {_q(e['trigger_component'])}"
                     "  # Appium 要素特定用（trigger_rect 自動取得の入力）")
        lines.append("    # trigger_rect: [x, y, w, h]")
        lines.append("    # via: 処理名   # 矢印の途中に処理ボックスを挟む場合")
        lines.append('    # pre: ["要素ID"]   # トリガー前にタップする要素（金額プリセット等）')
    return "\n".join(lines) + "\n"
