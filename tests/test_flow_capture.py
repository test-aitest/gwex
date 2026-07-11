"""flow_capture: Appium 実走によるスクショ+trigger_rect 収集（FakeClient で検証）。

- 要素待ち中のオーバーレイ dismiss
- requiredInputs の適用（tap / テキスト入力）と spec への非転記
- rect の pt→px 換算（スクショ幅 ÷ ウィンドウ幅）
- スクショ保存と spec の image / trigger_rect 更新
- 末尾 '*' の前方一致（「次へ」→「次へ (1)」の動的名対応）
"""

from __future__ import annotations

import io
import re

import pytest

from gwex.domains import flow_capture

PIL = pytest.importorskip("PIL.Image")

WIN_PT = 100          # ウィンドウ幅(pt)
PNG_W, PNG_H = 200, 400  # スクショ(px) → 倍率 2.0


def _png_bytes() -> bytes:
    buf = io.BytesIO()
    PIL.new("RGB", (PNG_W, PNG_H), (30, 30, 30)).save(buf, format="PNG")
    return buf.getvalue()


class FakeClient:
    """画面ごとの要素と rect を持ち、click で画面を進める Appium 互換フェイク。"""

    def __init__(self, screens_elements, transitions, overlay=None):
        self.elements = screens_elements  # screen -> {elementValue: rect(pt)}
        self.transitions = transitions    # (screen, elementValue) -> next screen
        self.current = next(iter(screens_elements))
        self.overlay = overlay            # 表示中オーバーレイの要素名（dismiss で消える）
        self.log: list[tuple] = []

    def find(self, value, using="accessibility id"):
        names = [self.overlay] if self.overlay else list(self.elements[self.current])
        if using == "-ios predicate string":
            # `name BEGINSWITH "..."` のみエミュレート
            m = re.match(r'name BEGINSWITH "(.*)"$', value)
            return next((n for n in names if m and n.startswith(m.group(1))), None)
        return value if value in names else None

    def rect(self, el):
        return self.elements[self.current][el]

    def click(self, el):
        self.log.append(("click", self.current, el))
        if self.overlay and el == self.overlay:
            self.overlay = None
            return
        nxt = self.transitions.get((self.current, el))
        if nxt:
            self.current = nxt

    def send_keys(self, el, text):
        self.log.append(("input", self.current, el, text))

    def screenshot_png(self):
        return _png_bytes()

    def window_width_pt(self):
        return WIN_PT


GRAPH = {
    "Login": {
        "components": [],
        "navigators": [{
            "targetScreen": "Home", "type": "TAP", "triggerComponent": "submit",
            "requiredInputs": {"idfield": "user-1", "select_tab": "tap"},
        }],
    },
    "Home": {"components": [], "navigators": []},
}

SPEC = {
    "sheet": "F",
    "nodes": [
        {"id": "Login", "name": "Login", "image": "caps/Login.png", "col": 0, "row": 0},
        {"id": "Home", "name": "Home", "image": "caps/Home.png", "col": 1, "row": 0},
    ],
    "edges": [{"from": "Login", "to": "Home", "label": "ログイン", "trigger_component": "submit"}],
}


def _client(overlay=None):
    return FakeClient(
        screens_elements={
            "Login": {"submit": {"x": 10, "y": 20, "width": 30, "height": 15},
                      "idfield": {"x": 0, "y": 0, "width": 50, "height": 10},
                      "select_tab": {"x": 0, "y": 40, "width": 50, "height": 10}},
            "Home": {"menu": {"x": 0, "y": 0, "width": 10, "height": 10}},
        },
        transitions={("Login", "submit"): "Home"},
        overlay=overlay,
    )


def test_capture_fills_rect_and_images(tmp_path):
    import copy
    spec = copy.deepcopy(SPEC)
    client = _client()
    out = flow_capture.capture_flow(spec, GRAPH, client, str(tmp_path / "caps"), settle=0)

    # trigger_rect は pt×2 の px 値
    assert out["edges"][0]["trigger_rect"] == [20, 40, 60, 30]
    # スクショが保存され image が実パスに更新される
    assert out["nodes"][0]["image"].endswith("Login.png")
    assert out["nodes"][1]["image"].endswith("Home.png")
    assert (tmp_path / "caps" / "Login.png").exists()
    assert (tmp_path / "caps" / "Home.png").exists()
    # requiredInputs: テキスト入力と tap が実行され、spec に認証情報は残らない
    assert ("input", "Login", "idfield", "user-1") in client.log
    assert ("click", "Login", "select_tab") in client.log
    assert "user-1" not in str(out)


def test_overlay_dismissed_before_trigger(tmp_path):
    import copy
    client = _client(overlay="閉じる")
    flow_capture.capture_flow(copy.deepcopy(SPEC), GRAPH, client, str(tmp_path / "caps"), settle=0)
    assert client.log[0] == ("click", "Login", "閉じる")  # 先にオーバーレイを閉じる


def test_prefix_match_dismiss(tmp_path):
    """dismiss の '次へ*' が選択数つきの「次へ (1)」にマッチして突破できる。"""
    import copy
    client = _client(overlay="次へ (1)")  # ピッカー相当のオーバーレイ（動的名）
    out = flow_capture.capture_flow(copy.deepcopy(SPEC), GRAPH, client,
                                    str(tmp_path / "caps"), settle=0,
                                    dismiss=("次へ*",))
    assert client.log[0] == ("click", "Login", "次へ (1)")  # 前方一致でタップ
    assert out["edges"][0]["trigger_rect"] == [20, 40, 60, 30]


def test_prefix_match_wait_target():
    """待ち対象そのものにも '*' が使える。"""
    client = _client()
    el = flow_capture._wait_element(client, "sub*", retries=1, interval=0)
    assert el == "submit"


def test_stale_element_retried(tmp_path):
    """send_keys が一度失敗（stale相当）しても、要素を取り直して成功する。"""
    import copy

    class FlakyClient(FakeClient):
        def __init__(self, *a, **kw):
            super().__init__(*a, **kw)
            self.fail_once = True

        def send_keys(self, el, text):
            if self.fail_once:
                self.fail_once = False
                raise RuntimeError("stale element reference")
            super().send_keys(el, text)

    client = FlakyClient(
        screens_elements=_client().elements, transitions={("Login", "submit"): "Home"})
    out = flow_capture.capture_flow(copy.deepcopy(SPEC), GRAPH, client,
                                    str(tmp_path / "caps"), settle=0)
    assert ("input", "Login", "idfield", "user-1") in client.log  # リトライで成功
    assert out["edges"][0]["trigger_rect"] == [20, 40, 60, 30]


def test_detour_edge_visits_modal_and_returns(tmp_path):
    """detour: トリガー→モーダル撮影→戻り要素タップで復帰。dismiss と競合しない。"""
    import copy

    client = FakeClient(
        screens_elements={
            "Login": {"submit": {"x": 10, "y": 20, "width": 30, "height": 15},
                      "dots": {"x": 80, "y": 20, "width": 10, "height": 10},
                      "idfield": {"x": 0, "y": 0, "width": 50, "height": 10},
                      "select_tab": {"x": 0, "y": 40, "width": 50, "height": 10}},
            "Modal": {"閉じる": {"x": 0, "y": 90, "width": 20, "height": 8}},
            "Home": {"menu": {"x": 0, "y": 0, "width": 10, "height": 10}},
        },
        transitions={("Login", "submit"): "Home",
                     ("Login", "dots"): "Modal",
                     ("Modal", "閉じる"): "Login"},
    )
    spec = copy.deepcopy(SPEC)
    spec["nodes"].append({"id": "Modal", "name": "モーダル", "image": "caps/Modal.png",
                          "col": 1, "row": 1})
    spec["edges"] = [
        {"from": "Login", "to": "Modal", "trigger_component": "dots", "detour": "閉じる"},
        spec["edges"][0],
    ]
    out = flow_capture.capture_flow(spec, GRAPH, client, str(tmp_path / "caps"), settle=0)

    assert (tmp_path / "caps" / "Modal.png").exists()
    assert out["edges"][0]["trigger_rect"] == [160, 40, 20, 20]  # dots の rect ×2
    assert ("click", "Modal", "閉じる") in client.log            # 戻りタップ
    assert client.current == "Home"                              # その後 通常エッジで遷移完了


def test_optional_edge_skipped_when_absent(tmp_path):
    """optional エッジ: トリガーが画面に無ければスキップして続行する。"""
    import copy
    spec = copy.deepcopy(SPEC)
    spec["edges"].insert(0, {"from": "Login", "to": "Home", "optional": True,
                             "trigger_component": "gone_button"})
    out = flow_capture.capture_flow(spec, GRAPH, _client(), str(tmp_path / "caps"), settle=0)
    assert "trigger_rect" not in out["edges"][0]                  # スキップされた
    assert out["edges"][1]["trigger_rect"] == [20, 40, 60, 30]    # 後続エッジは実走


def test_skip_capture_edge_ignored(tmp_path):
    """skip_capture エッジは実走されない（描画専用）。"""
    import copy
    spec = copy.deepcopy(SPEC)
    spec["edges"].insert(0, {"from": "Home", "to": "Login", "skip_capture": True,
                             "trigger_component": "no_such"})
    out = flow_capture.capture_flow(spec, GRAPH, _client(), str(tmp_path / "caps"), settle=0)
    assert "trigger_rect" not in out["edges"][0]          # 実走されていない
    assert out["edges"][1]["trigger_rect"] == [20, 40, 60, 30]  # 通常エッジは処理される


def test_missing_trigger_raises(tmp_path):
    import copy
    spec = copy.deepcopy(SPEC)
    spec["edges"][0]["trigger_component"] = "no_such_btn"
    with pytest.raises(RuntimeError, match="no_such_btn"):
        flow_capture.capture_flow(spec, GRAPH, _client(), str(tmp_path / "caps"), settle=0)


def test_missing_trigger_component_key(tmp_path):
    import copy
    spec = copy.deepcopy(SPEC)
    spec["edges"][0]["trigger_component"] = ""
    with pytest.raises(ValueError, match="trigger_component"):
        flow_capture.capture_flow(spec, GRAPH, _client(), str(tmp_path / "caps"), settle=0)
