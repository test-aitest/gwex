"""Appium でフローを実走し、スクショ+trigger_rect を収集して spec を完成させる（Phase 3）。

flow-spec が出力した spec（nodes/edges、trigger_component 付き）を入力に、
シミュレータ上でアプリを経路どおりに操作しながら:

- 各画面のスクリーンショットを caps_dir/<画面ID>.png に保存（image を実パスに更新）
- trigger_component の要素矩形を取得し、**スクショ画像のピクセル座標**へ換算して
  trigger_rect に記入（Appium の rect は pt、スクショは px。倍率はウィンドウ幅比で算出）
- 遷移に必要な入力（graph の requiredInputs: タップ/テキスト入力）を実行
- ダッシュボードの促進モーダル等は dismiss ボタン（「閉じる」など）をタップして回避

認証情報は graph から読み取って入力にのみ使い、spec には書き込まない。
"""

from __future__ import annotations

import base64
import io
import os
import time
from typing import Optional

import requests

_FIND_TIMEOUT = 15
_DEFAULT_DISMISS = ("閉じる", "close coachmark")


class AppiumClient:
    """Appium (W3C WebDriver) の最小 HTTP クライアント。"""

    def __init__(self, server: str = "http://127.0.0.1:4723"):
        self.server = server.rstrip("/")
        self.sid: Optional[str] = None

    def _url(self, path: str) -> str:
        return f"{self.server}/session/{self.sid}{path}"

    def start(self, udid: str, bundle_id: str,
              extra: Optional[dict] = None) -> None:
        """extra: 追加 capability（実機の WDA 署名 appium:xcodeOrgId 等）。"""
        always = {
            "platformName": "iOS",
            "appium:automationName": "XCUITest",
            "appium:udid": udid,
            "appium:bundleId": bundle_id,
            "appium:noReset": True,
            "appium:autoAcceptAlerts": True,  # 通知許可等のシステムアラートを自動承諾
            "appium:newCommandTimeout": 600,
        }
        always.update(extra or {})
        r = requests.post(f"{self.server}/session",
                          json={"capabilities": {"alwaysMatch": always}}, timeout=300)
        r.raise_for_status()
        self.sid = r.json()["value"]["sessionId"]

    def quit(self) -> None:
        if self.sid:
            requests.delete(f"{self.server}/session/{self.sid}", timeout=30)
            self.sid = None

    def screenshot_png(self) -> bytes:
        r = requests.get(self._url("/screenshot"), timeout=60)
        r.raise_for_status()
        return base64.b64decode(r.json()["value"])

    def window_width_pt(self) -> float:
        r = requests.get(self._url("/window/rect"), timeout=30)
        r.raise_for_status()
        return float(r.json()["value"]["width"])

    def find(self, value: str, using: str = "accessibility id") -> Optional[str]:
        """要素IDを返す。見つからなければ None（no such element）。"""
        r = requests.post(self._url("/element"),
                          json={"using": using, "value": value}, timeout=_FIND_TIMEOUT)
        if r.status_code != 200:
            return None
        v = r.json()["value"]
        return next(iter(v.values())) if isinstance(v, dict) and v else None

    def rect(self, el: str) -> dict:
        r = requests.get(self._url(f"/element/{el}/rect"), timeout=30)
        r.raise_for_status()
        return r.json()["value"]  # {x, y, width, height} (pt)

    def click(self, el: str) -> None:
        requests.post(self._url(f"/element/{el}/click"), json={}, timeout=60).raise_for_status()

    def send_keys(self, el: str, text: str) -> None:
        requests.post(self._url(f"/element/{el}/value"),
                      json={"text": text}, timeout=60).raise_for_status()

    def visible_names(self, limit: int = 30) -> list[str]:
        """画面上の要素 name 一覧（デバッグ・エラー文脈用）。"""
        import re as _re

        r = requests.get(self._url("/source"), timeout=60)
        if r.status_code != 200 or not isinstance(r.json().get("value"), str):
            return []
        names = _re.findall(r'name="([^"]{1,40})"', r.json()["value"])
        return list(dict.fromkeys(n for n in names if n.strip()))[:limit]


def _find(client, value: str):
    """要素検索。値の末尾が '*' なら前方一致（iOS predicate BEGINSWITH）で探す。

    PHPicker の「次へ」が選択数で「次へ (1)」に名前が変わる等、動的な名前の
    要素に使う（完全一致の accessibility id では突破できない）。
    """
    if value.endswith("*") and len(value) > 1:
        prefix = value[:-1].replace('"', '\\"')
        return client.find(f'name BEGINSWITH "{prefix}"', using="-ios predicate string")
    return client.find(value)


def _nav_for_edge(screens: dict, edge: dict) -> dict:
    """edge に対応する graph の navigator（requiredInputs の取得元）。"""
    navs = [n for n in (screens.get(edge["from"]) or {}).get("navigators") or []
            if n.get("targetScreen") == edge["to"]]
    for n in navs:
        if n.get("triggerComponent") == edge.get("trigger_component"):
            return n
    return navs[0] if navs else {}


def _wait_element(client, value: str, *, dismiss=_DEFAULT_DISMISS,
                  retries: int = 8, interval: float = 1.5) -> str:
    """要素が現れるまで待つ。見つからない間はオーバーレイの dismiss を試す。"""
    for _ in range(retries):
        el = _find(client, value)
        if el:
            return el
        # 見つかった邪魔要素は全てタップする（break すると「写真選択→次へ」のような
        # 2 段の interstitial で最初の要素だけをトグルし続けて詰む）
        for d in dismiss:
            de = _find(client, d)
            if de:
                client.click(de)
        time.sleep(interval)
    names = getattr(client, "visible_names", lambda: [])()
    raise RuntimeError(
        f"要素が見つかりません: {value!r}（オーバーレイ dismiss {dismiss} も試行済み。"
        f"現在画面の要素: {names}）")


def _clear_overlays(client, dismiss=_DEFAULT_DISMISS, *, interval: float = 0.5,
                    rounds: int = 4) -> None:
    """表示中の邪魔要素（促進モーダル/コーチマーク等）を居なくなるまで閉じる。

    対象要素はオーバーレイの下でも要素ツリーに残り find が成功してしまうため、
    タップやスクショの**前に**明示的に掃除する必要がある。
    """
    for _ in range(rounds):
        hit = False
        for d in dismiss:
            el = _find(client, d)
            if el:
                client.click(el)
                hit = True
                time.sleep(interval)
        if not hit:
            return


def _act(client, value: str, action, *, dismiss=_DEFAULT_DISMISS,
         interval: float = 1.5, attempts: int = 3):
    """要素を探して action(el) を実行。stale 等で失敗したら取り直してリトライ。"""
    for i in range(attempts):
        el = _wait_element(client, value, dismiss=dismiss, interval=interval)
        try:
            return action(el)
        except Exception:
            if i == attempts - 1:
                raise
            time.sleep(max(interval, 0.5))


def _transition_marker(screens: dict, src: str, dst: str) -> Optional[str]:
    """遷移成功の目印: dst にあって src に無い最初のコンポーネントID。"""
    src_ids = {c.get("accessibilityId") or c.get("name")
               for c in (screens.get(src) or {}).get("components") or []}
    for c in (screens.get(dst) or {}).get("components") or []:
        mid = c.get("accessibilityId") or c.get("name")
        if mid and mid not in src_ids:
            return mid
    return None


def _click_through(client, trig: str, marker: Optional[str], *,
                   dismiss=_DEFAULT_DISMISS, interval: float = 1.5,
                   attempts: int = 4) -> None:
    """trig をタップし、遷移が観測できるまで（掃除して）タップし直す。

    成功条件: marker（遷移先の目印要素）が現れる、または trig が画面から消える。
    タップの瞬間にモーダル/コーチマークが被さってタップが食われるケースへの対策。
    """
    for _ in range(attempts):
        _clear_overlays(client, dismiss)
        el = _find(client, trig)
        if el:
            try:
                client.click(el)
            except Exception:
                pass  # stale 等は取り直して再試行
        time.sleep(max(interval, 1.0))
        if marker and _find(client, marker):
            return
        if not _find(client, trig):
            return
    names = getattr(client, "visible_names", lambda: [])()
    raise RuntimeError(
        f"タップしても遷移しません: {trig!r}（期待した目印: {marker!r}。"
        f"現在画面の要素: {names}）")


def _shot(client, caps_dir: str, node: dict) -> tuple[str, float]:
    """現在画面を保存し、(保存パス, px/pt倍率) を返す。"""
    from PIL import Image as PILImage

    png = client.screenshot_png()
    path = os.path.join(caps_dir, f'{node["id"]}.png')
    with open(path, "wb") as f:
        f.write(png)
    with PILImage.open(io.BytesIO(png)) as im:
        scale = im.size[0] / client.window_width_pt()
    return path, scale


def rewind(client, back, then: Optional[str] = None, *,
           dismiss=_DEFAULT_DISMISS, max_backs: int = 8, settle: float = 1.0) -> None:
    """既知の「戻る」ボタン群を見えなくなるまで連打して起点近くへ戻す（実走前リセット用）。

    アプリが前回の途中状態を復元して起動するケースに使う。back は文字列または
    候補リスト（画面により戻るが「01 btn back off」だったり「closemark」だったり
    するため、各ラウンドで最初に見つかった候補をタップする）。then を指定すると
    最後にそれ（ホームタブ等）をタップして起点画面を確定させる。
    """
    backs = [back] if isinstance(back, str) else list(back)
    for _ in range(max_backs):
        _clear_overlays(client, dismiss)
        el = None
        for b in backs:
            el = _find(client, b)
            if el:
                break
        if not el:
            break
        client.click(el)
        time.sleep(max(settle, 0.5))
    if then:
        el = _find(client, then)
        if el:
            client.click(el)
            time.sleep(max(settle, 0.5))


def capture_flow(spec: dict, screens: dict, client, caps_dir: str,
                 *, settle: float = 2.0, dismiss=_DEFAULT_DISMISS) -> dict:
    """spec の edges を順に実走してスクショ+trigger_rect を埋め、更新済み spec を返す。

    client は AppiumClient 互換（find/rect/click/send_keys/screenshot_png/window_width_pt）。
    edges の並び＝実行順。edge.from のスクショは「入力適用後・タップ直前」の状態。
    edges に現れない row バリエーションのノードは対象外（image は手動配置のまま）。
    """
    os.makedirs(caps_dir, exist_ok=True)
    by_id = {n["id"]: n for n in spec.get("nodes") or []}
    interval = settle if settle > 0 else 0.05

    for edge in spec.get("edges") or []:
        if edge.get("skip_capture"):
            continue  # 描画専用エッジ（合流矢印等）。実走・撮影はしない
        node = by_id[edge["from"]]
        trig = edge.get("trigger_component") or ""
        if not trig:
            raise ValueError(f"trigger_component がありません: {edge['from']} → {edge['to']}")
        try:
            _wait_element(client, trig, dismiss=dismiss, interval=interval)  # 画面到達+オーバーレイ回避
        except RuntimeError:
            if edge.get("optional"):
                continue  # 任意エッジ: 通らない経路（自動ログイン済みでログイン画面が出ない等）
            raise
        _clear_overlays(client, dismiss)  # 被さっているモーダル/コーチマークを掃除

        if edge.get("detour"):
            # spec 拡張: detour = 寄り道エッジ。トリガーをタップして to（モーダル等）を
            # 撮影し、detour で指定した要素をタップして from へ戻る。戻り要素が
            # 「閉じる」等の dismiss 語と同名でも競合しないよう、この間は掃除しない。
            path, scale = _shot(client, caps_dir, node)
            node["image"] = path
            r = _act(client, trig, client.rect, dismiss=dismiss, interval=interval)
            edge["trigger_rect"] = [round(r["x"] * scale), round(r["y"] * scale),
                                    round(r["width"] * scale), round(r["height"] * scale)]
            _act(client, trig, client.click, dismiss=dismiss, interval=interval)
            time.sleep(settle)
            tgt = by_id[edge["to"]]
            tgt["image"] = _shot(client, caps_dir, tgt)[0]
            _act(client, str(edge["detour"]), client.click, dismiss=(), interval=interval)
            time.sleep(settle)
            continue

        for key, val in (_nav_for_edge(screens, edge).get("requiredInputs") or {}).items():
            if str(val).strip().lower() == "tap":
                _act(client, key, client.click, dismiss=dismiss, interval=interval)
            else:
                _act(client, key, lambda el, v=str(val): client.send_keys(el, v),
                     dismiss=dismiss, interval=interval)
            time.sleep(min(1.0, max(interval, 0.3)))  # 入力反映（タブ切替アニメ等）を待つ

        # spec 拡張: pre = トリガー前にタップする要素（金額プリセット等。
        # ボタンが入力するまで無効なケースに使う）
        for pv in edge.get("pre") or []:
            _act(client, str(pv), client.click, dismiss=dismiss, interval=interval)
            time.sleep(min(1.0, max(interval, 0.3)))

        _clear_overlays(client, dismiss)  # スクショ・タップ直前にもう一度掃除
        path, scale = _shot(client, caps_dir, node)
        node["image"] = path
        r = _act(client, trig, client.rect, dismiss=dismiss, interval=interval)  # 取り直し（stale対策）
        edge["trigger_rect"] = [round(r["x"] * scale), round(r["y"] * scale),
                                round(r["width"] * scale), round(r["height"] * scale)]
        _click_through(client, trig, _transition_marker(screens, edge["from"], edge["to"]),
                       dismiss=dismiss, interval=max(settle, 1.0))
        time.sleep(settle)

    edges = spec.get("edges") or []
    if edges:  # 最終画面のスクショ
        last = by_id[edges[-1]["to"]]
        time.sleep(settle)
        path, _ = _shot(client, caps_dir, last)
        last["image"] = path
    return spec
