---
name: capture-screen
description: 起動中の iOS シミュレータ / Android エミュレータの画面をキャプチャし、テスト仕様シートの「画面イメージ」シートの所定セルに、修正前/後ラベル付きで添付する。実装の修正前後の見た目を概要設計書に残したいときに使う。
---

# capture-screen

**今表示中**のシミュレータ/エミュレータ画面をキャプチャし、`gwex set-image` で
「画面イメージ」シートの指定セルへ添付するスキル。画像は自動縮小して入れる
（Google はセル内画像が base64 でサイズ上限があるため必須）。

> ビルド切替・アプリ起動・対象画面への遷移は、このスキルの対象外。
> あなた（agent）が対象アプリのリポジトリで bash により行い、目的の画面を表示した状態にしてから本スキルを使う。

## 入力

- `--target <Google URL | .xlsx パス>`：書き込み先。
- `--sheet <名>`：`2.画面イメージ(iOS)` または `2.画面イメージ(Android)` 等。
- `--cell <アンカー>`：画像枠の左上セル（例 `C7`）。**毎回明示**。
- `--range <範囲>`：画像枠として結合する範囲（例 `C7:F20`）。指定すると範囲を結合し、画像を**縦横比維持**で枠内にフィットさせ、セルの**上にオーバーレイ**で乗せる。
- `--insert-rows`：枠ぶんの行を先に挿入し、既存内容を押し下げて確実に空枠を作る（既存レイアウトを壊さない）。
- `--platform ios|android`。
- `--label 修正前|修正後`（任意）：隣接セルにラベルを書く。
- 端末指定（任意）：iOS `--udid <UDID>` / Android `--serial <serial>`。

## 手順

1. **キャプチャ**（一時 PNG に保存）：
   - iOS：`xcrun simctl io booted screenshot /tmp/cap.png`（`--udid` 指定時は `booted` の代わりに UDID）。
   - Android：`adb exec-out screencap -p > /tmp/cap.png`（`--serial` 指定時は `adb -s <serial> ...`）。
2. **確認提示**：対象シート・セル・ラベル・撮れた画像（パス）をユーザーに示し、**承認を待つ**。
3. **添付（承認後）**：範囲を結合し画像を縦横比維持でオーバーレイ。`--insert-rows` で確実に空枠を作る。
   ```
   gwex set-image "<target>" --sheet "<sheet>" --cell "<左上セル>" --range "<範囲>" --insert-rows --image /tmp/cap.png --max-dim 1024
   ```
   - ラベル指定時は隣接セル（例：画像セルの1つ上 or 左）に
     ```
     gwex set-text "<target>" --sheet "<sheet>" --cell "<label_cell>" --text "<修正前|修正後>"
     ```
4. **修正前/後の運用**：修正前ビルドで表示→本スキルで `--label 修正前` を所定セルへ、
   修正後ビルドで表示→隣接セルへ `--label 修正後`。ビルド切替は agent の bash 作業。

## 注意

- 事前に**シミュレータ/エミュレータが起動**し、**対象画面が表示**されていること。
- シート書き込みは破壊的。**承認後**に実行する。
- Google スプレッドシートは書込みスコープ認証済み（`gwex auth`）＋ セル内画像は Apps Script Web アプリ（`GWEX_APPSCRIPT_URL`）が必要。
- 大きなスクショは `--max-dim`（既定運用 1024）で縮小して添付する。
