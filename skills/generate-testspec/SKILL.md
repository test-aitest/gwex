---
name: generate-testspec
description: PR/差分から結合テスト項目（確認内容＋実施手順）を生成し、レビュー後に gwex でテスト仕様シートへ追記/更新する。対象は Excel(.xlsx) または Google スプレッドシート。実装からテスト観点を起こしたいときに使う。
---

# generate-testspec

実装の差分（PR/ブランチ）を読み、**結合テスト項目**を生成して既存テスト仕様シートに
**追記/更新**するスキル。生成・判断はあなた（LLM）が行い、シートへの配置は `gwex` CLI が担う
（gwex は LLM を持たない typed 書き込みツール）。

## 入力（ユーザーから受け取る）

- `--pr <番号>`：GitHub PR の差分を使う（主）。無ければ `--branch <base>`（`git diff <base>...HEAD`）。
- `--target <Google URL | .xlsx パス>`：書き込み先シートファイル。
- `--sheet <名>`：対象シート（既定 `3.結合テスト項目`）。
- `--mapping <toml>`（任意）：列マッピング。省略時は gwex 既定（実テンプレ準拠）。

## 手順

1. **差分取得**：`gh pr diff <番号>`（無ければ `git diff <base>...HEAD`）。変更ファイル一覧と内容を把握し、
   関連する画面/コンポーネントのソースも読んで文脈を掴む。
2. **既存仕様の取得**：
   ```
   gwex testspec "<target>" --sheet "<sheet>" --to json
   ```
   返る TestSpec（各ケースに `source_row` 付き）で、既存の画面名/中項目名の命名・既存ケースを把握する。
3. **生成**：下記スキーマの TestSpec(JSON) を作る。**スタイルガイドと生成範囲**に従う。
4. **マージ分類**（既存と突き合わせ）：
   - **完全一致**（画面+中項目+小項目+確認内容が同じ）→ skip（gwex の `--dedup` でも自動除外）。
   - **近似**（同じ画面+中項目+小項目だが確認内容が違う）→ 「同一テストの**更新**か**別物**か」をあなたが判断。
   - **新規** → 追記。
5. **レビュー提示**：「新規 N 件 / 更新 M 件 / skip K 件」を Markdown の表で人に見せ、**承認を待つ**。
6. **書き込み（承認後）**：
   - 新規：新規ケースだけの TestSpec を一時 JSON に保存し
     ```
     gwex write-testspec "<target>" --sheet "<sheet>" --json /tmp/new.json --append --dedup
     ```
     （`--append` で既存の下に追記、`--dedup` で完全一致除外＝冪等。test_no は既存グループ最大Noから自動継続）
   - 更新：近似で「更新」と判断したものは、既存ケースの `source_row` を使い
     ```
     gwex update-case "<target>" --sheet "<sheet>" --row <source_row> \
       --verification "<確認内容>" --steps $'①…\n②…'
     ```

## TestSpec JSON スキーマ（生成形）

```json
{
  "sheet": "3.結合テスト項目",
  "screens": [
    {
      "screen_name": "画面名",
      "groups": [
        {
          "medium_category": "中項目名",
          "small_category": "小項目名（無ければ空文字）",
          "cases": [
            {
              "test_no": 1,
              "verification_content": "…であること",
              "execution_steps": ["①…する", "②…を押下する", "③…を確認する"]
            }
          ]
        }
      ]
    }
  ]
}
```
- `test_no` は仮で 1.. 連番でよい（gwex が追記時に既存最大Noから振り直す）。
- 階層は「画面 → (中項目, 小項目) → ケース」。親は繰り返さない（ネストで表現）。

## スタイルガイド

- `verification_content`：**期待結果を1文**（例「…がマスキングされていること」「…画面に遷移すること」）。
- `execution_steps`：**操作手順を ①②③ の配列**で（前提・操作・確認の順）。
- 命名は既存シートの画面名/中項目名に**揃える**（手順2で取得した既存を参照）。

## 生成範囲

- 差分が触る画面/機能を**中心**に生成する。
- 加えて、明らかな**影響範囲**（共通部品・連携先API・デグレ確認観点）も補う。
- 削除された機能の既存テスト項目は**自動削除しない**（必要なら人に確認）。

## 注意

- シート書き込みは破壊的操作。**必ずレビュー承認後**に実行する。
- Google スプレッドシートは事前に書込みスコープ認証済みであること（`gwex auth`）。
- ネイティブ Google シートの既存 testspec 抽出（手順2）は未対応の場合がある。その場合は Excel か Drive アップロード xlsx を対象にする。
