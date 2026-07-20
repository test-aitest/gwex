# gwex — Google Workspace + Excel ドキュメント変換／読み書きツール

`gwex`（**G**oogle **W**orkspace + **Ex**cel）は、Google Workspace（Docs / Sheets / Slides）と
Microsoft Office（docx / xlsx / pptx）を **Markdown / typed JSON** に変換し、さらに
**Excel / Google スプレッドシートへ typed に書き戻す** CLI／MCP ツールです。

「設計書・テスト仕様をプログラムから安全に読み書きする」ことに特化しており、
**Excel でも Google スプレッドシートでも同じコマンド**で動くのが最大の特徴です。

### 3つの設計の柱

- **取得(I/O) と 解析(pure) をプロセス分離** — 解析を `sandbox-exec`（network / ファイル書き込み拒否）で
  隔離し、悪意あるファイルを開いても外部通信や改竄ができないように、安全性を運用でなく**アーキテクチャ**で担保。
- **正規化 Document IR**（pydantic v2）を中間に挟み、Markdown と typed JSON を**同一の中間表現**から出力。
- **テスト仕様（結合テスト項目）の typed 読み書きブリッジ** — 階層抽出・追記（冪等／体裁整形）・
  画面イメージの before/after 画像添付を、Excel と Google で同一コマンドで実行。

> [!IMPORTANT]
> **gwex 自身は LLM を持ちません。** 生成・判断（テスト項目を考える、差分を解釈する等）は
> 呼び出し側（Claude などのエージェント）が行い、gwex は typed な読み書きと配置だけを担う「手」です。
> エージェントと組み合わせる使い方は [スキル（ユースケース）](#スキルユースケース) を参照してください。

---

## 目次

1. [このツールが解決すること](#このツールが解決すること)
2. [対応フォーマット](#対応フォーマット)
3. [クイックスタート](#クイックスタート)
4. [導入手順（詳細）](#導入手順詳細)
5. [CLI コマンドリファレンス](#cli-コマンドリファレンス)
6. [スキル（ユースケース）](#スキルユースケース)
7. [MCP サーバ](#mcp-サーバ)
8. [テスト仕様（testspec）と列マッピング](#テスト仕様testspecと列マッピング)
9. [アーキテクチャ](#アーキテクチャ)
10. [開発](#開発)
11. [セキュリティ](#セキュリティ)

---

## このツールが解決すること

設計書やテスト仕様を Excel / Google スプレッドシートで管理しているチームでは、こんな手間が発生します。

- 実装の差分を見ながら**テスト項目を手で書き起こし**、シートの体裁（罫線・採番・親項目の縦結合）を崩さないよう追記する。
- 画面の修正前/修正後スクリーンショットを、シートの所定セルに**サイズを整えて貼り付ける**。
- Excel と Google スプレッドシートで**操作方法が違う**ため、どちらを使うかで手順が変わる。

`gwex` はこれらを **typed なコマンド** に落とし込みます。LLM（エージェント）が「中身」を考え、
`gwex` が「シートへの正確な配置」を担うことで、**人間のレビューを挟みつつ半自動化**できます。

---

## 対応フォーマット

| 入力（変換 / 抽出） | 書き戻し |
|---|---|
| Google Docs / Sheets / Slides（ネイティブ API） | ローカル `.xlsx` |
| Microsoft `.docx` / `.xlsx` / `.pptx` | Google スプレッドシート（ネイティブ） |
| Drive にアップロードされた Office ファイル | |

- 変換出力: **Markdown** または **typed JSON**（同一 IR から生成）。
- 書き戻し: セルテキスト / TestSpec 構造 / 画像（Excel = 埋め込み、Google = Apps Script オーバーレイ）/ 行・書式操作。

---

## クイックスタート

```bash
# 1. インストール
git clone https://github.com/test-aitest/gwex.git ~/lesson/gwex && cd ~/lesson/gwex
uv sync
uv tool install --editable .

# 2. ローカル .xlsx ならこれだけで動く（Google 連携は不要）
gwex convert ./report.xlsx --to json -o report.json
gwex testspec ./design.xlsx --sheet "3.結合テスト項目" --to md

# 3. Google 連携を使う場合は OAuth 設定後に認証（下の「導入手順」参照）
gwex auth
gwex testspec "https://docs.google.com/spreadsheets/d/XXXX/edit" --sheet "3.結合テスト項目"
```

> ローカル `.xlsx` だけを扱うなら **Google 連携の設定は一切不要**です。Google URL を渡したときだけ認証が必要になります。

---

## 導入手順（詳細）

### 1. 前提

- macOS / Linux、**Python 3.12+**、[uv](https://docs.astral.sh/uv/)
- Google 連携を使う場合のみ: 自分の **GCP プロジェクト**（OAuth Desktop クライアント）

> sandbox 隔離（解析の confine）は **macOS の `sandbox-exec` でのみ有効**です。Linux では隔離が無効のまま
> 実行されます（信頼できる入力のみを扱う前提）。詳細は [セキュリティ](#セキュリティ) を参照。

### 2. インストール

```bash
git clone https://github.com/test-aitest/gwex.git ~/lesson/gwex && cd ~/lesson/gwex
uv sync                      # 依存を解決（dev も入れるなら --group dev）
uv tool install --editable . # `gwex` / `gwex-mcp` を PATH に導入
gwex --help
```

### 3. Google OAuth の設定（Google 連携を使う場合のみ）

1. GCP で **Drive / Docs / Sheets / Slides API** を有効化。
2. **OAuth 2.0 クライアント（デスクトップ）** を作成し、クライアント ID / シークレットを取得。
   - 同意画面が「テスト」状態なら、自分を **テストユーザー**に追加。
3. リポジトリ直下に `.env` を作成（値は自分のもの。`.env` は `.gitignore` 済みでコミットされません）:
   ```dotenv
   GWEX_CLIENT_ID="xxxxxxxx.apps.googleusercontent.com"
   GWEX_CLIENT_SECRET_VALUE="GOCSPX-xxxxxxxx"
   # 任意: GWEX_CLIENT_SECRET=/path/to/client_secret.json （JSON ファイルで渡す場合）
   # 任意: GWEX_CONFIG_DIR=~/.config/gwex （token 保存先の上書き）
   ```
4. 認証（ブラウザが開く。token は `~/.config/gwex/token.json` にキャッシュ）:
   ```bash
   gwex auth
   ```
   付与スコープ（全5つ）:
   `drive.readonly` / **`drive.file`（共有・アップロード）** / `documents.readonly` /
   **`spreadsheets`（読み書き）** / `presentations.readonly`。

> **組織の Google Workspace 配下の場合**：管理者が「未承認のサードパーティアプリ」をブロックしていると
> `gwex auth` が弾かれます。GCP プロジェクトの作成可否と、発行した OAuth クライアント ID を
> 「信頼済みアプリ」に登録してもらえるか、管理者に確認してください。

### 4. Google スプレッドシートへの画像挿入を使う場合（Apps Script）

Sheets REST API にはセル内画像挿入が無いため、Apps Script Web アプリを経由します（公開リンク不要・PII 露出なし）。

1. <https://script.google.com> で新規プロジェクト → `src/gwex/writer/appscript/gwex_image.gs` を貼付。
2. 「デプロイ」→「ウェブアプリ」（実行 = 自分 / アクセス = 自分のみ）。
3. 発行された `/exec` URL を `.env` に登録:
   ```dotenv
   GWEX_APPSCRIPT_URL="https://script.google.com/macros/s/XXXX/exec"
   ```
   詳細は `src/gwex/writer/appscript/README.md`。

---

## CLI コマンドリファレンス

`<source>` / `<target>` は **ローカルパス** か **Google URL** のどちらでも可（自動判別）。
一部の書式系コマンド（後述）は **`.xlsx` 専用**です。`-o, --output` を省略すると、多くの書き込み系は
**in-place（その場で上書き）** になります。

全コマンドの概要:

| コマンド | 用途 | 対象 |
|---|---|---|
| [`convert`](#変換抽出) | 文書 → Markdown / JSON | 両対応 |
| [`testspec`](#変換抽出) | テスト仕様を階層抽出 | 両対応 |
| [`extract-sheet`](#変換抽出) | 指定シートだけを抽出 | 両対応 |
| [`auth`](#認証) | OAuth 認証 | Google |
| [`write-testspec`](#テスト仕様の書き込み) | TestSpec をシートに記入 | 両対応 |
| [`update-case`](#テスト仕様の書き込み) | 既存ケース行を更新 | 両対応 |
| [`set-text`](#セル操作) | セルにテキスト | 両対応 |
| [`set-image`](#セル操作) | 画像をセルに配置 | 両対応 |
| [`annotate-image`](#セル操作) | 画像上に赤枠+番号図形を注入 | xlsx 専用 |
| [`clear-annotations`](#セル操作) | 注釈図形を除去 | xlsx 専用 |
| [`draw-flow`](#セル操作) | 画面遷移図を生成（スクショ+矢印） | xlsx 専用 |
| [`flow-spec`](#セル操作) | スクリーングラフから遷移図 spec を生成 | YAML |
| [`flow-capture`](#セル操作) | Appium 実走でスクショ+ボタン矩形を収集 | iOS |
| [`set-section`](#セル操作) | before/after セクション作成 | 両対応 |
| [`insert-rows`](#行書式操作) | 行を挿入・シフト | 両対応 |
| [`delete-rows`](#行書式操作) | 行を削除・シフト | 両対応 |
| [`copy-row-format`](#行書式操作) | 行の書式をコピー | 両対応 |
| [`strip-instruction-row`](#行書式操作) | テンプレの指示行を除去 | 両対応 |
| [`autofit-rows`](#行書式操作) | 行高を自動調整 | Google |
| [`set-row-height`](#行書式操作) | 行高を設定 | xlsx 専用 |
| [`set-alignment`](#行書式操作) | セルの揃え方を設定 | xlsx 専用 |
| [`set-merge`](#行書式操作) | セルを結合 | xlsx 専用 |
| [`clear-images`](#品質チェッククリーンアップ) | シートの画像を削除 | 両対応 |
| [`lint`](#品質チェッククリーンアップ) | 設計書の残骸を検出 | 両対応 |
| [`diff`](#品質チェッククリーンアップ) | 2つの .xlsx を比較 | xlsx |
| [`share`](#drive-連携) | Drive にアップロードし共有 URL を返す | ローカル → Google |
| [`rename`](#drive-連携) | Drive 上のファイル名を変更 | Google |
| [`export-pdf`](#drive-連携) | スプレッドシートを PDF / PNG 出力 | Google |

---

### 変換・抽出

#### `gwex convert <source>` — 文書 → Markdown / JSON
| オプション | 説明 | 既定 |
|---|---|---|
| `--to md\|json` | 出力形式 | `md` |
| `-o, --output <path>` | 出力先（省略時は標準出力） | — |
| `--no-sandbox` | 解析を sandbox 隔離せず実行（開発用） | off |
```bash
gwex convert "https://docs.google.com/document/d/XXXX/edit" --to md
gwex convert ./report.xlsx --to json -o report.json
```

#### `gwex testspec <source> --sheet <名>` — テスト仕様の階層抽出
xlsx / Drive アップロード xlsx / ネイティブ Google スプレッドシートから、画面→中項目→小項目→ケースの階層を抽出。
| オプション | 説明 | 既定 |
|---|---|---|
| `--sheet <名>` *(必須)* | 対象シート（例 `3.結合テスト項目`） | — |
| `--to md\|json` | 出力形式 | `md` |
| `--mapping <toml>` | 列マッピング | 既定テンプレ |
| `-o, --output <path>` | 出力先 | — |
```bash
gwex testspec ./design.xlsx --sheet "3.結合テスト項目" --to json -o spec.json
```

#### `gwex extract-sheet <target> --sheet <名> -o <path>` — 指定シートのみ抽出
ブック全体から1シートだけを取り出して新規ファイルに保存。

---

### 認証

#### `gwex auth` — OAuth 認証
不足スコープがあればブラウザで再同意。token を `~/.config/gwex/token.json` に保存。

---

### テスト仕様の書き込み

#### `gwex write-testspec <target> --sheet <名> --json <file>` — TestSpec をシートに記入
| オプション | 説明 |
|---|---|
| `--sheet <名>` *(必須)* | 書き込み先シート |
| `--json <file>` *(必須)* | TestSpec JSON |
| `--append` | 既存データの下に追記（採番は既存グループ最大 No から継続） |
| `--dedup` | 既存と完全一致するケースを除外（冪等） |
| `--format` | 追記ブロックに既存体裁（罫線 B〜S / 親縦結合 / 番号）を付与 |
| `--clear-rows` | 書き込み前に対象範囲をクリア |
| `--start-row <N>` | 書き込み開始行を明示 |
| `--mapping <toml>` / `-o, --output` | 列マッピング / 出力先 |
```bash
gwex write-testspec ./design.xlsx --sheet "3.結合テスト項目" --json new.json --append --dedup --format
```

#### `gwex update-case <target> --sheet <名> --row <N>` — 既存ケースの更新
| オプション | 説明 |
|---|---|
| `--row <N>` *(必須)* | 更新対象行（抽出時の `source_row`） |
| `--verification <文>` *(必須)* | 確認内容 |
| `--steps <改行区切り>` | 実施手順 |
| `--mapping <toml>` | 列マッピング |
```bash
gwex update-case ./design.xlsx --sheet "3.結合テスト項目" --row 42 \
  --verification "ログイン後にダッシュボードへ遷移すること" --steps $'①…\n②…'
```

---

### セル操作

#### `gwex add-sheet <xlsx> --sheet <名>` — 空シートを追加
`--index N` で挿入位置（0始まり・省略時は末尾）、`--exist-ok` で同名シートがあってもエラーにしない。
レビュー結果シートの追加など、既存ブックに新しいシートを足すときに使う。xlsx のみ。

```bash
gwex add-sheet ./design.xlsx --sheet "レビュー結果" --exist-ok
```

#### `gwex set-table <xlsx> --sheet <名> --start-cell <C> --json <rows.json>` — 表を一括書き込み
2次元配列（JSON）を start-cell を左上として書き込む。先頭行は見出し（太字＋薄青塗り）、全セルに罫線と
折り返しが付く。`--no-header` で見出し扱いを外す。**`set-text` を何十回も叩く代わりに使う**（xlsx のみ）。

```bash
echo '[["No","指摘","重大度"],["1","フォントサイズ不一致","高"]]' > rows.json
gwex set-table ./design.xlsx --sheet "レビュー結果" --start-cell B10 --json rows.json
```

#### `gwex set-text <target> --sheet <名> --cell <C> --text <文字>` — セルにテキスト
`--type` で値の型（文字列／数値等）を指定可。
```bash
gwex set-text ./design.xlsx --sheet "3.結合テスト項目" --cell B2 --text "確認済み"
```

#### `gwex set-image <target> --sheet <名> --cell <C> --image <png>` — 画像を配置
| オプション | 説明 | 既定 |
|---|---|---|
| `--cell <C>` *(必須)* | アンカーセル（`--range` 時は左上） | — |
| `--image <path>` *(必須)* | 画像ファイル | — |
| `--range <C7:F20>` | 画像枠の結合範囲（結合 + オーバーレイ） | — |
| `--insert-rows` | 枠ぶんの行を挿入して確実に空枠を作る | off |
| `--width / --height <px>` | 表示サイズ | — |
| `--max-dim <px>` | 長辺上限（超過時に自動縮小） | — |
| `--scale <倍率>` | Google: 枠フィット後の倍率（`0.8` = 80%・中央寄せ） | `1.0` |
```bash
gwex set-image ./design.xlsx --sheet "2.画面イメージ(iOS)" --cell C7 --range C7:F25 --image cap.png --max-dim 1024
```

#### `gwex annotate-image <xlsx> --sheet <名> --rect "x,y,w,h[:label]"` — 画像上に赤枠+番号を注入
埋め込み画像の上へ、変更箇所を示す**赤枠矩形と番号バッジをネイティブ図形**（DrawingML `xdr:sp`）として注入する（**xlsx 専用**。Google スプレッドシートには図形 API が無い）。結合テスト項目から「①が変更点」と番号参照する用途。図形は Excel 上で移動・削除できる。
| オプション | 説明 | 既定 |
|---|---|---|
| `--rect "x,y,w,h[:label]"` *(必須・複数可)* | 赤枠の位置と寸法（**元画像のピクセル座標**）。label 省略時は 1,2,… の連番 | — |
| `--name <画像名>` | 対象画像（例: `capture2`）。シートに1枚なら省略可 | — |
| `--color <RRGGBB>` | 枠・バッジの色 | `FF0000` |
| `--line-pt <pt>` | 枠線の太さ | `2.25` |
| `--no-badge` | 番号バッジを付けない | badge 有り |
```bash
gwex annotate-image ./design.xlsx --sheet "2.画面イメージ(iOS)" --name capture2 --rect "100,508,128,66:1"
```
注意: openpyxl 系コマンド（`set-text` 等）で保存すると**そのシートの図形が消える**ことがある。
テンプレ由来の図形・画像は `xlsx_zip.restore_drawings()` が自動で復元するが（2026-07-14 追加）、
**同じシートに画像がある場合、後から openpyxl 系を流すと annotate-image の赤枠だけ落ちる**（画像は openpyxl が引き継ぐため復元対象外になる）。
→ 注釈はこれまで通り**ワークフローの最終段**で実行する（消えたら再実行すれば同位置に戻る）。

#### `gwex clear-annotations <xlsx> --sheet <名>` — 注釈図形を除去
`annotate-image` が注入した赤枠・番号バッジを全て除去する（画像は残す）。`-o` 省略で in-place。

#### `gwex draw-flow <出力.xlsx> --spec <flow.yaml>` — 画面遷移図を生成
スクリーンショットを格子配置（列=遷移ステップ、行=同一画面のバリエーション）し、
トリガーボタンの赤枠・矢印（直線/カギ線）・条件ラベル・処理ボックス付きの遷移図 xlsx を**新規生成**する（**xlsx 専用**）。
spec の形式は `gwex.writer.xlsx_flow` の docstring を参照（nodes: 画面ID/画面名/画像/col/row、edges: from/to/trigger_rect/label/via/dash）。
モーダル等への `detour` エッジは既定で**点線**（`dash: true/false` で任意のエッジを明示指定可）。
画像パスは spec ファイルのディレクトリからの相対で解決。
```bash
gwex draw-flow ./振込フロー.xlsx --spec flow.yaml
```
注意: 生成した遷移図は矢印・赤枠が図形なので、openpyxl 系コマンドで再保存すると落ちうる。編集せず成果物として扱う（変更は spec を直して再生成）。

#### `gwex flow-spec <graph.yaml> --from <画面> --to <画面>` — 遷移図 spec の雛形を生成
UI クローラーが構築したスクリーングラフ（`screens/components/navigators` 形式、例: devpilot-graph.yaml）から
最短経路を BFS で抽出し、`draw-flow` 用 spec の雛形 YAML を出力する。
矢印ラベルはトリガーコンポーネントの表示ラベルから自動解決。認証情報（requiredInputs）は転記しない。
| オプション | 説明 | 既定 |
|---|---|---|
| `--from / --to <画面名>` *(必須)* | 経路の起点/終点 | — |
| `--via <画面名>` | 経由地（複数可・順序どおり） | — |
| `--sheet <名>` | シート名 | 経路から自動 |
| `--image-dir <dir>` | スクショ配置予定ディレクトリ | `caps` |
```bash
gwex flow-spec devpilot-graph.yaml --from LoginViewController --to TransferModalViewController -o flow.yaml
# → caps/ にスクショを置き、必要なら trigger_rect を追記して: gwex draw-flow 遷移図.xlsx --spec flow.yaml
```

#### `gwex flow-capture <flow.yaml> --graph <graph.yaml> --udid <UDID> --bundle-id <ID>` — 実走収集
Appium でシミュレータ/デバイス上のアプリを spec の経路どおりに操作し、
各画面のスクリーンショットと trigger_component の**要素矩形（画像ピクセル座標）を自動計測**して spec を完成させる。
要: Appium サーバ稼働（`--appium`、既定 `http://127.0.0.1:4723`）・対象デバイス起動・アプリ導入済み。
| オプション | 説明 | 既定 |
|---|---|---|
| `--dismiss <要素ID>` | 邪魔要素（初回設定画面の完了ボタン等）を閉じる追加候補（複数可）。**末尾 `*` で前方一致**（例: `'次へ*'` は選択数つき「次へ (1)」にもマッチ） | 「閉じる」等 |
| `--settle <秒>` | 遷移後の待ち | `2.0` |
| `--draw <出力.xlsx>` | 続けて draw-flow で遷移図まで生成 | — |
| `-o <yaml>` | 完成 spec の出力先 | `<spec>_captured.yaml` |
spec の edge 拡張: `pre: ["＋1,000"]`（トリガーが入力するまで無効なケースで、先にタップする要素）。
遷移が観測できるまで「オーバーレイ掃除→再タップ」するため、促進モーダル/コーチマークに強い。
graph の requiredInputs（認証情報）は入力にのみ使い、spec へは書き込まない。

#### `gwex set-section <target> --sheet <名> --top-row <N> --title <題>` — before/after セクション
見出し + 修正前/修正後ラベル + 青背景 + 箱罫線を作り、画像を**枠の80%・比率維持・中央配置**。**Excel / Google 両対応**。
| オプション | 説明 | 既定 |
|---|---|---|
| `--top-row <N>` *(必須)* | セクション見出しの行 | — |
| `--title <題>` *(必須)* | 見出し（画面名など） | — |
| `--before / --after <png>` | 修正前 / 修正後 画像 | — |
| `--cols <C,L>` | セクション左右端の列 | `C,L` |
| `--split <H>` | 修正後の開始列 | `H` |
| `--scale <倍率>` | 画像を枠の何倍にするか | `0.8` |
| `--n-rows <N>` | セクションの行数 | — |
```bash
gwex set-section ./design.xlsx --sheet "2.画面イメージ(iOS)" --top-row 201 \
  --title "ダッシュボード（情報行 追加）" --before before.png --after after.png
```

---

### 行・書式操作

| コマンド | 説明 |
|---|---|
| `gwex insert-rows <target> --sheet <名> --at <N>` | `N` 行目に行を挿入してシフト。`--count`（既定 1）/ `--template-row` で雛形書式を継承 |
| `gwex delete-rows <target> --sheet <名> --start <N> --count <C>` | `N` から `C` 行を削除してシフト |
| `gwex copy-row-format <src> <dst> --src-sheet <名> --dst-sheet <名> --src-rows <範囲> --dst-rows <範囲>` | 行の書式をコピー |
| `gwex strip-instruction-row <target>` | テンプレの「指示行」を削除して詰める（`--sheet` 省略時は全シート自動） |
| `gwex autofit-rows <target> --sheet <名> --start-row <N> --end-row <M>` | Google スプレッドシートの行高を内容に合わせ自動調整 |
| `gwex set-row-height <xlsx> --sheet <名> --row <N> --height <px>` | 行高を設定（**xlsx 専用**） |
| `gwex set-alignment <xlsx> --sheet <名> --cell <C> [--cell ...]` | セルの揃え（`--horizontal` / `--vertical` / `--wrap-text`）を設定（**xlsx 専用**） |
| `gwex set-merge <xlsx> --sheet <名> --range <C7:F9> [--range ...]` | セルを結合（**xlsx 専用**） |

---

### 品質チェック・クリーンアップ

#### `gwex lint <target>` — 設計書の残骸を検出
テンプレのサンプル文字列や記入漏れプレースホルダ（消し忘れの「○○を入力」等）を検出。
`--sheet`（複数可）で対象を限定、`--ignore`（複数可）で除外語を指定、`--to summary|json`、`-o`。

#### `gwex diff <a> <b>` — 2つの .xlsx を比較
セル単位の差分を出力。`--sheet` / `--to summary|json` / `--max-row` / `--max-col` / `-o`。

#### `gwex clear-images <target> --sheet <名>` — シートの画像を削除
貼り直し前のクリーンアップに。`-o` 省略で in-place。

---

### Drive 連携

#### `gwex share <path>` — Drive にアップロードし共有 URL を返す
ローカルの Office / PDF を Drive にアップロードし、共有 URL（webViewLink）を返す。
| オプション | 説明 | 既定 |
|---|---|---|
| `--convert` | Google ネイティブ形式（スプレッドシート等）に変換してアップロード | off（元形式のまま） |
| `--private` | anyone-with-link 共有をしない | off（anyone reader で共有） |
| `--name <名>` | Drive 上の表示名（中間ファイル名のままになる事故を防ぐため正式名を渡す） | — |
```bash
gwex share ./design.xlsx --convert --name "結合テスト仕様書 v1.2"
```

> [!WARNING]
> 既定では **anyone-with-link（リンクを知る全員が閲覧可）** で共有します。組織が外部共有を禁止している場合や
> 社外に出したくない場合は `--private` を付けてください。

#### `gwex rename <target> --name <新名称>` — Drive 上のファイル名を変更

#### `gwex export-pdf <target> --sheet <名>` — スプレッドシートを PDF / PNG 出力
`--sheet`（複数可）で対象シート、`--png` で PNG 出力、`--landscape` で横向き、`-o` で出力先（既定 `/tmp`）。

---

## スキル（ユースケース）

`gwex` の真価は、Claude Code のエージェントと組み合わせたときに発揮されます。
`skills/` に **Claude Code スキルの正本**を同梱しており、利用側プロジェクト（テスト対象アプリのリポジトリ）の
`.claude/skills/` に配置すると `/generate-testspec`・`/capture-screen` として呼べます。

| スキル | ユースケース |
|---|---|
| **`generate-testspec`** | PR / ブランチの差分をエージェントが読み、結合テスト項目（確認内容＋実施手順）を生成。既存シートと突き合わせて「新規 / 更新 / skip」を分類し、**人間のレビュー承認後**に `gwex` でシートへ追記・更新する。 |
| **`capture-screen`** | 起動中の iOS シミュレータ / Android エミュレータの画面をキャプチャし、`gwex set-image` で「画面イメージ」シートの所定セルに修正前/後ラベル付きで添付する。 |

配置例:
```bash
mkdir -p .claude/skills
cp -R ~/lesson/gwex/skills/generate-testspec .claude/skills/
cp -R ~/lesson/gwex/skills/capture-screen   .claude/skills/
```

### ベストプラクティス（推奨ワークフロー）

1. **抽出 → 把握**：`gwex testspec ... --to json` で既存仕様を取得し、命名規則・既存ケースを把握する。
2. **生成は LLM、配置は gwex**：エージェントがテスト項目を生成し、人間がレビュー。承認後にのみ書き込む
   （シート書き込みは破壊的操作）。
3. **追記は冪等に**：`write-testspec --append --dedup --format` を使えば、再実行しても重複せず体裁も揃う。
4. **更新は `update-case`**：既存ケースの修正は `source_row` を使ってピンポイントに。全書き換えを避ける。
5. **Google 画像は小さく**：セル内画像は base64 でサイズ上限があるため、`--max-dim 1024` 程度に縮小して添付する。
6. **共有はスコープに注意**：社外秘は `gwex share --private`。表示名は `--name` で正式名にする。

---

## MCP サーバ

`gwex-mcp` で MCP（FastMCP）サーバとして起動。提供ツール（全6種）:

| ツール | 対応 CLI | 説明 |
|---|---|---|
| `convert` | `convert` | 文書を MD / JSON に変換 |
| `extract_testspec` | `testspec` | テスト仕様を JSON 抽出 |
| `write_testspec` | `write-testspec` | TestSpec をシートに記入（`append` / `dedup` 対応） |
| `update_case` | `update-case` | 既存ケース行を更新 |
| `set_cell_text` | `set-text` | セルにテキスト書き込み |
| `set_cell_image` | `set-image` | 画像をセルに配置 |

Claude Code 等の `.mcp.json` 例:
```json
{ "mcpServers": { "gwex": { "command": "gwex-mcp" } } }
```

---

## テスト仕様（testspec）と列マッピング

階層モデル: **TestSpec → Screen（画面名）→ Group（中項目・小項目）→ Case（確認内容 / 実施手順）**。

既定マッピング（実テンプレ「3.結合テスト項目」準拠）: `header_row=7` / `data_start_row=10`、列は
画面名 `C` / 中項目 `E` / 小項目 `G` / **確認内容 `H` / 実施手順 `I`**。
実テンプレはケースごとの No 専用列を持たないため、`Case.test_no` は既定マッピングでは出力されません
（採番が必要なテンプレでは `--mapping` で `test_no` 列を割り当てると追記時に自動継続されます）。
体裁整形では親 No 列 `B`(画面) / `D`(中項目) / `F`(小項目)、箱の右端 `S` も使用。
別テンプレに合わせるときは `--mapping <toml>` で上書きします。

TestSpec JSON の形（`write-testspec --json` に渡す形）:
```json
{
  "sheet": "3.結合テスト項目",
  "screens": [
    {
      "screen_name": "ログイン画面",
      "groups": [
        {
          "medium_category": "認証",
          "small_category": "正常系",
          "cases": [
            {
              "test_no": 1,
              "verification_content": "正しい資格情報でダッシュボードへ遷移すること",
              "execution_steps": ["①ID/PWを入力する", "②ログインを押下する", "③遷移先を確認する"]
            }
          ]
        }
      ]
    }
  ]
}
```
`test_no` は任意です（省略可）。`test_no` 列を持つマッピングでは仮の連番で構いません
（`--append` 時に既存グループの最大 No から自動で振り直されます）。既定マッピングでは出力されません。

---

## アーキテクチャ

```
source(URL/path) ─▶ fetcher(privileged: I/O)            ─▶ raw(JSON/bytes)
                                                            │
                          confined parser (sandbox-exec)  ◀─┘  network / fs-write 拒否
                                   │
                              Document IR (pydantic v2)
                              ├─▶ serializers ─▶ Markdown / typed JSON
                              └─▶ domains(testspec) ─▶ writer ─▶ xlsx / Google
```

- **pure 層**（`gwex.ir` / `parsers` / `serializers` / `domains`）は I/O・fetcher を import 不可。
  `import-linter` が CI で機械的に強制（解析層に発信能力が無いことを保証）。
- **privileged 層**（`fetcher` / `writer` / `sandbox`）のみ I/O を行う。
- Google は **戦略B**（Drive export を経由せず、Docs / Sheets / Slides API の構造化 JSON を直接読む）。
  testspec 抽出のみ Drive で xlsx にエクスポートして openpyxl で読む。
- ディレクトリ: `src/gwex/{ir,parsers,serializers,domains,fetcher,writer,surfaces,sandbox}`。

---

## 開発

```bash
uv run pytest            # 単体テスト
uv run lint-imports      # pure 層の依存契約を検証（Contracts: 1 kept, 0 broken）
```
- 解析は既定で sandbox 経由（`convert --no-sandbox` で無効化可）。
- CLI / MCP のエントリポイントは `src/gwex/surfaces/`（`cli.py` / `mcp.py`）。

---

## セキュリティ

- **解析の confine**: 信頼できないドキュメントの解析を、network / ファイル書き込みを奪った別プロセス
  （macOS `sandbox-exec`）で実行します。これにより、悪意あるファイルを開いても SSRF や改竄ができません。
  **この隔離は現状 macOS のみ有効**で、Linux では隔離されずに実行されます。信頼できない入力を Linux で
  扱う場合は注意してください。
- **認証情報**: OAuth クライアントの値は `.env`（`.gitignore` 済み）に置き、token は `~/.config/gwex/token.json`
  にキャッシュされます。どちらもリポジトリにはコミットされません。秘密情報をコミットしないよう注意してください。
- **共有**: `gwex share` は既定で anyone-with-link 共有です。社外秘は必ず `--private` を付けてください。

---

## ライセンス

[MIT License](LICENSE) © 2026 Tatsuki Yabe
