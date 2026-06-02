# gwex — Google Workspace + Excel ドキュメント変換／読み書きツール

`gwex`（**G**oogle **W**orkspace + **Ex**cel）は、Google Workspace（Docs / Sheets / Slides）と
Microsoft Office（docx / xlsx / pptx）を **Markdown / typed JSON** に変換し、さらに **Excel /
Google スプレッドシートへ typed に書き戻す**ツールです。次の3点を設計の柱としています。

- **取得(I/O) と 解析(pure) をプロセス分離**し、解析を `sandbox-exec`（network / fs-write 拒否）で confine。安全性を運用でなくアーキテクチャで担保。
- **正規化 Document IR**（pydantic v2）を中間に挟み、Markdown と typed JSON を同一 IR から出力。
- **テスト仕様（結合テスト項目）の typed 読み書きブリッジ**：階層抽出・追記（冪等 / 体裁整形）・画面イメージの before/after 画像添付を、**Excel と Google で同一コマンド**で実行。

> gwex は **LLM を持ちません**。生成・判断は呼び出し側（エージェント等）が行い、gwex は
> typed な読み書きと配置だけを担います。

---

## 目次
1. [対応フォーマット](#対応フォーマット)
2. [導入手順](#導入手順)
3. [CLI コマンドリファレンス](#cli-コマンドリファレンス)
4. [MCP サーバ](#mcp-サーバ)
5. [テスト仕様（testspec）と列マッピング](#テスト仕様testspecと列マッピング)
6. [アーキテクチャ](#アーキテクチャ)
7. [開発](#開発)

---

## 対応フォーマット

| 入力（変換 / 抽出） | 書き戻し |
|---|---|
| Google Docs / Sheets / Slides（ネイティブ） | ローカル `.xlsx` |
| Microsoft `.docx` / `.xlsx` / `.pptx` | Google スプレッドシート（ネイティブ / アップロード） |
| Drive にアップロードされた Office ファイル | |

- 変換出力: **Markdown** または **typed JSON**（同一 IR から）。
- 書き戻し: セルテキスト / TestSpec 構造 / 画像（Excel=埋め込み、Google=Apps Script オーバーレイ）。

---

## 導入手順

### 1. 前提
- macOS / Linux、**Python 3.12+**、[uv](https://docs.astral.sh/uv/)
- Google 連携を使う場合: 自分の **GCP プロジェクト**（OAuth Desktop クライアント）

### 2. インストール
```bash
git clone <repo> ~/lesson/gwex && cd ~/lesson/gwex
uv sync                      # 依存を解決（dev 含むなら --group dev）
uv tool install --editable . # `gwex` / `gwex-mcp` を PATH に導入
gwex --help
```

### 3. Google OAuth の設定（Google 連携を使う場合のみ）
ローカル `.xlsx` だけ扱うならこの節は不要です。

1. GCP で **Drive / Docs / Sheets / Slides API** を有効化。
2. **OAuth 2.0 クライアント（デスクトップ）** を作成し、クライアント ID / シークレットを取得。
   - 自分を **テストユーザー**に追加（同意画面が「テスト」の場合）。
3. リポジトリ直下に `.env` を作成（値は自分のもの）:
   ```dotenv
   GWEX_CLIENT_ID="xxxxxxxx.apps.googleusercontent.com"
   GWEX_CLIENT_SECRET_VALUE="GOCSPX-xxxxxxxx"
   # 任意: GWEX_CLIENT_SECRET=/path/to/client_secret.json （JSON で渡す場合）
   # 任意: GWEX_CONFIG_DIR=~/.config/gwex （token 保存先の上書き）
   ```
4. 認証（ブラウザが開く。token は `~/.config/gwex/token.json` にキャッシュ）:
   ```bash
   gwex auth
   ```
   付与スコープ: `drive.readonly` / `documents.readonly` / `presentations.readonly` /
   **`spreadsheets`（読み書き）**。

### 4. Google スプレッドシートへの画像挿入を使う場合（Apps Script）
Sheets REST API にセル画像挿入が無いため、Apps Script Web アプリを経由します（公開リンク不要）。
1. <https://script.google.com> で新規プロジェクト → `src/gwex/writer/appscript/gwex_image.gs` を貼付。
2. 「デプロイ」→「ウェブアプリ」（実行=自分 / アクセス=自分のみ）。
3. 発行された `/exec` URL を `.env` に登録:
   ```dotenv
   GWEX_APPSCRIPT_URL="https://script.google.com/macros/s/XXXX/exec"
   ```
   詳細は `src/gwex/writer/appscript/README.md`。

---

## CLI コマンドリファレンス

`<source>` / `<target>` は **ローカルパス** か **Google URL** のどちらでも可（自動判別）。

### `gwex convert <source>` — ドキュメント → Markdown / JSON
| オプション | 説明 |
|---|---|
| `--to md\|json` | 出力形式（既定 `md`） |
| `-o, --output <path>` | 出力先（省略時は標準出力） |
| `--no-sandbox` | 解析を sandbox 隔離せず実行（開発用） |
```bash
gwex convert "https://docs.google.com/document/d/XXXX/edit" --to md
gwex convert ./report.xlsx --to json -o report.json
```

### `gwex testspec <source> --sheet <名>` — テスト仕様の階層抽出
xlsx / Drive アップロード xlsx / **ネイティブ Google スプレッドシート**から、
画面→中項目→小項目→ケースの階層を抽出。
| オプション | 説明 |
|---|---|
| `--sheet <名>` *(必須)* | 対象シート（例 `3.結合テスト項目`） |
| `--to md\|json` | 出力形式（既定 `md`） |
| `--mapping <toml>` | 列マッピング（省略時は既定テンプレ） |
| `-o, --output <path>` | 出力先 |
```bash
gwex testspec ./design.xlsx --sheet "3.結合テスト項目" --to json -o spec.json
gwex testspec "https://docs.google.com/spreadsheets/d/XXXX/edit" --sheet "3.結合テスト項目"
```

### `gwex auth` — OAuth 認証
不足スコープがあれば再同意。token を `~/.config/gwex/token.json` に保存。

### `gwex write-testspec <target> --sheet <名> --json <file>` — TestSpec をシートに記入
| オプション | 説明 |
|---|---|
| `--sheet <名>` *(必須)* | 書き込み先シート |
| `--json <file>` *(必須)* | TestSpec JSON |
| `--append` | 既存データの下に追記（採番は既存グループ最大Noから継続） |
| `--dedup` | 既存と完全一致するケースを除外（冪等） |
| `--format` | 追記ブロックに既存体裁（罫線 B〜S / 親縦結合 / 番号）を付与 |
| `--mapping <toml>` / `-o, --output` | 列マッピング / 出力先 |
```bash
gwex write-testspec ./design.xlsx --sheet "3.結合テスト項目" --json new.json --append --dedup --format
```

### `gwex update-case <target> --sheet <名> --row <N>` — 既存ケースの更新
| オプション | 説明 |
|---|---|
| `--row <N>` *(必須)* | 更新対象行（抽出時の `source_row`） |
| `--verification <文>` *(必須)* | 確認内容 |
| `--steps <改行区切り>` | 実施手順 |

### `gwex set-text <target> --sheet <名> --cell <C> --text <文字>` — セルにテキスト
```bash
gwex set-text ./design.xlsx --sheet "3.結合テスト項目" --cell B2 --text "確認済み"
```

### `gwex set-image <target> --sheet <名> --cell <C> --image <png>` — 画像を配置
| オプション | 説明 |
|---|---|
| `--cell <C>` *(必須)* | アンカーセル（`--range` 時は左上） |
| `--image <path>` *(必須)* | 画像ファイル |
| `--range <C7:F20>` | 画像枠の結合範囲（結合＋オーバーレイ） |
| `--insert-rows` | 枠ぶんの行を挿入して確実に空枠を作る |
| `--width / --height <px>` | 表示サイズ（px） |
| `--max-dim <px>` | 長辺上限（超過時に自動縮小） |
| `--scale <倍率>` | Google: 枠フィット後の倍率（`0.8`＝80%・中央寄せ） |
```bash
gwex set-image ./design.xlsx --sheet "2.画面イメージ(iOS)" --cell C7 --range C7:F25 --image cap.png --max-dim 1024
```

### `gwex set-section <target> --sheet <名> --top-row <N> --title <題>` — before/after セクション
見出し＋修正前/修正後ラベル＋青背景＋箱罫線を作り、画像を**枠の80%・比率維持・中央配置**。**Excel / Google 両対応**。
| オプション | 説明 |
|---|---|
| `--top-row <N>` *(必須)* | セクション見出しの行 |
| `--title <題>` *(必須)* | 見出し（画面名など） |
| `--before / --after <png>` | 修正前 / 修正後 画像 |
| `--cols C,L` / `--split H` | 左右端の列 / 修正後の開始列 |
| `--scale 0.8` | 画像を枠の何倍にするか（既定 0.8） |
```bash
gwex set-section ./design.xlsx --sheet "2.画面イメージ(iOS)" --top-row 201 \
  --title "ダッシュボード（情報行 追加）" --before before.png --after after.png
gwex set-section "https://docs.google.com/spreadsheets/d/XXXX/edit" --sheet "画面イメージ" \
  --top-row 2 --title "ダッシュボード" --before before.png --after after.png
```

---

## MCP サーバ

`gwex-mcp` で MCP（FastMCP）として起動。提供ツール:
`convert` / `extract_testspec` / `write_testspec` / `update_case` / `set_cell_text` / `set_cell_image`。

Claude Code 等の `.mcp.json` 例:
```json
{ "mcpServers": { "gwex": { "command": "gwex-mcp" } } }
```

---

## テスト仕様（testspec）と列マッピング

階層モデル: **TestSpec → Screen（画面名）→ Group（中項目・小項目）→ Case（No / 確認内容 / 実施手順）**。

既定マッピング（実テンプレ準拠）: `header_row=7` / `data_start_row=10`、列は
画面名 `C` / 中項目 `E` / 小項目 `G` / No `H` / 確認内容 `I` / 実施手順 `J`。
体裁整形では No 列 `B`(画面)/`D`(中項目)/`F`(小項目)、箱の右端 `S` も使用。
`--mapping <toml>` で上書き可能。

---

## アーキテクチャ

```
source(URL/path) ─▶ fetcher(privileged: I/O)            ─▶ raw(JSON/bytes)
                                                            │
                          confined parser (sandbox-exec)  ◀─┘  network/fs-write 拒否
                                   │
                              Document IR (pydantic v2)
                              ├─▶ serializers ─▶ Markdown / typed JSON
                              └─▶ domains(testspec) ─▶ writer ─▶ xlsx / Google
```
- **pure 層**（`gwex.ir` / `parsers` / `serializers` / `domains`）は I/O・fetcher を import 不可。`import-linter` が CI で強制。
- **privileged 層**（`fetcher` / `writer` / `sandbox`）のみ I/O。
- Google は戦略B（Drive export 経由でなく Docs/Sheets/Slides API の構造化 JSON を直接）。testspec 抽出のみ Drive で xlsx にエクスポートして openpyxl で読む。

---

## 開発
```bash
uv run pytest            # 単体テスト
uv run lint-imports      # pure 層の依存契約を検証（Contracts: 1 kept, 0 broken）
```
- 解析は既定で sandbox 経由（`convert --no-sandbox` で無効化可）。
- ディレクトリ: `src/gwex/{ir,parsers,serializers,domains,fetcher,writer,surfaces,sandbox}`。
