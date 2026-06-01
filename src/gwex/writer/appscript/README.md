# gwex 画像挿入 Apps Script Web アプリ（セットアップ）

Google Sheets のセルに**ネイティブのセル内画像**を base64 で埋め込むための Web アプリ。
Sheets REST API には画像挿入が無いため Apps Script を経由する（公開リンク不要・PII 露出なし）。

## デプロイ手順（一度だけ）

1. <https://script.google.com> で新規プロジェクト作成。
2. `gwex_image.gs` の内容を貼り付けて保存。
3. （推奨）プロジェクトを **同じ GCP プロジェクト**に紐付け：
   プロジェクト設定 → 「Google Cloud Platform（GCP）プロジェクト」を gwex と同じ番号に変更。
4. 「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
   - 次のユーザーとして実行：**自分**
   - アクセスできるユーザー：**自分のみ**
5. 発行された **ウェブアプリ URL（/exec）** を控える。
6. gwex 側に登録（`.env` か環境変数）：
   ```
   GWEX_APPSCRIPT_URL="https://script.google.com/macros/s/XXXX/exec"
   ```

## 呼び出し

gwex は OAuth Bearer トークン付きで POST する。
ボディ: `{spreadsheetId, sheet, cell, dataUrl, width?, height?}`。

> 「自分のみ」デプロイの Web アプリを Bearer トークンで叩く場合、トークンのスコープ次第で
> 認可される。もし 401/403 になる場合は、アクセスを「リンクを知っている全員」にして
> URL を秘匿運用する方法もある（その場合 URL 自体がシークレット）。
