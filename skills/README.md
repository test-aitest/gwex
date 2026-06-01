# gwex Claude Code スキル

`gwex` を使う Claude Code スキルの正本。**利用側プロジェクト（テスト対象アプリのリポジトリ）の
`.claude/skills/` に配置**して有効化する（プロジェクトスキル）。

| スキル | 用途 |
|---|---|
| `generate-testspec` | PR/差分 → 結合テスト項目を生成し、レビュー後にシートへ追記/更新（A） |
| `capture-screen` | シミュレータ/エミュレータ画面をキャプチャしシートの所定セルへ添付（B） |

## 前提

- `gwex` が PATH にあること（`uv tool install --editable ~/lesson/gwex`）。
- Google 書き込み時は `gwex auth` で書込みスコープ認証済み。セル内画像は `GWEX_APPSCRIPT_URL` 設定済み。

## 配置（利用側プロジェクトで）

```bash
mkdir -p .claude/skills
cp -R ~/lesson/gwex/skills/generate-testspec .claude/skills/
cp -R ~/lesson/gwex/skills/capture-screen   .claude/skills/
# または symlink:
# ln -s ~/lesson/gwex/skills/generate-testspec .claude/skills/generate-testspec
```

配置後、Claude Code で `/generate-testspec` / `/capture-screen` として呼べる。
