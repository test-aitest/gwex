"""gwex の CLI。

    gwex convert <source> --to md|json
"""

from __future__ import annotations

from typing import Optional

import typer

from gwex import core

app = typer.Typer(add_completion=False, help="文書を Markdown / typed JSON に変換する")


@app.callback()
def _root() -> None:
    """gwex: Google Workspace + Microsoft 文書 → Markdown / typed JSON。"""


@app.command()
def convert(
    source: str = typer.Argument(..., help="Google URL またはローカル OOXML パス"),
    to: str = typer.Option("md", "--to", help="出力形式: md | json"),
    output: Optional[str] = typer.Option(None, "--output", "-o", help="出力先ファイル（省略時は標準出力）"),
    no_sandbox: bool = typer.Option(False, "--no-sandbox", help="解析を sandbox 隔離せず実行（開発用）"),
) -> None:
    result = core.convert(source, to=to, use_sandbox=not no_sandbox)
    if output:
        with open(output, "w", encoding="utf-8") as f:
            f.write(result)
        typer.echo(f"書き出しました: {output}")
    else:
        typer.echo(result)


@app.command()
def testspec(
    source: str = typer.Argument(..., help="Google URL またはローカル .xlsx パス"),
    sheet: str = typer.Option(..., "--sheet", help="対象シート名（例: 3.結合テスト項目）"),
    to: str = typer.Option("md", "--to", help="出力形式: md | json"),
    mapping: Optional[str] = typer.Option(None, "--mapping", help="列マッピング TOML（省略時は既定）"),
    output: Optional[str] = typer.Option(None, "--output", "-o", help="出力先ファイル（省略時は標準出力）"),
) -> None:
    cfg = core.load_mapping(mapping) if mapping else None
    spec = core.to_testspec(source, sheet, cfg)
    result = core.serialize_testspec(spec, to)
    if output:
        with open(output, "w", encoding="utf-8") as f:
            f.write(result)
        typer.echo(f"書き出しました: {output}")
    else:
        typer.echo(result)


@app.command()
def auth() -> None:
    """OAuth 認証（不足スコープがあれば再同意）。token を ~/.config/gwex に保存。"""
    from gwex.fetcher import auth as auth_mod

    creds = auth_mod.get_credentials()
    typer.echo("認証 OK。付与スコープ:")
    for s in creds.scopes or []:
        typer.echo(f"  - {s}")


@app.command(name="write-testspec")
def write_testspec(
    target: str = typer.Argument(..., help="書き込み先 .xlsx（テンプレ）"),
    sheet: str = typer.Option(..., "--sheet", help="対象シート名"),
    json_path: str = typer.Option(..., "--json", help="TestSpec JSON ファイル"),
    mapping: Optional[str] = typer.Option(None, "--mapping", help="列マッピング TOML（省略時は既定）"),
    append: bool = typer.Option(False, "--append", help="既存データの下に追記（採番は既存グループ最大Noから継続）"),
    dedup: bool = typer.Option(False, "--dedup", help="既存と完全一致するケースを除外（冪等）"),
    fmt: bool = typer.Option(False, "--format", help="追記ブロックに既存体裁（罫線/縦結合/親番号）を後付け（.xlsx）"),
    output: Optional[str] = typer.Option(None, "--output", "-o", help="出力先 .xlsx（既定 in-place）"),
) -> None:
    cfg = core.load_mapping(mapping) if mapping else None
    with open(json_path, encoding="utf-8") as f:
        spec_json = f.read()
    dest = core.write_testspec(spec_json, target, sheet, cfg, append=append, dedup=dedup, apply_format=fmt, output=output)
    typer.echo(f"テスト仕様を書き込みました: {dest}")


@app.command(name="update-case")
def update_case(
    target: str = typer.Argument(..., help="ローカル .xlsx パス または Google スプレッドシート URL"),
    sheet: str = typer.Option(..., "--sheet", help="対象シート名"),
    row: int = typer.Option(..., "--row", help="更新対象の行番号（source_row）"),
    verification: str = typer.Option(..., "--verification", help="確認内容"),
    steps: str = typer.Option("", "--steps", help="実施手順（改行区切り）"),
    mapping: Optional[str] = typer.Option(None, "--mapping", help="列マッピング TOML"),
) -> None:
    cfg = core.load_mapping(mapping) if mapping else None
    step_list = [s for s in steps.split("\n") if s.strip()]
    dest = core.update_case(target, sheet, row, verification, step_list, cfg)
    typer.echo(f"ケースを更新しました（行 {row}）: {dest}")


@app.command(name="set-text")
def set_text(
    target: str = typer.Argument(..., help="ローカル .xlsx パス または Google スプレッドシート URL"),
    sheet: str = typer.Option(..., "--sheet", help="対象シート名"),
    cell: str = typer.Option(..., "--cell", help="セル（例: B2）"),
    text: str = typer.Option(..., "--text", help="書き込むテキスト"),
    output: Optional[str] = typer.Option(None, "--output", "-o", help="出力先 .xlsx（xlsx のみ・既定 in-place）"),
) -> None:
    from gwex.writer import base

    dest = base.set_cell_text(target, sheet, cell, text, output=output)
    typer.echo(f"書き込みました: {dest}")


@app.command(name="set-image")
def set_image(
    target: str = typer.Argument(..., help="ローカル .xlsx パス または Google スプレッドシート URL"),
    sheet: str = typer.Option(..., "--sheet", help="対象シート名"),
    cell: str = typer.Option(..., "--cell", help="アンカーセル（例: B2）。--range 指定時は左上として扱う"),
    image: str = typer.Option(..., "--image", help="画像ファイルパス"),
    cell_range: Optional[str] = typer.Option(None, "--range", help="画像枠の結合範囲（例: C7:F20）。指定時は結合＋オーバーレイ"),
    insert_rows: bool = typer.Option(False, "--insert-rows", help="枠ぶんの行を挿入して確実に空枠を作る（既存を押し下げ）"),
    width: Optional[int] = typer.Option(None, "--width", help="幅(px)"),
    height: Optional[int] = typer.Option(None, "--height", help="高さ(px)"),
    max_dim: Optional[int] = typer.Option(None, "--max-dim", help="長辺上限px（超過時に自動縮小）"),
    scale: float = typer.Option(1.0, "--scale", help="枠フィット後の倍率（Google: 0.8で枠の80%＝余白付き・中央寄せ）"),
    output: Optional[str] = typer.Option(None, "--output", "-o", help="出力先 .xlsx（xlsx のみ・既定 in-place）"),
) -> None:
    from gwex.writer import base

    dest = base.set_cell_image(
        target, sheet, cell, image,
        cell_range=cell_range, insert_rows=insert_rows, width=width, height=height,
        max_dim=max_dim, scale=scale, output=output,
    )
    typer.echo(f"画像を挿入しました: {dest}")


@app.command(name="set-section")
def set_section(
    target: str = typer.Argument(..., help="ローカル .xlsx パス"),
    sheet: str = typer.Option(..., "--sheet", help="対象シート名（例: 2.画面イメージ(iOS)）"),
    top_row: int = typer.Option(..., "--top-row", help="セクション見出しの行（1始まり）"),
    title: str = typer.Option(..., "--title", help="セクション見出し（画面名など）"),
    before: Optional[str] = typer.Option(None, "--before", help="修正前 画像パス"),
    after: Optional[str] = typer.Option(None, "--after", help="修正後 画像パス"),
    left_cols: str = typer.Option("C,L", "--cols", help="セクション左右端の列（既定 C,L）"),
    split_col: str = typer.Option("H", "--split", help="修正後の開始列（既定 H）"),
    output: Optional[str] = typer.Option(None, "--output", "-o", help="出力先 .xlsx（既定 in-place）"),
) -> None:
    """枠付き before/after セクションを作成して画像を配置する。

    画像は横幅=カラム幅で固定、縦は画像高さに収まる行数だけ確保（空行を作らない）。
    """
    from gwex.writer import xlsx_format

    lc = tuple(left_cols.split(","))
    dest = xlsx_format.create_before_after_section(
        target, sheet, top_row, title, before, after,
        left_cols=lc, split_col=split_col, output=output,
    )
    typer.echo(f"枠付きセクションを作成しました: {dest}")


def main() -> None:
    app()


if __name__ == "__main__":
    main()
