"""gwmd の CLI。

    gwmd convert <source> --to md|json
"""

from __future__ import annotations

from typing import Optional

import typer

from gwmd import core

app = typer.Typer(add_completion=False, help="文書を Markdown / typed JSON に変換する")


@app.callback()
def _root() -> None:
    """gwmd: Google Workspace + Microsoft 文書 → Markdown / typed JSON。"""


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
    """OAuth 認証（不足スコープがあれば再同意）。token を ~/.config/gwmd に保存。"""
    from gwmd.fetcher import auth as auth_mod

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
    output: Optional[str] = typer.Option(None, "--output", "-o", help="出力先 .xlsx（既定 in-place）"),
) -> None:
    cfg = core.load_mapping(mapping) if mapping else None
    with open(json_path, encoding="utf-8") as f:
        spec_json = f.read()
    dest = core.write_testspec(spec_json, target, sheet, cfg, output=output)
    typer.echo(f"テスト仕様を書き込みました: {dest}")


@app.command(name="set-text")
def set_text(
    target: str = typer.Argument(..., help="ローカル .xlsx パス または Google スプレッドシート URL"),
    sheet: str = typer.Option(..., "--sheet", help="対象シート名"),
    cell: str = typer.Option(..., "--cell", help="セル（例: B2）"),
    text: str = typer.Option(..., "--text", help="書き込むテキスト"),
    output: Optional[str] = typer.Option(None, "--output", "-o", help="出力先 .xlsx（xlsx のみ・既定 in-place）"),
) -> None:
    from gwmd.writer import base

    dest = base.set_cell_text(target, sheet, cell, text, output=output)
    typer.echo(f"書き込みました: {dest}")


@app.command(name="set-image")
def set_image(
    target: str = typer.Argument(..., help="ローカル .xlsx パス または Google スプレッドシート URL"),
    sheet: str = typer.Option(..., "--sheet", help="対象シート名"),
    cell: str = typer.Option(..., "--cell", help="アンカーセル（例: B2）"),
    image: str = typer.Option(..., "--image", help="画像ファイルパス"),
    width: Optional[int] = typer.Option(None, "--width", help="幅(px)"),
    height: Optional[int] = typer.Option(None, "--height", help="高さ(px)"),
    output: Optional[str] = typer.Option(None, "--output", "-o", help="出力先 .xlsx（xlsx のみ・既定 in-place）"),
) -> None:
    from gwmd.writer import base

    dest = base.set_cell_image(target, sheet, cell, image, width=width, height=height, output=output)
    typer.echo(f"画像を挿入しました: {dest}")


def main() -> None:
    app()


if __name__ == "__main__":
    main()
