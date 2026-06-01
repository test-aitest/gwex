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


def main() -> None:
    app()


if __name__ == "__main__":
    main()
