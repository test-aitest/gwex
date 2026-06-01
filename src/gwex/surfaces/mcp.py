"""gwex の MCP 表面（typed read/write をエージェントに公開）。

gwex は LLM を持たず、構造化データの読み書きだけを提供する。生成・判断・
キャプチャ・Slack 応答は呼び出し側エージェントの責務。

起動: `gwex-mcp`（stdio）。Claude Code 等の MCP 設定から接続する。
"""

from __future__ import annotations

from typing import Optional

from mcp.server.fastmcp import FastMCP

from gwex import core
from gwex.writer import base as writer_base

mcp = FastMCP("gwex")


@mcp.tool()
def convert(uri: str, format: str = "md") -> str:
    """文書（Google URL / ローカル OOXML）を Markdown か typed JSON に変換する。format: md|json"""
    return core.convert(uri, to=format)


@mcp.tool()
def extract_testspec(uri: str, sheet: str, mapping_path: Optional[str] = None) -> str:
    """結合テスト項目シートを階層 TestSpec(JSON) として抽出する。"""
    cfg = core.load_mapping(mapping_path) if mapping_path else None
    spec = core.to_testspec(uri, sheet, cfg)
    return core.serialize_testspec(spec, "json")


@mcp.tool()
def write_testspec(
    target: str,
    sheet: str,
    spec_json: str,
    mapping_path: Optional[str] = None,
    append: bool = False,
    dedup: bool = False,
) -> str:
    """TestSpec(JSON) をシートに記入。append=既存の下に追記、dedup=完全一致を除外（冪等）。"""
    cfg = core.load_mapping(mapping_path) if mapping_path else None
    return core.write_testspec(spec_json, target, sheet, cfg, append=append, dedup=dedup)


@mcp.tool()
def update_case(
    target: str,
    sheet: str,
    row: int,
    verification_content: str,
    steps: list[str],
    mapping_path: Optional[str] = None,
) -> str:
    """既存ケース行(source_row)の確認内容・実施手順を上書きする（意味的マージの更新用）。"""
    cfg = core.load_mapping(mapping_path) if mapping_path else None
    return core.update_case(target, sheet, row, verification_content, steps, cfg)


@mcp.tool()
def set_cell_text(target: str, sheet: str, cell: str, text: str) -> str:
    """セルにテキストを書き込む（target: ローカル .xlsx / Google URL）。"""
    return writer_base.set_cell_text(target, sheet, cell, text)


@mcp.tool()
def set_cell_image(
    target: str,
    sheet: str,
    cell: str,
    image_path: str,
    cell_range: Optional[str] = None,
    insert_rows: bool = False,
    width: Optional[int] = None,
    height: Optional[int] = None,
    max_dim: Optional[int] = None,
) -> str:
    """画像をセルの上にオーバーレイで乗せる。cell_range 指定で結合枠にフィット、insert_rows で空枠を挿入。max_dim で長辺自動縮小。"""
    return writer_base.set_cell_image(
        target, sheet, cell, image_path,
        cell_range=cell_range, insert_rows=insert_rows, width=width, height=height, max_dim=max_dim,
    )


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
