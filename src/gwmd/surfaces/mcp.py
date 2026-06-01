"""gwmd の MCP 表面（typed read/write をエージェントに公開）。

gwmd は LLM を持たず、構造化データの読み書きだけを提供する。生成・判断・
キャプチャ・Slack 応答は呼び出し側エージェントの責務。

起動: `gwmd-mcp`（stdio）。Claude Code 等の MCP 設定から接続する。
"""

from __future__ import annotations

from typing import Optional

from mcp.server.fastmcp import FastMCP

from gwmd import core
from gwmd.writer import base as writer_base

mcp = FastMCP("gwmd")


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
def write_testspec(target: str, sheet: str, spec_json: str, mapping_path: Optional[str] = None) -> str:
    """TestSpec(JSON) をシートに展開記入する（target: ローカル .xlsx / Google URL）。"""
    cfg = core.load_mapping(mapping_path) if mapping_path else None
    return core.write_testspec(spec_json, target, sheet, cfg)


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
    width: Optional[int] = None,
    height: Optional[int] = None,
) -> str:
    """セルに画像を挿入する（Excel=埋め込み / Google=Apps Script でセル内画像）。"""
    return writer_base.set_cell_image(target, sheet, cell, image_path, width=width, height=height)


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
