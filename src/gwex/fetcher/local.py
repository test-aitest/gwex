"""ローカルファイル取得（privileged: ディスク I/O はこの層だけ）。

parser には**バイト列**だけを渡す。parser 側は BytesIO で純粋に解析し、
パスを開かない（＝任意ファイル読取の余地を持たない）。
"""

from __future__ import annotations

from pathlib import Path

# 拡張子 → source 種別（parser ディスパッチ用）
EXT_TO_SOURCE = {
    ".docx": "docx",
    ".xlsx": "xlsx",
    ".pptx": "pptx",
}


def detect_source(path: str | Path) -> str | None:
    """拡張子から source 種別を判定。未対応なら None。"""
    return EXT_TO_SOURCE.get(Path(path).suffix.lower())


def read_bytes(path: str | Path) -> bytes:
    """ファイルをバイト列として読む。"""
    return Path(path).read_bytes()
