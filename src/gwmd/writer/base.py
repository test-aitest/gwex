"""書き戻しの共通インターフェース。target でバックエンドを分岐する。

target がローカル .xlsx パス → xlsx_writer、Google URL → gsheet_writer。
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional


def _backend(target: str):
    from gwmd.fetcher import google

    if google.detect(target) is not None:
        from gwmd.writer import gsheet_writer

        return gsheet_writer, True  # (module, is_google)
    if Path(target).suffix.lower() != ".xlsx":
        raise ValueError("書き戻し対象はローカル .xlsx か Google スプレッドシート URL です。")
    from gwmd.writer import xlsx_writer

    return xlsx_writer, False


def set_cell_text(target: str, sheet: str, cell: str, text: str, *, output: Optional[str] = None) -> str:
    backend, is_google = _backend(target)
    if is_google:
        return backend.set_cell_text(target, sheet, cell, text)
    return backend.set_cell_text(target, sheet, cell, text, output=output)


def set_cell_image(
    target: str,
    sheet: str,
    anchor: str,
    image_path: str,
    *,
    width: Optional[int] = None,
    height: Optional[int] = None,
    output: Optional[str] = None,
) -> str:
    backend, is_google = _backend(target)
    if is_google:
        return backend.set_cell_image(target, sheet, anchor, image_path, width=width, height=height)
    return backend.set_cell_image(
        target, sheet, anchor, image_path, width=width, height=height, output=output
    )
