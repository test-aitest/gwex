"""書き戻しの共通インターフェース。target でバックエンドを分岐する。

target がローカル .xlsx パス → xlsx_writer、Google URL → gsheet_writer。
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional


def _backend(target: str):
    from gwex.fetcher import google

    if google.detect(target) is not None:
        from gwex.writer import gsheet_writer

        return gsheet_writer, True  # (module, is_google)
    if Path(target).suffix.lower() != ".xlsx":
        raise ValueError("書き戻し対象はローカル .xlsx か Google スプレッドシート URL です。")
    from gwex.writer import xlsx_writer

    return xlsx_writer, False


def set_cell_text(target: str, sheet: str, cell: str, text: str,
                  *, value_type: Optional[str] = None, output: Optional[str] = None) -> str:
    """セルにテキストを書き込む。value_type='int' を指定すると整数として書き込む。"""
    backend, is_google = _backend(target)
    value: object = text
    if value_type == "int":
        value = int(text)
    elif value_type == "float":
        value = float(text)
    if is_google:
        return backend.set_cell_text(target, sheet, cell, value)
    return backend.set_cell_text(target, sheet, cell, value, output=output)


def set_cell_image(
    target: str,
    sheet: str,
    anchor: str,
    image_path: str,
    *,
    cell_range: Optional[str] = None,
    insert_rows: bool = False,
    width: Optional[int] = None,
    height: Optional[int] = None,
    max_dim: Optional[int] = None,
    scale: float = 1.0,
    output: Optional[str] = None,
) -> str:
    backend, is_google = _backend(target)
    if is_google:
        return backend.set_cell_image(
            target, sheet, anchor, image_path,
            cell_range=cell_range, insert_rows=insert_rows, width=width, height=height,
            max_dim=max_dim, scale=scale,
        )
    # ローカル .xlsx の画像注入は xlsx_zip 経由（最小タッチの ZIP 注入）。openpyxl 版
    # （xlsx_writer.set_cell_image）も画像自体は保持するが、round-trip でグラフ/ピボット等の
    # 高度要素をドロップし得る・既存 media を再圧縮するため、任意 doc への安全側として使わない。
    from gwex.writer import xlsx_zip

    return xlsx_zip.set_cell_image(
        target, sheet, anchor, image_path,
        cell_range=cell_range, insert_rows=insert_rows, width=width, height=height, max_dim=max_dim, output=output,
    )


def set_row_height(target: str, sheet: str, row: int, height: float, *, output: Optional[str] = None) -> str:
    """行高を設定する（xlsx のみ対応）。"""
    backend, is_google = _backend(target)
    if is_google:
        raise ValueError("set-row-height は xlsx のみ対応です。")
    return backend.set_row_height(target, sheet, row, height, output=output)


def set_col_width(target: str, sheet: str, col: str, width: float, *, output: Optional[str] = None) -> str:
    """列幅を設定する（xlsx=Excel 列幅単位 / Google スプシ=pixel 単位）。

    前後の画像枠が左右で非対称な列幅のとき、狭い列を他列と同じ幅に揃えて
    set-image の枠フィットが前後同サイズになるようにするのに使う（画像は無加工のまま）。
    """
    backend, is_google = _backend(target)
    if is_google:
        return backend.set_col_width(target, sheet, col, int(width))
    return backend.set_col_width(target, sheet, col, width, output=output)


def set_merge_cells(target: str, sheet: str, ranges: list[str], *, output: Optional[str] = None) -> str:
    """セル範囲を結合する（xlsx のみ対応）。"""
    backend, is_google = _backend(target)
    if is_google:
        raise ValueError("set-merge は xlsx のみ対応です（Google スプシは未対応）。")
    return backend.set_merge_cells(target, sheet, ranges, output=output)


def autofit_rows(target: str, sheet: str, start_row: int, end_row: int) -> str:
    """行高をコンテンツに合わせて自動調整する（Google スプシのみ対応）。

    xlsx は Excel がファイルを開くときに customHeight=False の行を自動フィットする。
    Google スプシは AutoResizeDimensionsRequest API を呼び出す。
    """
    backend, is_google = _backend(target)
    if not is_google:
        raise ValueError(
            "autofit-rows は Google スプレッドシートのみ対応です。"
            "xlsx は Excel が開いたときに customHeight=False の行を自動フィットします。"
        )
    return backend.autofit_rows(target, sheet, start_row, end_row)
