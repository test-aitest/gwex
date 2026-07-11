"""Google スプレッドシートへの書き戻し（privileged）。

- テキスト/構造: Sheets API `values.update` / `values.batchUpdate`。
- 画像: Sheets REST に画像挿入が無いため、Apps Script Web アプリ経由で
  ネイティブのセル内画像を base64 埋め込み（公開リンク不要）。
"""

from __future__ import annotations

import base64
import mimetypes
import os
from typing import Optional

from googleapiclient.discovery import build

from gwex.domains.model import DEFAULT_MAPPING, MappingConfig
from gwex.fetcher import auth, google

_DATA_ROLES = (
    "screen_name", "medium_category", "small_category",
    "test_no", "verification_content", "execution_steps",
)


def _service():
    creds = auth.get_credentials()
    return build("sheets", "v4", credentials=creds, cache_discovery=False), creds


def _spreadsheet_id(target: str) -> str:
    fid = google.detect(target)
    if fid is None:
        raise ValueError(f"Google スプレッドシート URL ではありません: {target}")
    return fid


def set_cell_text(target: str, sheet: str, cell: str, text: str) -> str:
    svc, _ = _service()
    sid = _spreadsheet_id(target)
    svc.spreadsheets().values().update(
        spreadsheetId=sid,
        range=f"'{sheet}'!{cell}",
        valueInputOption="RAW",
        body={"values": [[text]]},
    ).execute()
    return target


def write_cells(target: str, sheet: str, assignments: list[tuple[str, object]]) -> str:
    """[(A1, value), ...] を values.batchUpdate で一括書込。"""
    svc, _ = _service()
    sid = _spreadsheet_id(target)
    data = [
        {"range": f"'{sheet}'!{cell}", "values": [[value]]}
        for cell, value in assignments
    ]
    if data:
        svc.spreadsheets().values().batchUpdate(
            spreadsheetId=sid,
            body={"valueInputOption": "RAW", "data": data},
        ).execute()
    return target


def last_data_row(target: str, sheet: str, mapping: Optional[MappingConfig] = None) -> int:
    """testspec のデータ列に値がある最終行（1始まり）。データ無しなら data_start_row-1。"""
    cfg = mapping or DEFAULT_MAPPING
    cols = [i for i in (cfg.column_index(r) for r in _DATA_ROLES) if i is not None]
    svc, _ = _service()
    sid = _spreadsheet_id(target)
    rows = (
        svc.spreadsheets().values()
        .get(spreadsheetId=sid, range=f"'{sheet}'", majorDimension="ROWS")
        .execute()
        .get("values", [])
    )
    last = cfg.data_start_row - 1
    for r_idx, row in enumerate(rows, start=1):
        if r_idx < cfg.data_start_row:
            continue
        if any(c < len(row) and str(row[c]).strip() for c in cols):
            last = r_idx
    return last


def _get_sheet_id(svc, sid: str, sheet: str) -> int:
    """シート名からシートIDを取得する。"""
    meta = svc.spreadsheets().get(spreadsheetId=sid).execute()
    for s in meta["sheets"]:
        if s["properties"]["title"] == sheet:
            return s["properties"]["sheetId"]
    raise ValueError(f"シートが見つかりません: {sheet}")


def set_row_height(target: str, sheet: str, row: int, height_pt: float) -> str:
    """指定行の高さを設定する（Excel pt 単位 → Sheets API pixel 変換）。

    row: 1始まりの行番号。
    height_pt: Excel ポイント単位の行高（例: 66.75）。
               Sheets API は pixel 単位で受け取る（1pt × 96/72 = 1.333...px）。
    """
    svc, _ = _service()
    sid = _spreadsheet_id(target)
    sheet_id = _get_sheet_id(svc, sid, sheet)
    height_px = max(1, round(height_pt * 96 / 72))
    svc.spreadsheets().batchUpdate(
        spreadsheetId=sid,
        body={
            "requests": [{
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "ROWS",
                        "startIndex": row - 1,  # 0-indexed
                        "endIndex": row,
                    },
                    "properties": {"pixelSize": height_px},
                    "fields": "pixelSize",
                }
            }]
        },
    ).execute()
    return target


def set_col_width(target: str, sheet: str, col: str, width_px: int) -> str:
    """指定列の幅を設定する（Sheets API・pixel 単位）。col は列文字（例: C）。"""
    from openpyxl.utils.cell import column_index_from_string

    svc, _ = _service()
    sid = _spreadsheet_id(target)
    sheet_id = _get_sheet_id(svc, sid, sheet)
    idx = column_index_from_string(col) - 1  # 0-indexed
    svc.spreadsheets().batchUpdate(
        spreadsheetId=sid,
        body={
            "requests": [{
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": idx,
                        "endIndex": idx + 1,
                    },
                    "properties": {"pixelSize": max(1, int(width_px))},
                    "fields": "pixelSize",
                }
            }]
        },
    ).execute()
    return target


def autofit_rows(target: str, sheet: str, start_row: int, end_row: int) -> str:
    """行高をコンテンツに合わせて自動調整する（Sheets API AutoResizeDimensionsRequest）。

    start_row, end_row: 1始まり、end_row は含む（inclusive）。
    セルの wrap_text が True の行は折り返し後のテキスト全体が見えるよう自動フィットされる。
    """
    svc, _ = _service()
    sid = _spreadsheet_id(target)
    sheet_id = _get_sheet_id(svc, sid, sheet)
    svc.spreadsheets().batchUpdate(
        spreadsheetId=sid,
        body={
            "requests": [{
                "autoResizeDimensions": {
                    "dimensions": {
                        "sheetId": sheet_id,
                        "dimension": "ROWS",
                        "startIndex": start_row - 1,  # 0-indexed (inclusive)
                        "endIndex": end_row,           # 0-indexed (exclusive) = end_row (1-based inclusive)
                    }
                }
            }]
        },
    ).execute()
    return target


def set_cell_image(
    target: str,
    sheet: str,
    cell: str,
    image_path: str,
    *,
    cell_range: Optional[str] = None,
    insert_rows: bool = False,
    width: Optional[int] = None,
    height: Optional[int] = None,
    max_dim: Optional[int] = None,
    scale: float = 1.0,
) -> str:
    """Apps Script Web アプリ経由でセル内画像を base64 埋め込みする。

    scale: 枠フィット後にさらに掛ける倍率（0.8 で枠の80%＝周囲に余白）。中央寄せされる。

    要: 環境変数 GWEX_APPSCRIPT_URL（デプロイ済み Web アプリ URL）。
    画像は data URL で送るため公開リンク不要・PII 露出なし。
    """
    import requests

    app_url = os.environ.get("GWEX_APPSCRIPT_URL")
    if not app_url:
        raise RuntimeError(
            "GWEX_APPSCRIPT_URL が未設定です。appscript/gwex_image.gs を Web アプリとして"
            "デプロイし、その URL を環境変数に設定してください。"
        )
    from gwex.writer._image import downscale, fit_pixels

    # キャプチャ画像そのものを使い、寸法だけ縮小する。
    # Apps Script は 100万画素・2MB 上限のため、max_dim 指定後も画素上限を必ず満たす。
    src = downscale(image_path, max_dim)
    src = fit_pixels(src)
    mime = mimetypes.guess_type(src)[0] or "image/png"
    with open(src, "rb") as f:
        data_url = f"data:{mime};base64," + base64.b64encode(f.read()).decode("ascii")

    _, creds = _service()
    resp = requests.post(
        app_url,
        headers={"Authorization": f"Bearer {creds.token}"},
        json={
            "spreadsheetId": _spreadsheet_id(target),
            "sheet": sheet,
            "cell": cell,
            "range": cell_range,        # 指定時は結合してオーバーレイ
            "insertRows": insert_rows,  # 枠ぶんの行を挿入して確実に空枠を作る
            "dataUrl": data_url,
            "width": width,
            "height": height,
            "scale": scale,
        },
        timeout=60,
    )
    resp.raise_for_status()
    result = resp.json()
    if not result.get("ok"):
        raise RuntimeError(f"Apps Script 画像挿入に失敗: {result}")
    return target
