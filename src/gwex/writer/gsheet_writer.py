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
) -> str:
    """Apps Script Web アプリ経由でセル内画像を base64 埋め込みする。

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
    from gwex.writer._image import downscale

    src = downscale(image_path, max_dim)
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
        },
        timeout=60,
    )
    resp.raise_for_status()
    result = resp.json()
    if not result.get("ok"):
        raise RuntimeError(f"Apps Script 画像挿入に失敗: {result}")
    return target
