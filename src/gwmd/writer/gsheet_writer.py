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

from gwmd.fetcher import auth, google


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


def set_cell_image(
    target: str,
    sheet: str,
    cell: str,
    image_path: str,
    *,
    width: Optional[int] = None,
    height: Optional[int] = None,
) -> str:
    """Apps Script Web アプリ経由でセル内画像を base64 埋め込みする。

    要: 環境変数 GWMD_APPSCRIPT_URL（デプロイ済み Web アプリ URL）。
    画像は data URL で送るため公開リンク不要・PII 露出なし。
    """
    import requests

    app_url = os.environ.get("GWMD_APPSCRIPT_URL")
    if not app_url:
        raise RuntimeError(
            "GWMD_APPSCRIPT_URL が未設定です。appscript/gwmd_image.gs を Web アプリとして"
            "デプロイし、その URL を環境変数に設定してください。"
        )
    mime = mimetypes.guess_type(image_path)[0] or "image/png"
    with open(image_path, "rb") as f:
        data_url = f"data:{mime};base64," + base64.b64encode(f.read()).decode("ascii")

    _, creds = _service()
    resp = requests.post(
        app_url,
        headers={"Authorization": f"Bearer {creds.token}"},
        json={
            "spreadsheetId": _spreadsheet_id(target),
            "sheet": sheet,
            "cell": cell,
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
