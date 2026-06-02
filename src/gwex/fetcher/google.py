"""Google Workspace / Drive 取得層（privileged: ネットワーク I/O はここだけ）。

URL から file_id を取り出し、Drive API で**実 mimeType を判定**してから取得方法を分岐する:
  - ネイティブ Google 形式 → Docs/Sheets/Slides API の構造化 JSON（戦略B）
  - アップロードされた Office 形式 → Drive で実バイトをダウンロード → OOXML パーサへ

いずれも raw バイト列を返し、解析は同一の confined parser に合流する。
"""

from __future__ import annotations

import json
import re

from googleapiclient.discovery import build

from gwex.fetcher import auth

# Google の各種 URL から file_id を取り出す
_ID_RE = re.compile(
    r"(?:/d/|/document/d/|/spreadsheets/d/|/presentation/d/|[?&]id=)([A-Za-z0-9_-]{20,})"
)

# 実 mimeType → source 種別
_MIME_TO_KIND = {
    "application/vnd.google-apps.document": "gdocs",
    "application/vnd.google-apps.spreadsheet": "gsheet",
    "application/vnd.google-apps.presentation": "gslide",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "docx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "xlsx",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation": "pptx",
}

_GOOGLE_NATIVE = {"gdocs", "gsheet", "gslide"}


def detect(source: str) -> str | None:
    """Google/Drive の URL なら file_id を返す。"""
    if "docs.google.com" not in source and "drive.google.com" not in source:
        return None
    m = _ID_RE.search(source)
    return m.group(1) if m else None


def fetch(file_id: str) -> tuple[str, bytes]:
    """file_id を解決し (source種別, raw バイト列) を返す。

    raw は Google ネイティブなら構造化 JSON、Office なら元ファイルのバイト列。
    """
    creds = auth.get_credentials()
    drive = build("drive", "v3", credentials=creds, cache_discovery=False)
    meta = (
        drive.files()
        .get(fileId=file_id, fields="mimeType,name", supportsAllDrives=True)
        .execute()
    )
    mime = meta.get("mimeType", "")
    kind = _MIME_TO_KIND.get(mime)
    if kind is None:
        raise ValueError(f"未対応の mimeType です: {mime}（{meta.get('name')}）")

    if kind in _GOOGLE_NATIVE:
        return kind, json.dumps(_fetch_native(kind, file_id, creds)).encode("utf-8")
    # アップロードされた Office ファイル: 実バイトをダウンロード
    data = drive.files().get_media(fileId=file_id, supportsAllDrives=True).execute()
    return kind, data


_XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def fetch_xlsx(file_id: str) -> bytes:
    """file_id を xlsx バイト列として取得する（testspec 抽出など openpyxl 用）。

    - ネイティブ Google スプレッドシート → Drive API で xlsx に**エクスポート**（結合セル含む）
    - アップロード済み .xlsx → 実バイトをダウンロード
    それ以外は ValueError。
    """
    creds = auth.get_credentials()
    drive = build("drive", "v3", credentials=creds, cache_discovery=False)
    meta = (
        drive.files()
        .get(fileId=file_id, fields="mimeType,name", supportsAllDrives=True)
        .execute()
    )
    mime = meta.get("mimeType", "")
    if mime == "application/vnd.google-apps.spreadsheet":
        return drive.files().export_media(fileId=file_id, mimeType=_XLSX_MIME).execute()
    if mime == _XLSX_MIME:
        return drive.files().get_media(fileId=file_id, supportsAllDrives=True).execute()
    raise ValueError(
        f"testspec 抽出は Google スプレッドシートまたは .xlsx のみ対応です: {mime}（{meta.get('name')}）"
    )


def _fetch_native(kind: str, file_id: str, creds) -> dict:
    if kind == "gdocs":
        svc = build("docs", "v1", credentials=creds, cache_discovery=False)
        return svc.documents().get(documentId=file_id).execute()
    if kind == "gsheet":
        svc = build("sheets", "v4", credentials=creds, cache_discovery=False)
        return (
            svc.spreadsheets()
            .get(spreadsheetId=file_id, includeGridData=True)
            .execute()
        )
    svc = build("slides", "v1", credentials=creds, cache_discovery=False)
    return svc.presentations().get(presentationId=file_id).execute()
