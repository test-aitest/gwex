"""Drive 共有 URL の取得。

ローカルの Office ファイル(.xlsx/.docx/.pptx/.pdf)を Google Drive にアップロードし、
「リンクを知っている全員が閲覧可(anyone reader)」に設定して、埋め込み/閲覧用の共有 URL を返す。
アプリ(Engram)はこの URL を DB に保持し、右ペインで preview 埋め込み表示する。

`drive.file` スコープが必要(gwex auth を再実行して付与)。
"""

from __future__ import annotations

import os

from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

from gwex.fetcher import auth

_OOXML = {
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "pdf": "application/pdf",
}
_GOOGLE = {
    "xlsx": "application/vnd.google-apps.spreadsheet",
    "docx": "application/vnd.google-apps.document",
    "pptx": "application/vnd.google-apps.presentation",
}


def share_file(path: str, *, convert: bool = False, public: bool = True) -> str:
    """Drive にアップロードして共有 URL を返す。

    convert=True なら Google ネイティブ形式(スプレッドシート/ドキュメント/スライド)に変換して
    アップロード(ネイティブ共有 URL になる)。public=True で anyone-with-link 閲覧可に設定。
    """
    creds = auth.get_credentials()
    drive = build("drive", "v3", credentials=creds, cache_discovery=False)

    name = os.path.basename(path)
    ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
    body: dict[str, str] = {"name": name}
    if convert and ext in _GOOGLE:
        body["mimeType"] = _GOOGLE[ext]
    media = MediaFileUpload(path, mimetype=_OOXML.get(ext), resumable=False)

    f = (
        drive.files()
        .create(body=body, media_body=media, fields="id,webViewLink", supportsAllDrives=True)
        .execute()
    )
    file_id = f["id"]
    if public:
        drive.permissions().create(
            fileId=file_id,
            body={"role": "reader", "type": "anyone"},
            supportsAllDrives=True,
        ).execute()
    return f.get("webViewLink") or f"https://drive.google.com/file/d/{file_id}/view"
