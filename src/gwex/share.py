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


def share_file(path: str, *, convert: bool = False, public: bool = True, name: str | None = None) -> str:
    """Drive にアップロードして共有 URL を返す。

    convert=True なら Google ネイティブ形式(スプレッドシート/ドキュメント/スライド)に変換して
    アップロード(ネイティブ共有 URL になる)。public=True で anyone-with-link 閲覧可に設定。
    name を渡すと Drive 上の表示名をそれにする。**中間ファイル名(例 2026-000xxx_概要設計書_結合テスト項目)がそのまま
    スプシ名になる事故を防ぐため、変換前の画像なし版を share するときは正式名を --name で渡す。**
    """
    creds = auth.get_credentials()
    drive = build("drive", "v3", credentials=creds, cache_discovery=False)

    fname = os.path.basename(path)
    ext = fname.rsplit(".", 1)[-1].lower() if "." in fname else ""
    # 表示名: --name があればそれを正本に。無ければファイル名(convert 時は拡張子を落とす)。
    display = name or (fname.rsplit(".", 1)[0] if convert else fname)
    body: dict[str, str] = {"name": display}
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


def rename_file(source: str, name: str) -> str:
    """既存の Drive ファイル(スプシ/ドキュメント等)の表示名を変更する。

    source は URL でも file_id でも可。share 後に表示名がズレていたときの修正に使う
    (例: 中間ファイル名のままになったスプシを正式名に直す)。
    """
    from gwex.fetcher import google as _google

    file_id = _google.detect(source) or source.strip()
    creds = auth.get_credentials()
    drive = build("drive", "v3", credentials=creds, cache_discovery=False)
    f = (
        drive.files()
        .update(fileId=file_id, body={"name": name}, fields="id,name,webViewLink", supportsAllDrives=True)
        .execute()
    )
    url = f.get("webViewLink") or f"https://drive.google.com/file/d/{file_id}/view"
    return f.get("name", name), url
