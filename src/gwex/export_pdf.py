"""Google スプレッドシートの指定シートを PDF（任意で PNG）に書き出す。

anken-verify のレンダリング目視（画像枠内・体裁確認）で使う。
従来は呼び出し側でアドホックな urllib スクリプトを書いていたが、再現性のため
gwex コマンド（``gwex export-pdf``）に集約した。

Google 認証は既存の ``auth.get_credentials()`` を再利用する。Sheets API には
シート単位の PDF export が無いため、``docs.google.com/.../export`` エンドポイントを
Bearer トークンで叩く（gid・range・余白などを指定できる唯一の経路）。
"""

from __future__ import annotations

import os
import re
import subprocess
import urllib.request
from typing import Optional

from googleapiclient.discovery import build

from gwex.fetcher import google as _google
from gwex.fetcher.auth import get_credentials

_RAW_ID_RE = re.compile(r"^[A-Za-z0-9_-]{20,}$")


def _resolve_file_id(source: str) -> str:
    """URL なら file_id を抽出。生 ID ならそのまま返す。"""
    fid = _google.detect(source)
    if fid:
        return fid
    if _RAW_ID_RE.match(source.strip()):
        return source.strip()
    raise ValueError(f"スプレッドシートの URL/ID として解釈できません: {source}")


def _resolve_gids(file_id: str, creds) -> dict[str, int]:
    svc = build("sheets", "v4", credentials=creds, cache_discovery=False)
    meta = (
        svc.spreadsheets()
        .get(spreadsheetId=file_id, fields="sheets(properties(sheetId,title))")
        .execute()
    )
    return {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta["sheets"]}


def export_pdf(
    source: str,
    sheets: list[str],
    out_dir: str = "/tmp",
    *,
    png: bool = False,
    portrait: bool = True,
    fit_width: bool = True,
    gridlines: bool = True,
    prefix: str = "gwex_pdf_",
) -> list[dict]:
    """``sheets`` の各シートを PDF に書き出す。``png=True`` で sips により PNG も生成。

    返り値は ``[{"sheet","pdf","bytes"[,"png"]} | {"sheet","error"}]``。
    """
    file_id = _resolve_file_id(source)
    creds = get_credentials()
    gids = _resolve_gids(file_id, creds)
    token = creds.token
    os.makedirs(out_dir, exist_ok=True)

    results: list[dict] = []
    for title in sheets:
        if title not in gids:
            results.append({"sheet": title, "error": f"シートが見つかりません（実在: {', '.join(gids)}）"})
            continue
        gid = gids[title]
        safe = re.sub(r"[^A-Za-z0-9_.-]", "_", title)
        pdf_path = os.path.join(out_dir, f"{prefix}{safe}.pdf")
        params = [
            "format=pdf",
            f"gid={gid}",
            f"portrait={'true' if portrait else 'false'}",
            f"fitw={'true' if fit_width else 'false'}",
            f"gridlines={'true' if gridlines else 'false'}",
            "top_margin=0.20",
            "bottom_margin=0.20",
            "left_margin=0.20",
            "right_margin=0.20",
        ]
        url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?" + "&".join(params)
        req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
        with urllib.request.urlopen(req) as resp, open(pdf_path, "wb") as f:
            f.write(resp.read())
        entry: dict = {"sheet": title, "pdf": pdf_path, "bytes": os.path.getsize(pdf_path)}
        if png:
            png_path = os.path.join(out_dir, f"{prefix}{safe}.png")
            subprocess.run(
                ["sips", "-s", "format", "png", pdf_path, "--out", png_path],
                check=True,
                capture_output=True,
            )
            entry["png"] = png_path
        results.append(entry)
    return results
