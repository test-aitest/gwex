"""Google OAuth2（InstalledApp フロー）。privileged 層。

client_secret.json（自分の GCP で発行した Desktop クライアント）を読み、
トークンを ~/.config/gwmd/token.json にキャッシュする。スコープは readonly のみ。
"""

from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

# プロジェクト root / カレントの .env を読み込む（既存の環境変数は上書きしない）。
load_dotenv()

SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/documents.readonly",
    "https://www.googleapis.com/auth/spreadsheets",  # 読み書き（書き戻しに必要）
    "https://www.googleapis.com/auth/presentations.readonly",
]

CONFIG_DIR = Path(os.environ.get("GWMD_CONFIG_DIR", Path.home() / ".config" / "gwmd"))
TOKEN_PATH = CONFIG_DIR / "token.json"


def _client_secret_path() -> Path:
    env = os.environ.get("GWMD_CLIENT_SECRET")
    return Path(env) if env else CONFIG_DIR / "client_secret.json"


def _build_flow() -> InstalledAppFlow:
    """OAuth フローを構築する。優先順位:

    1. client_secret.json ファイル（GWMD_CLIENT_SECRET / ~/.config/gwmd/）
    2. 環境変数 GWMD_CLIENT_ID + GWMD_CLIENT_SECRET_VALUE（JSON 不要）
    """
    secret = _client_secret_path()
    if secret.exists():
        return InstalledAppFlow.from_client_secrets_file(str(secret), SCOPES)

    client_id = os.environ.get("GWMD_CLIENT_ID")
    client_secret = os.environ.get("GWMD_CLIENT_SECRET_VALUE")
    if client_id and client_secret:
        config = {
            "installed": {
                "client_id": client_id,
                "client_secret": client_secret,
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "redirect_uris": ["http://localhost"],
            }
        }
        return InstalledAppFlow.from_client_config(config, SCOPES)

    raise FileNotFoundError(
        "OAuth クライアントが見つかりません。次のいずれかを設定してください:\n"
        f"  (a) {secret} に client_secret.json を配置\n"
        "  (b) 環境変数 GWMD_CLIENT_ID と GWMD_CLIENT_SECRET_VALUE を設定（JSON 不要）"
    )


def _granted_scopes() -> set[str]:
    """token.json に実際に保存された付与スコープを読む（要求スコープではなく）。"""
    import json

    try:
        return set(json.loads(TOKEN_PATH.read_text()).get("scopes", []))
    except Exception:
        return set()


def get_credentials() -> Credentials:
    """有効な認証情報を返す。無ければ／付与スコープ不足なら OAuth 同意フローを走らせる。"""
    creds: Credentials | None = None
    # 実際の付与スコープが要求スコープを満たす時だけ既存トークンを使う。
    if TOKEN_PATH.exists() and set(SCOPES).issubset(_granted_scopes()):
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
        if creds.valid:
            return creds
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
            CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            TOKEN_PATH.write_text(creds.to_json(), encoding="utf-8")
            return creds
    # トークン無し or スコープ不足 → 新規同意（ブラウザ）
    creds = _build_flow().run_local_server(port=0)
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    TOKEN_PATH.write_text(creds.to_json(), encoding="utf-8")
    return creds
