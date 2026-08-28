"""Google OAuth2（InstalledApp フロー）。privileged 層。

OAuth クライアント（自分の GCP で発行した Desktop クライアント）を
**環境変数 → ~/.config/gwex/client_secret.json → .env** の優先順位で解決し、
トークンを ~/.config/gwex/token.json にキャッシュする。

要求スコープは 5 つ（``SCOPES`` 参照）:
Drive / Docs / Slides は readonly、``spreadsheets`` は読み書き（シートへの
書き戻しに必要）、``drive.file`` は自分が作成したファイルの作成・共有（share 用）。
"""

from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path
from typing import Optional

from dotenv import find_dotenv, load_dotenv
from google.auth.exceptions import RefreshError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

# .env 読み込み前から設定されていた環境変数を記録する
# （環境変数 → client_secret.json → .env の優先順位判定に使う）。
_ENV_BEFORE_DOTENV = frozenset(k for k in os.environ if k.startswith("GWEX_"))

# .env を読み込む（既存の環境変数は上書きしない）。
# 1) このモジュールの位置から上方探索 — editable インストールでリポジトリ直下の .env を拾う
# 2) カレントディレクトリから上方探索 — uv tool install 等の配布版でも、
#    リポジトリ内（.env のある場所以下）から実行すれば拾える
load_dotenv()
_dotenv_from_cwd = find_dotenv(usecwd=True)
if _dotenv_from_cwd:
    load_dotenv(_dotenv_from_cwd)

SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/drive.file",  # 自分が作成したファイルの作成/共有（share 用）
    "https://www.googleapis.com/auth/documents.readonly",
    "https://www.googleapis.com/auth/spreadsheets",  # 読み書き（書き戻しに必要）
    "https://www.googleapis.com/auth/presentations.readonly",
]

def _env(name: str) -> Optional[str]:
    """環境変数 GWEX_<name> を読む（.env 由来の値を含む）。"""
    return os.environ.get(f"GWEX_{name}")


def _real_env(name: str) -> Optional[str]:
    """.env 由来ではなく、プロセス起動時から設定されていた環境変数 GWEX_<name> を読む。"""
    key = f"GWEX_{name}"
    return os.environ.get(key) if key in _ENV_BEFORE_DOTENV else None


CONFIG_DIR = Path(_env("CONFIG_DIR") or (Path.home() / ".config" / "gwex"))
TOKEN_PATH = CONFIG_DIR / "token.json"


def _flow_from(
    secret_path: Optional[str],
    client_id: Optional[str],
    client_secret: Optional[str],
) -> Optional[InstalledAppFlow]:
    """client_secret.json のパス、または ID/シークレット値からフローを構築する。

    パスが未指定/存在しない かつ ID/シークレットも揃っていなければ None。
    """
    if secret_path and Path(secret_path).exists():
        return InstalledAppFlow.from_client_secrets_file(secret_path, SCOPES)
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
    return None


def _build_flow() -> InstalledAppFlow:
    """OAuth フローを構築する。優先順位:

    1. 環境変数（GWEX_CLIENT_SECRET のパス、または GWEX_CLIENT_ID + GWEX_CLIENT_SECRET_VALUE）
    2. ~/.config/gwex/client_secret.json（GWEX_CONFIG_DIR で変更可）
    3. .env に書かれた同名の変数（開発・editable インストール向けフォールバック）
    """
    # 1. 環境変数（.env 由来の値を除く）
    flow = _flow_from(
        _real_env("CLIENT_SECRET"), _real_env("CLIENT_ID"), _real_env("CLIENT_SECRET_VALUE")
    )
    if flow:
        return flow

    # 2. 設定ディレクトリの client_secret.json
    default_secret = CONFIG_DIR / "client_secret.json"
    if default_secret.exists():
        return InstalledAppFlow.from_client_secrets_file(str(default_secret), SCOPES)

    # 3. .env 由来の値（従来挙動との互換）
    flow = _flow_from(_env("CLIENT_SECRET"), _env("CLIENT_ID"), _env("CLIENT_SECRET_VALUE"))
    if flow:
        return flow

    raise FileNotFoundError(
        "OAuth クライアントが見つかりません。次のいずれかを設定してください:\n"
        "  (a) 環境変数 GWEX_CLIENT_ID と GWEX_CLIENT_SECRET_VALUE を設定（JSON 不要）\n"
        f"  (b) {default_secret} に client_secret.json を配置\n"
        "  (c) リポジトリ直下の .env に同名の変数を記述（開発・editable インストール向け）"
    )


def _granted_scopes() -> set[str]:
    """token.json に実際に保存された付与スコープを読む（要求スコープではなく）。"""
    import json

    try:
        return set(json.loads(TOKEN_PATH.read_text()).get("scopes", []))
    except Exception:
        return set()


def _backup_stale_token(reason: str) -> None:
    """使えなくなった token.json をタイムスタンプ付きで退避する（原因の追跡用）。"""
    if not TOKEN_PATH.exists():
        return
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    TOKEN_PATH.rename(TOKEN_PATH.with_name(f"{TOKEN_PATH.name}.{reason}_{ts}"))


def get_credentials() -> Credentials:
    """有効な認証情報を返す。無ければ／付与スコープ不足なら OAuth 同意フローを走らせる。"""
    creds: Credentials | None = None
    # 実際の付与スコープが要求スコープを満たす時だけ既存トークンを使う。
    if TOKEN_PATH.exists() and set(SCOPES).issubset(_granted_scopes()):
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
        if creds.valid:
            return creds
        if creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                CONFIG_DIR.mkdir(parents=True, exist_ok=True)
                TOKEN_PATH.write_text(creds.to_json(), encoding="utf-8")
                return creds
            except RefreshError:
                # refresh_token 自体が失効/取り消し（テストステータスの 7 日失効など）。
                # 古いトークンを退避し、下の新規同意フローへフォールバックする。
                _backup_stale_token("revoked")
    # トークン無し or スコープ不足 or refresh 失敗 → 新規同意（ブラウザ）
    creds = _build_flow().run_local_server(port=0)
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    TOKEN_PATH.write_text(creds.to_json(), encoding="utf-8")
    return creds
