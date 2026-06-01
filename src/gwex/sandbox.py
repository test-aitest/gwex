"""parser を network/fs-write を奪った別プロセスで実行する（macOS sandbox-exec）。

安全境界をプロセス境界にする設計の核。fetcher（特権）が取得した raw を、
発信能力の無い parser プロセスに渡す。悪意ある文書が来ても SSRF を起こせない。

NOTE: sandbox-exec は Apple では deprecated 扱いだが現状動作する。将来 bay の
Confine 層へ委譲する。
"""

from __future__ import annotations

import shutil
import subprocess
import sys

from gwex.ir import Document

# Seatbelt profile: 既定拒否、ファイル読取とプロセス起動のみ許可、
# **ネットワークと書込みは拒否**（SSRF・改竄・書込み経由の流出を封じる）。
_PROFILE = """(version 1)
(deny default)
(allow process-fork)
(allow process-exec)
(allow sysctl-read)
(allow mach-lookup)
(allow file-read*)
(deny file-write*)
(deny network*)
"""


def available() -> bool:
    return shutil.which("sandbox-exec") is not None


def run_parser_confined(kind: str, raw: bytes) -> Document:
    """raw を confined parser プロセスに渡し、IR Document を得る。"""
    proc = subprocess.run(
        ["sandbox-exec", "-p", _PROFILE, sys.executable, "-m", "gwex.parser_entry", kind],
        input=raw,
        capture_output=True,
    )
    if proc.returncode != 0:
        raise RuntimeError(
            f"confined parser が失敗しました (kind={kind}, rc={proc.returncode}):\n"
            + proc.stderr.decode("utf-8", "replace")
        )
    return Document.model_validate_json(proc.stdout)
