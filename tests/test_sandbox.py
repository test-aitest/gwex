"""sandbox（confined parser プロセス）の検証。

sandbox-exec が無い環境ではスキップ。
"""

from __future__ import annotations

import subprocess
import sys
from io import BytesIO

import openpyxl
import pytest

from gwmd import parser_entry, sandbox

pytestmark = pytest.mark.skipif(not sandbox.available(), reason="sandbox-exec が無い")


def _xlsx_bytes() -> bytes:
    wb = openpyxl.Workbook()
    wb.active.append(["a", "b"])
    wb.active.append([1, 2])
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_confined_parse_matches_inprocess():
    raw = _xlsx_bytes()
    confined = sandbox.run_parser_confined("xlsx", raw)
    inproc = parser_entry._parse("xlsx", raw)
    assert confined.model_dump_json() == inproc.model_dump_json()


def test_profile_denies_network():
    """defense-in-depth: confine 下では発信ができない（SSRF 不可）。"""
    proc = subprocess.run(
        ["sandbox-exec", "-p", sandbox._PROFILE, sys.executable, "-c",
         "import socket; socket.create_connection(('8.8.8.8', 53), 2)"],
        capture_output=True, text=True,
    )
    assert proc.returncode != 0
    assert "not permitted" in proc.stderr.lower()


def test_profile_denies_file_write():
    proc = subprocess.run(
        ["sandbox-exec", "-p", sandbox._PROFILE, sys.executable, "-c",
         "open('/tmp/gwmd_sbx_probe.txt','w').write('x')"],
        capture_output=True, text=True,
    )
    assert proc.returncode != 0
    assert "not permitted" in proc.stderr.lower()
