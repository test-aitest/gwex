"""fetcher（取得）と parser（解析）を調停するオーケストレーション層。

ここは pure ではない（fetcher を import してよい）。pure 層は parsers/serializers/ir のみ。
解析は既定で sandbox（network/fs-write deny の別プロセス）経由で行う。
"""

from __future__ import annotations

from pathlib import Path

from gwmd import parser_entry, sandbox
from gwmd.fetcher import local
from gwmd.ir import Document
from gwmd.serializers import json as json_ser
from gwmd.serializers import markdown as md_ser


def to_document(source: str, *, use_sandbox: bool = True) -> Document:
    """ソース指定（Google URL / ローカル OOXML パス）を IR Document に変換する。

    fetcher だけがネットワーク/ディスク I/O を行い、解析は raw（bytes / JSON）に対して
    confined parser プロセスで実行する。
    """
    from gwmd.fetcher import google

    file_id = google.detect(source)
    if file_id is not None:
        kind, raw = google.fetch(file_id)
        return _parse(kind, raw, use_sandbox)

    path = Path(source)
    if path.exists():
        kind = local.detect_source(path)
        if kind is None:
            raise ValueError(f"未対応のファイル形式です: {path.suffix}")
        return _parse(kind, local.read_bytes(path), use_sandbox)

    raise ValueError(f"ソースを解決できません: {source}")


def _parse(kind: str, raw: bytes, use_sandbox: bool) -> Document:
    if use_sandbox and sandbox.available():
        return sandbox.run_parser_confined(kind, raw)
    return parser_entry._parse(kind, raw)


def serialize(doc: Document, to: str) -> str:
    if to == "md":
        return md_ser.render(doc)
    if to == "json":
        return json_ser.render(doc)
    raise ValueError(f"未対応の出力形式: {to}（md|json）")


def convert(source: str, to: str = "md", *, use_sandbox: bool = True) -> str:
    return serialize(to_document(source, use_sandbox=use_sandbox), to)


# --------------------------------------------------------------------------
# テスト仕様（ドメイン抽出）
# --------------------------------------------------------------------------


def _xlsx_bytes(source: str) -> bytes:
    """source（ローカル .xlsx / Google URL）から xlsx バイト列を得る。"""
    from gwmd.fetcher import google

    file_id = google.detect(source)
    if file_id is not None:
        kind, raw = google.fetch(file_id)
        if kind != "xlsx":
            raise ValueError(
                "ネイティブ Google シートの testspec 抽出は後続対応です。"
                "アップロードされた Excel かローカル .xlsx を指定してください。"
            )
        return raw
    path = Path(source)
    if not path.exists():
        raise ValueError(f"ソースを解決できません: {source}")
    if local.detect_source(path) != "xlsx":
        raise ValueError("testspec 抽出は .xlsx のみ対応です。")
    return local.read_bytes(path)


def to_testspec(source: str, sheet: str, mapping=None):
    from gwmd.domains import testspec

    return testspec.extract(_xlsx_bytes(source), sheet, mapping)


def serialize_testspec(spec, to: str) -> str:
    from gwmd.serializers import testspec_json, testspec_md

    if to == "md":
        return testspec_md.render(spec)
    if to == "json":
        return testspec_json.render(spec)
    raise ValueError(f"未対応の出力形式: {to}（md|json）")


def write_testspec(source_json: str, target: str, sheet: str, mapping=None, *, output=None) -> str:
    """TestSpec(JSON) を target（現状ローカル .xlsx）のシートに展開記入する。"""
    from gwmd.domains import testspec_writer
    from gwmd.domains.model import TestSpec
    from gwmd.fetcher import google

    spec = TestSpec.model_validate_json(source_json)
    asgn = testspec_writer.assignments(spec, mapping)
    if google.detect(target) is not None:
        from gwmd.writer import gsheet_writer

        return gsheet_writer.write_cells(target, sheet, asgn)
    from gwmd.writer import xlsx_writer

    return xlsx_writer.write_cells(target, sheet, asgn, output=output)


def load_mapping(path: str):
    """TOML から MappingConfig を読む（privileged: ファイル I/O）。"""
    import tomllib

    from gwmd.domains.model import MappingConfig

    with open(path, "rb") as f:
        raw = tomllib.load(f)
    options = raw.pop("options", {})
    return MappingConfig(**raw, **options)
