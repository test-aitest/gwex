"""confined プロセスの entrypoint。

stdin から raw（OOXML はバイト列 / Google は JSON）を受け取り、pure parser で
IR に変換し、IR JSON を stdout に書く。**この入口は fetcher を import しない**
（pure layer のみに依存）。sandbox-exec 下で network/fs-write を奪われた状態で動く。

    python -m gwex.parser_entry <kind>   < raw  > ir.json
"""

from __future__ import annotations

import json
import sys

_GOOGLE = {"gdocs", "gsheet", "gslide"}
_OOXML = {"docx", "xlsx", "pptx"}


def _parse(kind: str, raw: bytes):
    if kind in _GOOGLE:
        data = json.loads(raw.decode("utf-8"))
        if kind == "gdocs":
            from gwex.parsers import gdocs

            return gdocs.parse(data)
        if kind == "gsheet":
            from gwex.parsers import gsheet

            return gsheet.parse(data)
        from gwex.parsers import gslide

        return gslide.parse(data)
    if kind in _OOXML:
        if kind == "docx":
            from gwex.parsers import docx

            return docx.parse(raw)
        if kind == "xlsx":
            from gwex.parsers import xlsx

            return xlsx.parse(raw)
        from gwex.parsers import pptx

        return pptx.parse(raw)
    raise ValueError(f"未対応の種別: {kind}")


def main() -> None:
    if len(sys.argv) != 2:
        sys.stderr.write("usage: python -m gwex.parser_entry <kind>\n")
        sys.exit(2)
    raw = sys.stdin.buffer.read()
    doc = _parse(sys.argv[1], raw)
    sys.stdout.write(doc.model_dump_json())


if __name__ == "__main__":
    main()
