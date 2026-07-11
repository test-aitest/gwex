"""ZIP 手術後の .xlsm/.xlsx を base と突き合わせて整合検査する（pure）。

図形1000個超・VBA 付きの巨大ブックを部分編集した際の「壊していないこと」の
機械証明。bytes dict 同士を比較し、以下を検査する:

- changed/added/removed エントリの列挙。expect_entries 指定時は
  それ以外の変更を ERROR（最小タッチ検証）。
- xl/vbaProject.bin のバイト一致（不一致・消失は ERROR）。
- 変更された worksheet XML: well-formed / <row r> の昇順・一意 /
  行内の全 <c r> の行番号が row と一致（過去に regex 編集で
  「row 154 のセルが row 155 に紛れる」破損が実際に起きたクラス）/
  行内セルの列順が昇順（r 属性は正しいのに列順を壊して注入される
  破損クラス。行番号一致だけでは検出できない）/
  mergeCell 参照の形式・向き・重複・count 属性。
- 変更された drawing XML: well-formed / pic・sp・cxnSp・grpSp 数の
  base との差分（減少は WARNING）/ cNvPr id 重複（ERROR）/
  r:embed が rels と実エントリに解決できるか（ERROR）。
- sharedStrings.xml の変更は WARNING（書き込みは inlineStr が方針）。

issue は {level, kind, entry, detail}。ERROR が1件でもあれば ok=False。
"""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from typing import Optional

from gwex.domains import sheet_scan

_WS_RE = re.compile(r"^xl/worksheets/sheet\d+\.xml$")
_DRAW_RE = re.compile(r"^xl/drawings/drawing\d+\.xml$")
_VBA = "xl/vbaProject.bin"


def _issue(issues: list[dict], level: str, kind: str, entry: str, detail: str) -> None:
    issues.append({"level": level, "kind": kind, "entry": entry, "detail": detail})


def _well_formed(name: str, data: bytes, issues: list[dict]) -> bool:
    try:
        ET.fromstring(data)
        return True
    except ET.ParseError as e:
        _issue(issues, "ERROR", "xml_malformed", name, f"XML パース不能: {e}")
        return False


def _check_worksheet(name: str, data: bytes, issues: list[dict]) -> None:
    if not _well_formed(name, data, issues):
        return
    text = data.decode("utf-8")

    prev = 0
    seen: set[int] = set()
    # 自己終了 <row/> を次行まで飲み込まない iter_row_blocks で行を切る
    # （<row\b([^>]*)(?:/>|>(.*?)</row>) 型 regex の偽陽性クラスを排除）。
    for attrs, body, _start, _end in sheet_scan.iter_row_blocks(text):
        rr = re.search(r'\br="(\d+)"', attrs)
        if not rr:
            _issue(issues, "ERROR", "row_no_r", name, f"r 属性の無い row: {attrs!r}")
            continue
        r = int(rr.group(1))
        if r in seen:
            _issue(issues, "ERROR", "row_duplicate", name, f"row r={r} が重複")
        if r <= prev:
            _issue(issues, "ERROR", "row_order", name, f"row r={r} が昇順でない（直前 r={prev}）")
        seen.add(r)
        prev = max(prev, r)
        prev_col, prev_ref = 0, ""
        for cm in re.finditer(r'<c\b[^>]*?\br="([A-Z]+)(\d+)"', body):
            ref = cm.group(1) + cm.group(2)
            if int(cm.group(2)) != r:
                _issue(issues, "ERROR", "cell_row_mismatch", name,
                       f"row r={r} の中にセル {ref} が紛れている")
            ci = _col_idx(cm.group(1))
            if ci <= prev_col:
                _issue(issues, "ERROR", "cell_col_order", name,
                       f"row r={r} のセル列順が昇順でない（{prev_ref} の後に {ref}）")
            prev_col, prev_ref = ci, ref

    refs = re.findall(r'<mergeCell ref="([^"]*)"/>', text)
    seen_refs: set[str] = set()
    for ref in refs:
        m = re.fullmatch(r"([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?", ref)
        if not m:
            _issue(issues, "ERROR", "merge_ref_invalid", name, f"mergeCell 参照が不正: {ref!r}")
            continue
        if m.group(3):
            c1 = _col_idx(m.group(1)), int(m.group(2))
            c2 = _col_idx(m.group(3)), int(m.group(4))
            if c2[0] < c1[0] or c2[1] < c1[1]:
                _issue(issues, "ERROR", "merge_ref_invalid", name, f"mergeCell の向きが不正: {ref!r}")
        if ref in seen_refs:
            _issue(issues, "ERROR", "merge_duplicate", name, f"mergeCell が重複: {ref!r}")
        seen_refs.add(ref)
    cm = re.search(r'<mergeCells count="(\d+)">', text)
    if cm and int(cm.group(1)) != len(refs):
        _issue(issues, "ERROR", "merge_count_mismatch", name,
               f"mergeCells count={cm.group(1)} だが実体は {len(refs)} 件")


def _col_idx(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _shape_counts(text: str) -> dict[str, int]:
    return {k: len(re.findall(r"<(?:xdr:)?%s[ >]" % k, text))
            for k in ("pic", "sp", "cxnSp", "grpSp")}


def _check_drawing(name: str, base_data: Optional[bytes], data: bytes,
                   edited: dict[str, bytes], issues: list[dict]) -> None:
    if not _well_formed(name, data, issues):
        return
    text = data.decode("utf-8")

    if base_data is not None:
        before = _shape_counts(base_data.decode("utf-8"))
        after = _shape_counts(text)
        for k in before:
            if after[k] < before[k]:
                _issue(issues, "WARNING", "shape_count_decreased", name,
                       f"{k}: {before[k]} → {after[k]}（意図的な削除でなければ喪失）")

    ids = re.findall(r'<(?:xdr:)?cNvPr [^>]*?id="(\d+)"', text)
    dup = sorted({i for i in ids if ids.count(i) > 1})
    if dup:
        _issue(issues, "ERROR", "cnvpr_id_duplicate", name, f"cNvPr id 重複: {', '.join(dup)}")

    rels_path = f"xl/drawings/_rels/{name.split('/')[-1]}.rels"
    rels = edited.get(rels_path, b"").decode("utf-8")
    for rid in sorted(set(re.findall(r'r:embed="(rId\d+)"', text))):
        rel = re.search(r'<Relationship\b[^>]*\bId="' + re.escape(rid) + r'"[^>]*/>', rels)
        if not rel:
            _issue(issues, "ERROR", "embed_unresolved", name, f"r:embed={rid} が rels に無い")
            continue
        if 'TargetMode="External"' in rel.group(0):
            continue
        mt = re.search(r'Target="([^"]+)"', rel.group(0))
        target = mt.group(1) if mt else ""
        if "../" in target:
            target = "xl/" + target.split("../", 1)[1]
        else:
            target = target.lstrip("/")
        if target not in edited:
            _issue(issues, "ERROR", "embed_target_missing", name,
                   f"r:embed={rid} の Target が実在しない: {target}")


def verify(base: dict[str, bytes], edited: dict[str, bytes],
           expect_entries: Optional[list[str]] = None) -> dict:
    issues: list[dict] = []
    changed = sorted(n for n in base if n in edited and base[n] != edited[n])
    added = sorted(n for n in edited if n not in base)
    removed = sorted(n for n in base if n not in edited)

    if expect_entries is not None:
        allow = set(expect_entries)
        for n in changed + added + removed:
            if n not in allow:
                _issue(issues, "ERROR", "unexpected_change", n, "expect-entries に無いエントリが変化した")

    for n in removed:
        _issue(issues, "ERROR", "entry_removed", n, "エントリが消失した（部分編集で削除は起きないはず）")

    if _VBA in base:
        if _VBA not in edited:
            _issue(issues, "ERROR", "vba_removed", _VBA, "vbaProject.bin が消失した")
        elif base[_VBA] != edited[_VBA]:
            _issue(issues, "ERROR", "vba_changed", _VBA, "vbaProject.bin のバイトが変化した")

    for n in changed + added:
        if _WS_RE.match(n):
            _check_worksheet(n, edited[n], issues)
        elif _DRAW_RE.match(n):
            _check_drawing(n, base.get(n), edited[n], edited, issues)

    if "xl/sharedStrings.xml" in changed:
        _issue(issues, "WARNING", "shared_strings_changed", "xl/sharedStrings.xml",
               "sharedStrings が変更された（新規テキストは inlineStr が方針）")

    errors = [i for i in issues if i["level"] == "ERROR"]
    return {
        "ok": not errors,
        "changed_entries": changed,
        "added_entries": added,
        "removed_entries": removed,
        "issues": issues,
        "counts": {"errors": len(errors),
                   "warnings": len([i for i in issues if i["level"] == "WARNING"])},
    }


def summarize(result: dict) -> str:
    lines = [
        f"{'OK' if result['ok'] else 'NG'}  changed={len(result['changed_entries'])} "
        f"added={len(result['added_entries'])} removed={len(result['removed_entries'])} "
        f"ERROR={result['counts']['errors']} WARN={result['counts']['warnings']}",
    ]
    for n in result["changed_entries"]:
        lines.append(f"  changed: {n}")
    for n in result["added_entries"]:
        lines.append(f"  added:   {n}")
    for n in result["removed_entries"]:
        lines.append(f"  removed: {n}")
    for i in result["issues"]:
        lines.append(f"  [{i['level']}] {i['kind']} {i['entry']}: {i['detail']}")
    return "\n".join(lines)
