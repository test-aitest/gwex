"""宣言的 ops による .xlsm/.xlsx の ZIP 手術（privileged）。

47MB/73MB のマクロ付き画面設計書（図形1000個超・vbaProject.bin あり）を、
openpyxl の load→save を通さずに部分編集する。触るのは対象エントリだけで、
それ以外（VBA・他シート・sharedStrings・media）はバイト単位で温存する。

方針（実測に基づく設計判断）:
- 新規テキストは sharedStrings に触らず inlineStr セルで書く（既存セルの
  style s= は温存）。既存セルの共有文字列参照を書き換えないのは、
  sharedStrings をバイト不変に保ち diff を最小にするため。
- 日本語の一致判定は必ず NFC 正規化を通す（NFD/NFC 罠）。
- 全 op 成功時のみ書き出し。適用後に domains.xlsm_verify の内部整合
  チェック（row/cell 整合ほか）を通し、通らなければ書き出さない。
  過去に regex アドホック編集で「row 154 のセルが row 155 に紛れる」破損が
  実際に起きたクラスを、書き出し前に機械検出するため。

ops 仕様は apply_ops docstring と surfaces/cli.py の apply-ops を参照。
"""

from __future__ import annotations

import datetime as _dt
import math
import os
import re
import unicodedata
from typing import Callable, Optional
from xml.sax.saxutils import escape

from gwex.domains import sheet_scan, xlsm_verify
from gwex.writer import xlsx_zip

_EMU = xlsx_zip._EMU  # 9525 = 1px
_EPOCH = _dt.date(1899, 12, 30)  # Excel シリアル値の起点

_WORKSHEET_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
_DRAWING_CT = "application/vnd.openxmlformats-officedocument.drawing+xml"
_WORKSHEET_RTYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"


class OpError(ValueError):
    """op の適用失敗（理由つき）。apply_ops が捕捉して失敗レポートにする。"""


def _nfc(s: str) -> str:
    return unicodedata.normalize("NFC", str(s))


def _require(params: dict, key: str, op: str):
    if key not in params or params[key] is None:
        raise OpError(f"{op}: 必須パラメータがありません: {key}")
    return params[key]


def _sheet_of(entries: dict[str, bytes], params: dict, default_sheet: Optional[str], op: str) -> str:
    sheet = params.get("sheet") or default_sheet
    if not sheet:
        raise OpError(f"{op}: sheet が未指定です（spec 先頭の sheet: か op 内 sheet: で指定）")
    try:
        sheet_scan.sheet_xml_path(entries, sheet)
    except ValueError as e:
        raise OpError(f"{op}: {e}") from e
    return sheet


# ---------------------------------------------------------------------------
# セル操作の低レベルヘルパ（sheet XML 文字列に対する最小タッチ upsert）
# ---------------------------------------------------------------------------

def _split_ref(ref: str) -> tuple[str, int]:
    m = re.fullmatch(r"([A-Z]+)(\d+)", ref.upper())
    if not m:
        raise OpError(f"セル参照が不正です: {ref!r}")
    return m.group(1), int(m.group(2))


# 開きタグは属性値対応で取る。<x\b[^>]*(?:/>|>.*?</x>) 型 regex は自己終了
# <row .../> / <c .../> で [^>]* が末尾 / を食い、「次の要素の閉じタグまで」
# 飲み込む（過去に row 154 のセルが row 155 に注入された破損クラスの真因）。
# 属性値内は "…" / '…' で読み飛ばし、値の外の > で必ず止まる。
_ATTRS = r"(?:[^>\"']|\"[^\"]*\"|'[^']*')*?"
_ROW_TAG_TMPL = r'<row\b%s\br="%%d"%s(/?)>' % (_ATTRS, _ATTRS)
_CELL_TAG_TMPL = r'<c\b%s\br="%%s"%s(/?)>' % (_ATTRS, _ATTRS)


def _find_row(sx: str, rownum: int) -> Optional[tuple[int, int, str]]:
    """row r=rownum の (start, end, block)。自己終了なら block は <row .../> のみ。"""
    m = re.search(_ROW_TAG_TMPL % rownum, sx)
    if not m:
        return None
    if m.group(1):
        return m.start(), m.end(), m.group(0)
    close = sx.find("</row>", m.end())
    if close == -1:
        return None
    end = close + len("</row>")
    return m.start(), end, sx[m.start():end]


def _find_cell(row_block: str, ref: str) -> Optional[tuple[int, int, str]]:
    """row block 内のセル r=ref の (start, end, block)。"""
    m = re.search(_CELL_TAG_TMPL % re.escape(ref), row_block)
    if not m:
        return None
    if m.group(1):
        return m.start(), m.end(), m.group(0)
    close = row_block.find("</c>", m.end())
    if close == -1:
        return None
    end = close + len("</c>")
    return m.start(), end, row_block[m.start():end]


def _cell_xml(ref: str, inner: str, t: Optional[str], s: Optional[str]) -> str:
    attrs = f' r="{ref}"'
    if s is not None:
        attrs += f' s="{s}"'
    if t is not None:
        attrs += f' t="{t}"'
    return f"<c{attrs}>{inner}</c>" if inner else f"<c{attrs}/>"


def _write_cell(sx: str, ref: str, inner: str, t: Optional[str],
                s_override: Optional[str] = None) -> str:
    """セルを upsert する。既存セルの s= は温存（s_override 指定時はそれを使う）。"""
    col, rownum = _split_ref(ref)
    col_idx = sheet_scan.col_to_index(col)
    rm = _find_row(sx, rownum)
    if rm:
        r_start, r_end, block = rm
        cm = _find_cell(block, ref)
        if cm:
            c_start, c_end, cell_block = cm
            s = s_override
            if s is None:
                sm = re.search(r'\bs="(\d+)"', cell_block[:cell_block.find(">")])
                s = sm.group(1) if sm else None
            new_block = block[:c_start] + _cell_xml(ref, inner, t, s) + block[c_end:]
        else:
            cell = _cell_xml(ref, inner, t, s_override)
            if block.endswith("/>"):  # 空行 <row .../> を展開
                new_block = block[:-2] + ">" + cell + "</row>"
            else:
                pos = None
                for existing in re.finditer(r'<c\b[^>]*?\br="([A-Z]+)\d+"', block):
                    if sheet_scan.col_to_index(existing.group(1)) > col_idx:
                        pos = existing.start()
                        break
                if pos is not None:
                    new_block = block[:pos] + cell + block[pos:]
                else:
                    new_block = block[: -len("</row>")] + cell + "</row>"
        return sx[:r_start] + new_block + sx[r_end:]

    row_xml = f'<row r="{rownum}">' + _cell_xml(ref, inner, t, s_override) + "</row>"
    for existing in re.finditer(r'<row\b[^>]*?\br="(\d+)"', sx):
        if int(existing.group(1)) > rownum:
            return sx[:existing.start()] + row_xml + sx[existing.start():]
    if "</sheetData>" in sx:
        return sx.replace("</sheetData>", row_xml + "</sheetData>", 1)
    m = re.search(r"<sheetData\s*/>", sx)
    if m:
        return sx[:m.start()] + "<sheetData>" + row_xml + "</sheetData>" + sx[m.end():]
    raise OpError("sheetData が見つかりません（worksheet XML でない可能性）")


def _extend_dimension(sx: str, row: int) -> str:
    m = re.search(r'<dimension ref="([A-Z]+\d+):([A-Z]+)(\d+)"/>', sx)
    if m and int(m.group(3)) < row:
        return sx[:m.start()] + f'<dimension ref="{m.group(1)}:{m.group(2)}{row}"/>' + sx[m.end():]
    return sx


def _inline_str(text: str) -> str:
    return f'<is><t xml:space="preserve">{escape(text)}</t></is>'


def _existing_cell_text(entries: dict[str, bytes], sx: str, ref: str) -> str:
    _, rownum = _split_ref(ref)
    rm = _find_row(sx, rownum)
    if not rm:
        return ""
    cm = _find_cell(rm[2], ref)
    if not cm:
        return ""
    cell = cm[2]
    gt = cell.find(">")
    attrs, inner = cell[:gt], (cell[gt + 1:-len("</c>")] if not cell.endswith("/>") else "")
    val = sheet_scan.cell_value(attrs, inner, sheet_scan.shared_strings(entries))
    return "" if val is None else str(val)


# ---------------------------------------------------------------------------
# 図形操作の低レベルヘルパ（drawing XML 文字列）
# ---------------------------------------------------------------------------

_SP_RE = re.compile(r"<(?:xdr:)?sp[ >].*?</(?:xdr:)?sp>", re.DOTALL)
_ANCHOR_RE = re.compile(
    r"<(?:xdr:)?(oneCellAnchor|twoCellAnchor|absoluteAnchor)\b[^>]*>.*?</(?:xdr:)?\1>",
    re.DOTALL,
)


def _drawing_of(entries: dict[str, bytes], sheet: str, op: str, create: bool = False) -> tuple[str, str]:
    sheet_xml = sheet_scan.sheet_xml_path(entries, sheet)
    found = sheet_scan.drawing_path_for_sheet(entries, sheet_xml)
    if found:
        return found
    if create:
        return xlsx_zip._drawing_for_sheet(entries, sheet_xml)
    raise OpError(f"{op}: シート {sheet} に drawing（図形レイヤ）がありません")


def _match_sp(dx: str, match_text: str, op: str) -> re.Match:
    """テキストが NFC 一致する sp を一意に特定する（0件/複数件は明示エラー）。"""
    target = _nfc(match_text)
    hits = [m for m in _SP_RE.finditer(dx) if sheet_scan.shape_text(m.group(0)) == target]
    if not hits:
        near = [sheet_scan.shape_text(m.group(0)) for m in _SP_RE.finditer(dx)
                if target[:6] and target[:6] in sheet_scan.shape_text(m.group(0))]
        hint = f"（部分一致の候補: {near[:3]}）" if near else ""
        raise OpError(f"{op}: テキストが一致する図形がありません: {match_text!r}{hint}")
    if len(hits) > 1:
        raise OpError(f"{op}: テキストが一致する図形が {len(hits)} 個あり一意に特定できません: {match_text!r}")
    return hits[0]


def _replace_sp_text(sp_block: str, new_text: str) -> str:
    """sp の txBody を新テキストで置き換える（先頭 run の rPr・pPr・endParaRPr を継承）。"""
    tb = re.search(r"<(?:xdr:)?txBody>(.*)</(?:xdr:)?txBody>", sp_block, re.DOTALL)
    if not tb:
        raise OpError("図形に txBody がありません（テキストを持てない図形）")
    inner = tb.group(1)
    paras = list(re.finditer(r"<a:p>(.*?)</a:p>", inner, re.DOTALL))
    if not paras:
        raise OpError("txBody に段落 <a:p> がありません")
    head = inner[: paras[0].start()]

    first_inner = paras[0].group(1)
    m = re.search(r"<a:(?:r|br|fld|endParaRPr)\b", first_inner)
    ppr = first_inner[: m.start()] if m else first_inner

    rpr = ""
    run = re.search(r"<a:r>(.*?)</a:r>", inner, re.DOTALL)
    if run:
        i = run.group(1).find("<a:t")
        rpr = run.group(1)[:i] if i >= 0 else run.group(1)

    last_inner = paras[-1].group(1)
    j = last_inner.find("<a:endParaRPr")
    tail = last_inner[j:] if j >= 0 else ""

    lines = _nfc(new_text).split("\n")
    brk = f"<a:br>{rpr}</a:br>" if rpr else "<a:br/>"
    runs = brk.join(f"<a:r>{rpr}<a:t>{escape(line)}</a:t></a:r>" for line in lines)
    new_inner = head + "<a:p>" + ppr + runs + tail + "</a:p>"
    return sp_block[: tb.start(1)] + new_inner + sp_block[tb.end(1):]


def _next_cnvpr_id(dx: str) -> int:
    return max((int(i) for i in re.findall(r'<(?:xdr:)?cNvPr [^>]*?id="(\d+)"', dx)), default=1) + 1


# ---------------------------------------------------------------------------
# 各 op（entries を in-place 変更し、監査ログ用サマリ文字列を返す）
# ---------------------------------------------------------------------------

def _number_repr(value) -> str:
    """number: の値を <v> の数値表現へ（int / float / 数字文字列。bool・NaN/inf は拒否）。"""
    if isinstance(value, bool):
        raise OpError(f"set_cell: number が数値ではありません: {value!r}")
    if not isinstance(value, (int, float)):
        s = str(value).strip()
        try:
            value = int(s)
        except ValueError:
            try:
                value = float(s)
            except ValueError:
                raise OpError(f"set_cell: number が数値ではありません: {value!r}")
    if isinstance(value, float):
        if not math.isfinite(value):
            raise OpError(f"set_cell: number は有限数のみ: {value!r}")
        return str(int(value)) if value.is_integer() else repr(value)
    return str(value)


def _op_set_cell(entries: dict[str, bytes], params: dict, default_sheet: Optional[str]) -> str:
    sheet = _sheet_of(entries, params, default_sheet, "set_cell")
    cell = str(_require(params, "cell", "set_cell"))
    text, number = params.get("text"), params.get("number")
    if (text is None) == (number is None):
        raise OpError("set_cell: text か number のどちらか一方を指定してください")
    mode = params.get("mode", "replace")
    if mode not in ("replace", "append"):
        raise OpError(f"set_cell: mode は replace | append: {mode!r}")

    sheet_xml = sheet_scan.sheet_xml_path(entries, sheet)
    sx = entries[sheet_xml].decode("utf-8")
    if number is not None:
        if mode != "replace":
            raise OpError("set_cell: number に mode=append は使えません")
        v = _number_repr(number)
        sx = _write_cell(sx, cell.upper(), f"<v>{v}</v>", None)  # t なし＝数値・既存 s= 温存
        detail = f"number {v}"
    else:
        new_text = _nfc(str(text))
        if mode == "append":
            old = _existing_cell_text(entries, sx, cell.upper())
            new_text = (old + "\n" + new_text) if old else new_text
        sx = _write_cell(sx, cell.upper(), _inline_str(new_text), "inlineStr")
        detail = f"{mode}, {len(new_text)}文字"
    _, rownum = _split_ref(cell)
    sx = _extend_dimension(sx, rownum)
    entries[sheet_xml] = sx.encode("utf-8")
    return f"set_cell {sheet}!{cell.upper()} ({detail})"


def _op_clear_cell(entries: dict[str, bytes], params: dict, default_sheet: Optional[str]) -> str:
    op = "clear_cell"
    sheet = _sheet_of(entries, params, default_sheet, op)
    cell = str(_require(params, "cell", op)).upper()
    _, rownum = _split_ref(cell)

    sheet_xml = sheet_scan.sheet_xml_path(entries, sheet)
    sx = entries[sheet_xml].decode("utf-8")
    rm = _find_row(sx, rownum)
    cm = _find_cell(rm[2], cell) if rm else None
    if not cm:
        return f"clear_cell {sheet}!{cell} (セルなし: no-op)"
    r_start, r_end, block = rm
    c_start, c_end, cell_block = cm
    sm = re.search(r'\bs="(\d+)"', cell_block[: cell_block.find(">")])
    new_block = (block[:c_start] + _cell_xml(cell, "", None, sm.group(1) if sm else None)
                 + block[c_end:])
    entries[sheet_xml] = (sx[:r_start] + new_block + sx[r_end:]).encode("utf-8")
    return f"clear_cell {sheet}!{cell}"


def _op_insert_rows(entries: dict[str, bytes], params: dict, default_sheet: Optional[str]) -> str:
    sheet = _sheet_of(entries, params, default_sheet, "insert_rows")
    at = int(_require(params, "at", "insert_rows"))
    count = int(params.get("count", 1))
    if at < 1 or count < 1:
        raise OpError(f"insert_rows: at/count は 1 以上: at={at} count={count}")
    copy_style_from = params.get("copy_style_from")

    sheet_xml = sheet_scan.sheet_xml_path(entries, sheet)
    found = sheet_scan.drawing_path_for_sheet(entries, sheet_xml)
    draw_path = found[0] if found else ""
    xlsx_zip._shift_rows_down(entries, sheet_xml, draw_path, at, count)

    sx = entries[sheet_xml].decode("utf-8")
    if copy_style_from is not None:
        src = int(copy_style_from)
        src += count if src >= at else 0
        rm = _find_row(sx, src)
        if not rm:
            raise OpError(f"insert_rows: copy_style_from の行がありません: {copy_style_from}")
        block = rm[2]
        open_tag = re.match(r"<row\b" + _ATTRS + r"/?>", block).group(0)
        attrs = re.sub(r'\s+r="\d+"', "", open_tag[len("<row"):].rstrip("/>").rstrip(">"))
        cells = []
        for cm in re.finditer(r"<c\b([^>]*?)(?:/>|>)", block):
            ref_m = re.search(r'\br="([A-Z]+)\d+"', cm.group(1))
            if not ref_m:
                continue
            s_m = re.search(r'\bs="(\d+)"', cm.group(1))
            cells.append((ref_m.group(1), s_m.group(1) if s_m else None))
        new_rows = ""
        for k in range(count):
            n = at + k
            cs = "".join(_cell_xml(f"{col}{n}", "", None, s or None) for col, s in cells)
            new_rows += f'<row r="{n}"{attrs}>{cs}</row>'
        pos = None
        for existing in re.finditer(r'<row\b[^>]*?\br="(\d+)"', sx):
            if int(existing.group(1)) >= at:
                pos = existing.start()
                break
        if pos is not None:
            sx = sx[:pos] + new_rows + sx[pos:]
        else:
            sx = sx.replace("</sheetData>", new_rows + "</sheetData>", 1)

    max_row = max((int(r) for r in re.findall(r'<row\b[^>]*?\br="(\d+)"', sx)), default=0)
    sx = _extend_dimension(sx, max_row)
    entries[sheet_xml] = sx.encode("utf-8")
    return f"insert_rows {sheet} at={at} count={count} copy_style_from={copy_style_from}"


def _merge_overlaps(ref: str, min_c: int, min_r: int, max_c: int, max_r: int) -> bool:
    parts = ref.split(":")
    a_col, a_row = _split_ref(parts[0])
    b_col, b_row = _split_ref(parts[1]) if len(parts) == 2 else (a_col, a_row)
    c1, c2 = sheet_scan.col_to_index(a_col), sheet_scan.col_to_index(b_col)
    lo_c, hi_c = min(c1, c2), max(c1, c2)
    lo_r, hi_r = min(a_row, b_row), max(a_row, b_row)
    return not (hi_c < min_c or lo_c > max_c or hi_r < min_r or lo_r > max_r)


def _op_set_merge(entries: dict[str, bytes], params: dict, default_sheet: Optional[str]) -> str:
    op = "set_merge"
    sheet = _sheet_of(entries, params, default_sheet, op)
    rng, ranges = params.get("range"), params.get("ranges")
    if (rng is None) == (ranges is None):
        raise OpError(f"{op}: range か ranges のどちらか一方を指定してください")
    if rng is not None:
        ranges = [rng]
    if not isinstance(ranges, (list, tuple)) or not ranges:
        raise OpError(f"{op}: ranges は結合範囲（例 B230:C233）のリストです: {ranges!r}")

    sheet_xml = sheet_scan.sheet_xml_path(entries, sheet)
    # mergeCells も </sheetData> も無いシートは _add_merge の挿入先が無いので先に展開する
    sx = entries[sheet_xml].decode("utf-8")
    if "<mergeCells" not in sx and "</sheetData>" not in sx:
        m = re.search(r"<sheetData\s*/>", sx)
        if not m:
            raise OpError(f"{op}: sheetData が見つかりません（worksheet XML でない可能性）")
        entries[sheet_xml] = (sx[:m.start()] + "<sheetData></sheetData>" + sx[m.end():]).encode("utf-8")

    done: list[str] = []
    notes: list[str] = []
    for raw in ranges:
        m = re.fullmatch(r"([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)", str(raw).strip())
        if not m:
            raise OpError(f"{op}: 結合範囲が不正です（例 B230:C233）: {raw!r}")
        c1, r1 = sheet_scan.col_to_index(m.group(1).upper()), int(m.group(2))
        c2, r2 = sheet_scan.col_to_index(m.group(3).upper()), int(m.group(4))
        if r1 < 1:
            raise OpError(f"{op}: 行番号は 1 以上: {raw!r}")
        if c2 < c1 or r2 < r1:
            raise OpError(f"{op}: 結合範囲の向きが不正です（左上:右下 の順）: {raw!r}")
        if (c1, r1) == (c2, r2):
            raise OpError(f"{op}: 1セルだけの範囲は結合できません: {raw!r}")
        ref = f"{m.group(1).upper()}{r1}:{m.group(3).upper()}{r2}"

        removed = [old for old in sheet_scan.merged_ranges(entries[sheet_xml].decode("utf-8"))
                   if _merge_overlaps(old, c1, r1, c2, r2)]
        if removed:
            xlsx_zip._remove_overlapping_merges(entries, sheet_xml, c1, r1, c2, r2)
            notes.append(f"{ref} と重なる既存結合を解除: {', '.join(removed)}")
        xlsx_zip._add_merge(entries, sheet_xml, ref)
        done.append(ref)

    note = f"（{'; '.join(notes)}）" if notes else ""
    return f"set_merge {sheet} {', '.join(done)}{note}"


def _op_append_history_row(entries: dict[str, bytes], params: dict,
                           default_sheet: Optional[str]) -> str:
    op = "append_history_row"
    sheet = params.get("sheet") or "改版履歴"
    params = dict(params, sheet=sheet)
    sheet = _sheet_of(entries, params, None, op)
    date = _require(params, "date", op)
    target = str(_require(params, "target", op))
    status = str(_require(params, "status", op))
    content = str(_require(params, "content", op))

    if isinstance(date, _dt.datetime):
        date = date.date()
    if not isinstance(date, _dt.date):
        try:
            date = _dt.date.fromisoformat(str(date))
        except ValueError as e:
            raise OpError(f"{op}: date は YYYY-MM-DD: {date!r}") from e
    serial = (date - _EPOCH).days

    sheet_xml = sheet_scan.sheet_xml_path(entries, sheet)
    sx = entries[sheet_xml].decode("utf-8")
    last = 0
    for row, _ref, col, _attrs, inner in sheet_scan.iter_cells(sx):
        if col == "A" and re.search(r"<v>|<is>", inner):
            last = max(last, row)
    if last == 0:
        raise OpError(f"{op}: A列にデータ行がありません: {sheet}")
    n = last + 1

    styles: dict[str, Optional[str]] = {}
    rm = _find_row(sx, last)
    for col in ("A", "B", "C", "E"):
        cm = _find_cell(rm[2], f"{col}{last}") if rm else None
        sm = re.search(r'\bs="(\d+)"', cm[2][: cm[2].find(">")]) if cm else None
        styles[col] = sm.group(1) if sm else None

    sx = _write_cell(sx, f"A{n}", f"<v>{serial}</v>", None, styles["A"])
    sx = _write_cell(sx, f"B{n}", _inline_str(_nfc(target)), "inlineStr", styles["B"])
    sx = _write_cell(sx, f"C{n}", _inline_str(_nfc(status)), "inlineStr", styles["C"])
    sx = _write_cell(sx, f"E{n}", _inline_str(_nfc(content)), "inlineStr", styles["E"])
    sx = _extend_dimension(sx, n)
    entries[sheet_xml] = sx.encode("utf-8")
    return f"append_history_row {sheet}!row{n} date={date.isoformat()}(serial {serial}) target={target}"


def _op_edit_shape_text(entries: dict[str, bytes], params: dict, default_sheet: Optional[str]) -> str:
    op = "edit_shape_text"
    sheet = _sheet_of(entries, params, default_sheet, op)
    match_text = str(_require(params, "match_text", op))
    new_text = str(_require(params, "new_text", op))
    draw_path, _ = _drawing_of(entries, sheet, op)
    dx = entries[draw_path].decode("utf-8")
    m = _match_sp(dx, match_text, op)
    entries[draw_path] = (dx[: m.start()] + _replace_sp_text(m.group(0), new_text)
                          + dx[m.end():]).encode("utf-8")
    return f"edit_shape_text {sheet} {match_text!r} → {new_text!r}"


def _op_add_shape(entries: dict[str, bytes], params: dict, default_sheet: Optional[str]) -> str:
    op = "add_shape"
    sheet = _sheet_of(entries, params, default_sheet, op)
    preset = str(_require(params, "preset", op))
    text = str(_require(params, "text", op))
    at_px = _require(params, "at_px", op)
    size_px = _require(params, "size_px", op)
    style_from = str(_require(params, "style_from_text", op))
    if not re.fullmatch(r"[A-Za-z][A-Za-z0-9]*", preset):
        raise OpError(f"{op}: preset 名が不正: {preset!r}")
    for name, v in (("at_px", at_px), ("size_px", size_px)):
        if not (isinstance(v, (list, tuple)) and len(v) == 2):
            raise OpError(f"{op}: {name} は [x, y] の2要素: {v!r}")

    draw_path, _ = _drawing_of(entries, sheet, op, create=True)
    dx = entries[draw_path].decode("utf-8")
    src = _match_sp(dx, style_from, op)
    clone = src.group(0)

    nid = _next_cnvpr_id(dx)
    cnv = re.search(r"<(?:xdr:)?cNvPr\b[^>]*(?:/>|>.*?</(?:xdr:)?cNvPr>)", clone, re.DOTALL)
    if not cnv:
        raise OpError(f"{op}: 複製元図形に cNvPr がありません")
    p = "xdr:" if "<xdr:wsDr" in dx else ""
    clone = clone[: cnv.start()] + f'<{p}cNvPr id="{nid}" name="gwexShape{nid}"/>' + clone[cnv.end():]

    x, y = int(at_px[0]) * _EMU, int(at_px[1]) * _EMU
    cx, cy = int(size_px[0]) * _EMU, int(size_px[1]) * _EMU
    clone = re.sub(r'<a:off x="-?\d+" y="-?\d+"/>', f'<a:off x="{x}" y="{y}"/>', clone, count=1)
    clone = re.sub(r'<a:ext cx="\d+" cy="\d+"/>', f'<a:ext cx="{cx}" cy="{cy}"/>', clone, count=1)

    src_preset = re.search(r'<a:prstGeom prst="(\w+)"', clone)
    if not src_preset:
        raise OpError(f"{op}: 複製元図形に prstGeom がありません（フリーフォームは複製元にできない）")
    if src_preset.group(1) != preset:
        # adjust 値（avLst）はプリセット固有なので、別プリセットにする時は初期化する
        clone = re.sub(r"<a:prstGeom prst=\"\w+\">.*?</a:prstGeom>",
                       f'<a:prstGeom prst="{preset}"><a:avLst/></a:prstGeom>', clone,
                       count=1, flags=re.DOTALL)
    clone = _replace_sp_text(clone, text)

    anchor = (f"<{p}absoluteAnchor><{p}pos x=\"{x}\" y=\"{y}\"/>"
              f"<{p}ext cx=\"{cx}\" cy=\"{cy}\"/>{clone}<{p}clientData/></{p}absoluteAnchor>")
    close = f"</{p}wsDr>"
    entries[draw_path] = dx.replace(close, anchor + close).encode("utf-8")
    return f"add_shape {sheet} {preset} {text!r} at_px={list(at_px)} size_px={list(size_px)}"


def _op_delete_shape(entries: dict[str, bytes], params: dict, default_sheet: Optional[str]) -> str:
    op = "delete_shape"
    sheet = _sheet_of(entries, params, default_sheet, op)
    match_text = str(_require(params, "match_text", op))
    draw_path, _ = _drawing_of(entries, sheet, op)
    dx = entries[draw_path].decode("utf-8")
    m = _match_sp(dx, match_text, op)

    for am in _ANCHOR_RE.finditer(dx):
        if am.start() <= m.start() and m.end() <= am.end():
            if re.search(r"<(?:xdr:)?grpSp[ >]", am.group(0)):
                raise OpError(f"{op}: 対象図形はグループ（grpSp）内にあり削除できません: {match_text!r}")
            entries[draw_path] = (dx[: am.start()] + dx[am.end():]).encode("utf-8")
            return f"delete_shape {sheet} {match_text!r}"
    raise OpError(f"{op}: 図形を含むアンカーが見つかりません: {match_text!r}")


# ---------------------------------------------------------------------------
# コネクタ（矢印）op
#
# 実ファイル（画面遷移図 .xlsm）の cxnSp 実測に基づく流儀:
# - 接続は <a:stCxn id idx> / <a:endCxn id idx>。矩形の接続点 idx は
#   0=上辺中点 / 1=左辺中点 / 2=下辺中点 / 3=右辺中点（左→右の水平接続は
#   st idx=3 → end idx=1 が実例の支配的パターン）。
# - 接続の有無に関わらず絶対 EMU の a:xfrm off/ext を必ず持つ。off=両端点の
#   min、ext=差の絶対値、終点が始点より左/上なら flipH/flipV（実例で xfrm の
#   端点と接続図形の辺中点が 0〜1 EMU 差で一致することを確認済み。1 EMU 差は
#   辺中点 y+h/2 の半上げ丸め）。
# - 線は <a:ln w="19050">（1.5pt）+ solidFill schemeClr tx1 + tailEnd triangle、
#   bentConnector3 は avLst に adj1=50000 を明示するのが実例の流儀。
# ---------------------------------------------------------------------------

_PIC_RE = re.compile(r"<(?:xdr:)?pic[ >].*?</(?:xdr:)?pic>", re.DOTALL)
_XFRM_RE = re.compile(
    r'<a:xfrm[^>]*><a:off x="(-?\d+)" y="(-?\d+)"/><a:ext cx="(\d+)" cy="(\d+)"/>')
_CONNECTOR_PRESETS = ("straightConnector1", "bentConnector2", "bentConnector3",
                      "bentConnector4", "bentConnector5")


def _anchor_of_match(dx: str, m: re.Match, op: str, what: str) -> re.Match:
    for am in _ANCHOR_RE.finditer(dx):
        if am.start() <= m.start() and m.end() <= am.end():
            return am
    raise OpError(f"{op}: {what} を含むアンカーが見つかりません")


def _cnvpr_id_of(block: str, op: str) -> int:
    m = re.search(r'<(?:xdr:)?cNvPr [^>]*?id="(\d+)"', block)
    if not m:
        raise OpError(f"{op}: 図形に cNvPr id がありません")
    return int(m.group(1))


def _bbox_of(block: str, op: str, what: str) -> tuple[int, int, int, int]:
    m = _XFRM_RE.search(block)
    if not m:
        raise OpError(f"{op}: {what} に a:xfrm（絶対 EMU 座標）がありません"
                      "（Excel 保存済みの実ファイルは必ず持つ。px 指定で回避可）")
    return int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))


def _top_level_pics(dx: str) -> list[str]:
    pics = []
    for am in _ANCHOR_RE.finditer(dx):
        block = am.group(0)
        if re.search(r"<(?:xdr:)?grpSp[ >]", block):
            continue
        pm = _PIC_RE.search(block)
        if pm:
            pics.append(pm.group(0))
    return pics


def _resolve_cxn_endpoint(dx: str, spec, op: str, which: str):
    """from/to の1端を (接続先id | None, bbox | None, 座標点 | None) に解決する。"""
    if not isinstance(spec, dict) or len(spec) != 1:
        raise OpError(f"{op}: {which} は shape_text / pic_index / px のどれか"
                      f"1キーの dict です: {spec!r}")
    (key, val), = spec.items()
    if key == "px":
        if not (isinstance(val, (list, tuple)) and len(val) == 2):
            raise OpError(f"{op}: {which}.px は [x, y] の2要素: {val!r}")
        return None, None, (int(val[0]) * _EMU, int(val[1]) * _EMU)
    if key == "shape_text":
        m = _match_sp(dx, str(val), op)
        am = _anchor_of_match(dx, m, op, f"{which} の図形")
        if re.search(r"<(?:xdr:)?grpSp[ >]", am.group(0)):
            raise OpError(f"{op}: {which} の図形はグループ（grpSp）内にあり接続できません: {val!r}")
        return _cnvpr_id_of(m.group(0), op), _bbox_of(m.group(0), op, f"{which} の図形"), None
    if key == "pic_index":
        pics = _top_level_pics(dx)
        i = int(val)
        if not (1 <= i <= len(pics)):
            raise OpError(f"{op}: {which}.pic_index={i} は範囲外（1〜{len(pics)}）")
        return _cnvpr_id_of(pics[i - 1], op), _bbox_of(pics[i - 1], op, f"{which} の画像"), None
    raise OpError(f"{op}: {which} のキーは shape_text | pic_index | px: {key!r}")


def _side_point(bbox: tuple[int, int, int, int], idx: int) -> tuple[int, int]:
    """接続点 idx の座標（辺中点。半分は実例に合わせ半上げ丸め）。"""
    x, y, w, h = bbox
    return {0: (x + (w + 1) // 2, y),
            1: (x, y + (h + 1) // 2),
            2: (x + (w + 1) // 2, y + h),
            3: (x + w, y + (h + 1) // 2)}[idx]


def _op_add_connector(entries: dict[str, bytes], params: dict, default_sheet: Optional[str]) -> str:
    op = "add_connector"
    sheet = _sheet_of(entries, params, default_sheet, op)
    preset = str(params.get("preset", "straightConnector1"))
    if preset not in _CONNECTOR_PRESETS:
        raise OpError(f"{op}: preset は {' | '.join(_CONNECTOR_PRESETS)}: {preset!r}")
    frm = _require(params, "from", op)
    to = _require(params, "to", op)
    try:
        w = int(round(float(params.get("line_pt", 1.5)) * sheet_scan.EMU_PER_PT))
    except (TypeError, ValueError):
        raise OpError(f"{op}: line_pt が数値ではありません: {params.get('line_pt')!r}")
    if w <= 0:
        raise OpError(f"{op}: line_pt は正の数: {params.get('line_pt')!r}")

    draw_path, _ = _drawing_of(entries, sheet, op, create=True)
    dx = entries[draw_path].decode("utf-8")
    st_id, st_bbox, st_pt = _resolve_cxn_endpoint(dx, frm, op, "from")
    en_id, en_bbox, en_pt = _resolve_cxn_endpoint(dx, to, op, "to")

    def _center(bbox, pt):
        if bbox is None:
            return pt
        x, y, bw, bh = bbox
        return x + (bw + 1) // 2, y + (bh + 1) // 2

    c1, c2 = _center(st_bbox, st_pt), _center(en_bbox, en_pt)
    dxv, dyv = c2[0] - c1[0], c2[1] - c1[1]
    if abs(dxv) >= abs(dyv):  # 支配軸が水平: 右辺(3)→左辺(1)（実例の支配的パターン）
        st_idx, en_idx = (3, 1) if dxv >= 0 else (1, 3)
    else:                     # 垂直: 下辺(2)→上辺(0)
        st_idx, en_idx = (2, 0) if dyv >= 0 else (0, 2)
    p1 = _side_point(st_bbox, st_idx) if st_bbox else st_pt
    p2 = _side_point(en_bbox, en_idx) if en_bbox else en_pt
    if p1 == p2:
        raise OpError(f"{op}: 始点と終点が同一点です（ゼロ長の矢印）: {p1}")

    ox, oy = min(p1[0], p2[0]), min(p1[1], p2[1])
    cx, cy = abs(p2[0] - p1[0]), abs(p2[1] - p1[1])
    flips = ('' + (' flipH="1"' if p2[0] < p1[0] else "")
             + (' flipV="1"' if p2[1] < p1[1] else ""))

    nid = _next_cnvpr_id(dx)
    p = "xdr:" if "<xdr:wsDr" in dx else ""
    stx = f'<a:stCxn id="{st_id}" idx="{st_idx}"/>' if st_id is not None else ""
    enx = f'<a:endCxn id="{en_id}" idx="{en_idx}"/>' if en_id is not None else ""
    avlst = ('<a:avLst><a:gd name="adj1" fmla="val 50000"/></a:avLst>'
             if preset == "bentConnector3" else "<a:avLst/>")
    anchor = (
        f'<{p}absoluteAnchor><{p}pos x="{ox}" y="{oy}"/><{p}ext cx="{cx}" cy="{cy}"/>'
        f'<{p}cxnSp macro=""><{p}nvCxnSpPr><{p}cNvPr id="{nid}" name="gwexConnector{nid}"/>'
        f'<{p}cNvCxnSpPr><a:cxnSpLocks/>{stx}{enx}</{p}cNvCxnSpPr></{p}nvCxnSpPr>'
        f'<{p}spPr><a:xfrm{flips}><a:off x="{ox}" y="{oy}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="{preset}">{avlst}</a:prstGeom>'
        f'<a:ln w="{w}"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
        f'<a:tailEnd type="triangle"/></a:ln></{p}spPr>'
        f'<{p}style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef>'
        f'<a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>'
        f'<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>'
        f'<a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></{p}style>'
        f'</{p}cxnSp><{p}clientData/></{p}absoluteAnchor>'
    )
    close = f"</{p}wsDr>"
    entries[draw_path] = dx.replace(close, anchor + close).encode("utf-8")
    st_desc = f"id{st_id}/idx{st_idx}" if st_id is not None else f"px{p1}"
    en_desc = f"id{en_id}/idx{en_idx}" if en_id is not None else f"px{p2}"
    return (f"add_connector {sheet} {preset} id={nid} st={st_desc} end={en_desc} "
            f"off=({ox},{oy}) ext=({cx},{cy}){flips}")


def _op_delete_connector(entries: dict[str, bytes], params: dict,
                         default_sheet: Optional[str]) -> str:
    op = "delete_connector"
    sheet = _sheet_of(entries, params, default_sheet, op)
    index = params.get("index")
    ctext = params.get("connected_to_text")
    if (index is None) == (ctext is None):
        raise OpError(f"{op}: index か connected_to_text のどちらか一方を指定してください")
    draw_path, _ = _drawing_of(entries, sheet, op)
    dx = entries[draw_path].decode("utf-8")

    anchors = [am for am in _ANCHOR_RE.finditer(dx)
               if re.search(r"<(?:xdr:)?cxnSp[ >]", am.group(0))
               and not re.search(r"<(?:xdr:)?grpSp[ >]", am.group(0))]

    if index is not None:
        i = int(index)
        if not (1 <= i <= len(anchors)):
            raise OpError(f"{op}: index={i} は範囲外（コネクタは {len(anchors)} 本）")
        targets = [anchors[i - 1]]
        what = f"index={i}"
    else:
        sid = _cnvpr_id_of(_match_sp(dx, str(ctext), op).group(0), op)
        targets = [am for am in anchors
                   if str(sid) in re.findall(r'<a:(?:st|end)Cxn id="(\d+)"', am.group(0))]
        if not targets:
            raise OpError(f"{op}: {ctext!r}（id={sid}）に接続しているコネクタがありません")
        if len(targets) > 1 and not params.get("all"):
            raise OpError(f"{op}: {ctext!r} に接続しているコネクタが {len(targets)} 本あり"
                          "一意に特定できません（all: true で全削除）")
        what = f"connected_to={ctext!r} {len(targets)}本"

    for am in sorted(targets, key=lambda m: m.start(), reverse=True):
        dx = dx[:am.start()] + dx[am.end():]
    entries[draw_path] = dx.encode("utf-8")
    return f"delete_connector {sheet} {what}"


def _png_size(data: bytes) -> Optional[tuple[int, int]]:
    if len(data) >= 24 and data[:8] == b"\x89PNG\r\n\x1a\n":
        return (int.from_bytes(data[16:20], "big"), int.from_bytes(data[20:24], "big"))
    return None


def _op_replace_image(entries: dict[str, bytes], params: dict, default_sheet: Optional[str]) -> str:
    op = "replace_image"
    sheet = _sheet_of(entries, params, default_sheet, op)
    image = str(_require(params, "image", op))
    index = params.get("index")
    name = params.get("name")
    if (index is None) == (name is None):
        raise OpError(f"{op}: index か name のどちらか一方を指定してください")
    if not os.path.isfile(image):
        raise OpError(f"{op}: 画像ファイルがありません: {image}")
    with open(image, "rb") as f:
        new_bytes = f.read()
    new_size = _png_size(new_bytes)
    if new_size is None:
        raise OpError(f"{op}: PNG ではありません（対応は png のみ）: {image}")

    draw_path, draw_rels = _drawing_of(entries, sheet, op)
    dx = entries[draw_path].decode("utf-8")
    pics = []
    for am in _ANCHOR_RE.finditer(dx):
        block = am.group(0)
        if re.search(r"<(?:xdr:)?grpSp[ >]", block) or not re.search(r"<(?:xdr:)?pic[ >]", block):
            continue
        nm = re.search(r'<(?:xdr:)?cNvPr [^>]*?name="([^"]*)"', block)
        rid = re.search(r'r:embed="(rId\d+)"', block)
        pics.append((_nfc(nm.group(1)) if nm else "", rid.group(1) if rid else None))
    if not pics:
        raise OpError(f"{op}: シート {sheet} に画像（pic）がありません")

    if index is not None:
        i = int(index)
        if not (1 <= i <= len(pics)):
            raise OpError(f"{op}: index={i} は範囲外（1〜{len(pics)}）")
        pic_name, rid = pics[i - 1]
    else:
        hits = [pc for pc in pics if pc[0] == _nfc(str(name))]
        if len(hits) != 1:
            raise OpError(f"{op}: name={name!r} の画像が {len(hits)} 件"
                          f"（候補: {[pc[0] for pc in pics]}）")
        pic_name, rid = hits[0]
    if not rid:
        raise OpError(f"{op}: 対象 pic に r:embed がありません")

    media = sheet_scan.media_for_rid(entries, draw_rels, rid)
    if not media or media not in entries or not media.startswith("xl/media/"):
        raise OpError(f"{op}: media を解決できません: rid={rid} → {media}")
    if not media.endswith(".png"):
        raise OpError(f"{op}: 差替できるのは png の media だけです: {media}")

    basename = media.split("/")[-1]
    refs = sum(
        1
        for n, data in entries.items()
        if n.endswith(".rels")
        for tgt in re.findall(r'Target="([^"]+)"', data.decode("utf-8", "replace"))
        if tgt.split("/")[-1] == basename
    )
    embeds_of_rid = len(re.findall(r'r:embed="%s"' % rid, dx))
    if refs > 1 or embeds_of_rid > 1:
        raise OpError(f"{op}: media {media} は複数の図/シートで共有されています"
                      f"（rels 参照 {refs} 件 / この drawing 内 embed {embeds_of_rid} 件）。"
                      "バイト差替は他の画像も変えてしまうため中止")

    old_size = _png_size(entries[media])
    note = ""
    if old_size and old_size != new_size:
        note = f"（注意: ピクセル寸法が {old_size[0]}x{old_size[1]} → {new_size[0]}x{new_size[1]} に変化。表示サイズは不変のため伸縮する）"
    entries[media] = new_bytes
    return f"replace_image {sheet} pic={pic_name!r} media={media}{note}"


def _op_clone_sheet(entries: dict[str, bytes], params: dict, default_sheet: Optional[str]) -> str:
    op = "clone_sheet"
    template = str(_require(params, "template", op))
    new_name = str(_require(params, "new_name", op))
    tab_after = params.get("tab_after")

    wb = entries["xl/workbook.xml"].decode("utf-8")
    if re.search(r'<sheet\b[^>]*name="%s"' % re.escape(escape(new_name, {'"': "&quot;"})), wb):
        raise OpError(f"{op}: シート名が既に存在します: {new_name}")
    try:
        tpl_xml = sheet_scan.sheet_xml_path(entries, template)
    except ValueError as e:
        raise OpError(f"{op}: {e}") from e

    def _next_num(pattern: str) -> int:
        return max((int(m) for n in entries for m in re.findall(pattern, n)), default=0) + 1

    new_num = _next_num(r"^xl/worksheets/sheet(\d+)\.xml$")
    new_sheet = f"xl/worksheets/sheet{new_num}.xml"

    # シート XML の複製。codeName は VBA のシート識別子で重複させられないため
    # 複製からは外す（Excel が開いた時に新しい codeName を採番する）。
    sxd = entries[tpl_xml].decode("utf-8")
    sxd = re.sub(r'(<sheetPr\b[^>]*?)\s+codeName="[^"]*"', r"\1", sxd, count=1)
    entries[new_sheet] = sxd.encode("utf-8")

    dup_notes = []
    tpl_rels = f"xl/worksheets/_rels/{tpl_xml.split('/')[-1]}.rels"
    if tpl_rels in entries:
        rels = entries[tpl_rels].decode("utf-8")
        dm = re.search(r'Target="\.\./drawings/(drawing\d+)\.xml"', rels)
        if dm:
            old_draw = f"xl/drawings/{dm.group(1)}.xml"
            dn = _next_num(r"^xl/drawings/drawing(\d+)\.xml$")
            new_draw = f"xl/drawings/drawing{dn}.xml"
            entries[new_draw] = entries[old_draw]
            old_drels = f"xl/drawings/_rels/{dm.group(1)}.xml.rels"
            if old_drels in entries:
                # drawing rels は同じ media を指してよい（media は読み取り共有）
                entries[f"xl/drawings/_rels/drawing{dn}.xml.rels"] = entries[old_drels]
            rels = rels.replace(f'Target="../drawings/{dm.group(1)}.xml"',
                                f'Target="../drawings/drawing{dn}.xml"')
            _add_ct_override(entries, f"/xl/drawings/drawing{dn}.xml", _DRAWING_CT)
            dup_notes.append(f"drawing→drawing{dn}")
        pm = re.search(r'Target="\.\./printerSettings/(printerSettings\d+)\.bin"', rels)
        if pm:
            pn = _next_num(r"^xl/printerSettings/printerSettings(\d+)\.bin$")
            entries[f"xl/printerSettings/printerSettings{pn}.bin"] = \
                entries[f"xl/printerSettings/{pm.group(1)}.bin"]
            rels = rels.replace(f'Target="../printerSettings/{pm.group(1)}.bin"',
                                f'Target="../printerSettings/printerSettings{pn}.bin"')
            dup_notes.append(f"printerSettings→{pn}")
        entries[f"xl/worksheets/_rels/sheet{new_num}.xml.rels"] = rels.encode("utf-8")

    ct = entries["[Content_Types].xml"].decode("utf-8")
    tm = re.search(r'<Override PartName="/%s" ContentType="([^"]+)"/>' % re.escape(tpl_xml), ct)
    _add_ct_override(entries, f"/{new_sheet}", tm.group(1) if tm else _WORKSHEET_CT)

    rid = xlsx_zip._add_rel(entries, "xl/_rels/workbook.xml.rels", _WORKSHEET_RTYPE,
                            f"worksheets/sheet{new_num}.xml")
    wb = entries["xl/workbook.xml"].decode("utf-8")
    sid = max((int(s) for s in re.findall(r'<sheet\b[^>]*sheetId="(\d+)"', wb)), default=0) + 1
    name_attr = escape(new_name, {'"': "&quot;"})
    # openpyxl 由来の workbook.xml はルートに xmlns:r が無い（各 <sheet> に
    # ローカル宣言する流儀）。無ければ挿入要素にも宣言を付けないと不正 XML になる。
    root_tag = re.match(r"<\?xml[^>]*\?>\s*<workbook\b[^>]*>|<workbook\b[^>]*>", wb)
    r_decl = ""
    if root_tag and "xmlns:r=" not in root_tag.group(0):
        r_decl = ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    sheet_el = f'<sheet{r_decl} name="{name_attr}" sheetId="{sid}" r:id="{rid}"/>'
    if tab_after:
        am = re.search(r'<sheet\b[^>]*name="%s"[^>]*/>' % re.escape(str(tab_after)), wb)
        if not am:
            raise OpError(f"{op}: tab_after のシートがありません: {tab_after}")
        wb = wb[: am.end()] + sheet_el + wb[am.end():]
    else:
        wb = wb.replace("</sheets>", sheet_el + "</sheets>", 1)
    entries["xl/workbook.xml"] = wb.encode("utf-8")
    return (f"clone_sheet {template} → {new_name} "
            f"(sheet{new_num}.xml, sheetId={sid}, {', '.join(dup_notes) or 'rels なし'})")


def _add_ct_override(entries: dict[str, bytes], partname: str, ctype: str) -> None:
    ct = entries["[Content_Types].xml"].decode("utf-8")
    if f'PartName="{partname}"' not in ct:
        ct = ct.replace("</Types>",
                        f'<Override PartName="{partname}" ContentType="{ctype}"/></Types>', 1)
        entries["[Content_Types].xml"] = ct.encode("utf-8")


# ---------------------------------------------------------------------------
# move_shape
#
# 実ファイル（画面設計書 .xlsm）の実測: 図形は twoCellAnchor（from/to の
# col/colOff/row/rowOff）で、spPr に絶対 EMU の <a:xfrm> off/ext キャッシュを持つ。
# セルアンカーの移動は「シートの実際の列幅・行高」で offset を再計算し、
# セル境界を跨いだら col/row を進めて正規化する（Excel と同じ正規形）。
# 列幅・行高→EMU 換算は xlsx_zip._anchor_origin_emu と同じ Excel 実式
# （列: round(文字幅×7)+5 px / 行: pt×12700 EMU）を列・行単位で使う。
# ---------------------------------------------------------------------------

_POS_RE = re.compile(r'<((?:xdr:)?)pos x="(-?\d+)" y="(-?\d+)"/>')
_FROM_TO_TMPL = (
    r"<((?:xdr:)?)%s>\s*"
    r"<(?:xdr:)?col>(\d+)</(?:xdr:)?col>\s*<(?:xdr:)?colOff>(-?\d+)</(?:xdr:)?colOff>\s*"
    r"<(?:xdr:)?row>(\d+)</(?:xdr:)?row>\s*<(?:xdr:)?rowOff>(-?\d+)</(?:xdr:)?rowOff>\s*"
    r"</(?:xdr:)?%s>"
)
_FROM_RE = re.compile(_FROM_TO_TMPL % ("from", "from"), re.DOTALL)
_TO_RE = re.compile(_FROM_TO_TMPL % ("to", "to"), re.DOTALL)


def _cell_geometry(sx: str) -> tuple[Callable[[int], int], Callable[[int], int]]:
    """シート XML → (0始まり列 index → 幅EMU, 0始まり行 index → 高さEMU)。"""
    cw = xlsx_zip._col_widths(sx)
    rh = xlsx_zip._row_heights(sx)
    dc, dr = xlsx_zip._sheet_defaults(sx)

    def col_w(col0: int) -> int:
        w = cw.get(col0 + 1, dc)
        return (round(w * 7.0) + 5) * _EMU if w else round(xlsx_zip._COL_PX * _EMU)

    def row_h(row0: int) -> int:
        h = rh.get(row0 + 1, dr)
        return round(h * sheet_scan.EMU_PER_PT) if h else round(xlsx_zip._ROW_PX * _EMU)

    return col_w, row_h


def _shift_anchor_point(idx: int, off: int, delta: int, size_of: Callable[[int], int],
                        op: str, what: str) -> tuple[int, int]:
    """セルアンカーの1軸 (index, offsetEMU) を delta EMU 動かし、境界跨ぎを正規化する。"""
    off += delta
    while off < 0:
        if idx == 0:
            raise OpError(f"{op}: 移動結果が負座標になります（{what}）")
        idx -= 1
        off += size_of(idx)
    while off >= size_of(idx):
        off -= size_of(idx)
        idx += 1
    return idx, off


def _move_from_to(block: str, regex: re.Pattern, tag: str, dxe: int, dye: int,
                  col_w, row_h, op: str) -> tuple[str, Optional[str]]:
    m = regex.search(block)
    if not m:
        return block, None
    p = m.group(1)
    col, coff, row, roff = (int(m.group(i)) for i in (2, 3, 4, 5))
    ncol, ncoff = _shift_anchor_point(col, coff, dxe, col_w, op, f"{tag} col")
    nrow, nroff = _shift_anchor_point(row, roff, dye, row_h, op, f"{tag} row")
    new = (f"<{p}{tag}><{p}col>{ncol}</{p}col><{p}colOff>{ncoff}</{p}colOff>"
           f"<{p}row>{nrow}</{p}row><{p}rowOff>{nroff}</{p}rowOff></{p}{tag}>")
    return (block[:m.start()] + new + block[m.end():],
            f"{tag} c{col}+{coff}→c{ncol}+{ncoff} r{row}+{roff}→r{nrow}+{nroff}")


def _move_xfrm_off(block: str, dxe: int, dye: int, op: str) -> tuple[str, Optional[str]]:
    """spPr の <a:xfrm> 絶対座標キャッシュがあれば off だけ delta 更新（ext 不変）。"""
    m = re.search(r'<a:xfrm[^>]*><a:off x="(-?\d+)" y="(-?\d+)"/>', block)
    if not m:
        return block, None
    nx, ny = int(m.group(1)) + dxe, int(m.group(2)) + dye
    if nx < 0 or ny < 0:
        raise OpError(f"{op}: 移動結果が負座標になります（xfrm off ({nx},{ny})）")
    new = block[:m.start(1)] + str(nx) + block[m.end(1):m.start(2)] + str(ny) + block[m.end(2):]
    return new, f"xfrm ({m.group(1)},{m.group(2)})→({nx},{ny})"


def _op_move_shape(entries: dict[str, bytes], params: dict, default_sheet: Optional[str]) -> str:
    op = "move_shape"
    sheet = _sheet_of(entries, params, default_sheet, op)
    match_text, sid = params.get("match_text"), params.get("id")
    if (match_text is None) == (sid is None):
        raise OpError(f"{op}: match_text か id のどちらか一方を指定してください")

    def _px(key: str) -> int:
        v = params.get(key, 0)
        if v is None:
            return 0
        if isinstance(v, bool) or not isinstance(v, int):
            raise OpError(f"{op}: {key} は整数(px)で指定してください: {v!r}")
        return v

    dx_px, dy_px = _px("dx_px"), _px("dy_px")
    if dx_px == 0 and dy_px == 0:
        raise OpError(f"{op}: 移動量が 0 です（dx_px / dy_px の少なくとも一方を指定）")

    draw_path, _ = _drawing_of(entries, sheet, op)
    dx = entries[draw_path].decode("utf-8")
    if match_text is not None:
        m = _match_sp(dx, str(match_text), op)
        am = _anchor_of_match(dx, m, op, "対象図形")
        desc = repr(str(match_text))
    else:
        try:
            want = int(sid)
        except (TypeError, ValueError):
            raise OpError(f"{op}: id が整数ではありません: {sid!r}")
        hits = [a for a in _ANCHOR_RE.finditer(dx)
                if str(want) in re.findall(r'<(?:xdr:)?cNvPr [^>]*?id="(\d+)"', a.group(0))]
        if not hits:
            raise OpError(f"{op}: cNvPr id={want} の図形がありません")
        if len(hits) > 1:
            raise OpError(f"{op}: cNvPr id={want} を含むアンカーが {len(hits)} 個あり一意に特定できません")
        am = hits[0]
        desc = f"id={want}"
    block = am.group(0)
    if re.search(r"<(?:xdr:)?grpSp[ >]", block):
        raise OpError(f"{op}: 対象図形はグループ（grpSp）内にあり移動できません: {desc}")
    kind = re.match(r"<(?:xdr:)?(oneCellAnchor|twoCellAnchor|absoluteAnchor)", block).group(1)

    dxe, dye = dx_px * _EMU, dy_px * _EMU
    details: list[str] = []
    if kind == "absoluteAnchor":
        pm = _POS_RE.search(block)
        if not pm:
            raise OpError(f"{op}: absoluteAnchor に pos がありません")
        nx, ny = int(pm.group(2)) + dxe, int(pm.group(3)) + dye
        if nx < 0 or ny < 0:
            raise OpError(f"{op}: 移動結果が負座標になります（pos ({nx},{ny})）")
        block = (block[:pm.start()] + f'<{pm.group(1)}pos x="{nx}" y="{ny}"/>' + block[pm.end():])
        details.append(f"pos ({pm.group(2)},{pm.group(3)})→({nx},{ny})")
    else:
        sheet_xml = sheet_scan.sheet_xml_path(entries, sheet)
        col_w, row_h = _cell_geometry(entries[sheet_xml].decode("utf-8"))
        block, d = _move_from_to(block, _FROM_RE, "from", dxe, dye, col_w, row_h, op)
        if d is None:
            raise OpError(f"{op}: アンカーに from がありません")
        details.append(d)
        if kind == "twoCellAnchor":
            block, d = _move_from_to(block, _TO_RE, "to", dxe, dye, col_w, row_h, op)
            if d is None:
                raise OpError(f"{op}: twoCellAnchor に to がありません")
            details.append(d)
    block, d = _move_xfrm_off(block, dxe, dye, op)
    if d:
        details.append(d)

    entries[draw_path] = (dx[:am.start()] + block + dx[am.end():]).encode("utf-8")
    return (f"move_shape {sheet} {desc} dx_px={dx_px} dy_px={dy_px} "
            f"({kind}: {'; '.join(details)})")


# ---------------------------------------------------------------------------
# add_image（新規 pic の挿入。replace_image は既存 pic の media バイト差替のみ）
# ---------------------------------------------------------------------------

_IMAGE_RTYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"


def _op_add_image(entries: dict[str, bytes], params: dict, default_sheet: Optional[str]) -> str:
    op = "add_image"
    sheet = _sheet_of(entries, params, default_sheet, op)
    image = str(_require(params, "image", op))
    at_px = _require(params, "at_px", op)
    size_px = params.get("size_px")
    if not (isinstance(at_px, (list, tuple)) and len(at_px) == 2):
        raise OpError(f"{op}: at_px は [x, y] の2要素: {at_px!r}")
    if size_px is not None and not (isinstance(size_px, (list, tuple)) and len(size_px) == 2):
        raise OpError(f"{op}: size_px は [w, h] の2要素: {size_px!r}")
    if not os.path.isfile(image):
        raise OpError(f"{op}: 画像ファイルがありません: {image}")
    with open(image, "rb") as f:
        data = f.read()
    nat = _png_size(data)
    if nat is None:
        raise OpError(f"{op}: PNG ではありません（対応は png のみ）: {image}")
    w_px, h_px = (int(size_px[0]), int(size_px[1])) if size_px is not None else nat
    if w_px <= 0 or h_px <= 0:
        raise OpError(f"{op}: size_px は正の数: {size_px!r}")
    x, y = int(at_px[0]) * _EMU, int(at_px[1]) * _EMU
    if x < 0 or y < 0:
        raise OpError(f"{op}: at_px は負にできません: {list(at_px)!r}")
    cx, cy = w_px * _EMU, h_px * _EMU

    draw_path, draw_rels = _drawing_of(entries, sheet, op, create=True)
    midx = xlsx_zip._next_media_index(entries)
    media = f"xl/media/image{midx}.png"
    entries[media] = data
    xlsx_zip._ensure_png_content_type(entries)
    rid = xlsx_zip._add_rel(entries, draw_rels, _IMAGE_RTYPE, f"../media/image{midx}.png")

    dx = entries[draw_path].decode("utf-8")
    nid = _next_cnvpr_id(dx)
    p = "xdr:" if "<xdr:wsDr" in dx else ""
    # 実ファイル（Excel 由来）の wsDr ルートは xmlns:r を宣言しないため、
    # set_cell_image と同じく a: / r: をローカル宣言して自己完結にする。
    a_ns = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    r_ns = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    anchor = (
        f'<{p}absoluteAnchor><{p}pos x="{x}" y="{y}"/><{p}ext cx="{cx}" cy="{cy}"/>'
        f'<{p}pic><{p}nvPicPr><{p}cNvPr id="{nid}" name="Picture {nid}"/>'
        f'<{p}cNvPicPr><a:picLocks {a_ns} noChangeAspect="1"/></{p}cNvPicPr></{p}nvPicPr>'
        f'<{p}blipFill><a:blip {a_ns} {r_ns} r:embed="{rid}"/>'
        f'<a:stretch {a_ns}><a:fillRect/></a:stretch></{p}blipFill>'
        f'<{p}spPr><a:xfrm {a_ns}><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom {a_ns} prst="rect"><a:avLst/></a:prstGeom></{p}spPr>'
        f'</{p}pic><{p}clientData/></{p}absoluteAnchor>'
    )
    close = f"</{p}wsDr>"
    entries[draw_path] = dx.replace(close, anchor + close).encode("utf-8")
    size_note = "" if size_px is not None else "（実画像サイズ）"
    return (f"add_image {sheet} {media} pic='Picture {nid}' rid={rid} "
            f"at_px={[int(at_px[0]), int(at_px[1])]} size_px=[{w_px}, {h_px}]{size_note}")


_OPS: dict[str, Callable] = {
    "set_cell": _op_set_cell,
    "clear_cell": _op_clear_cell,
    "insert_rows": _op_insert_rows,
    "set_merge": _op_set_merge,
    "append_history_row": _op_append_history_row,
    "edit_shape_text": _op_edit_shape_text,
    "add_shape": _op_add_shape,
    "delete_shape": _op_delete_shape,
    "move_shape": _op_move_shape,
    "add_connector": _op_add_connector,
    "delete_connector": _op_delete_connector,
    "replace_image": _op_replace_image,
    "add_image": _op_add_image,
    "clone_sheet": _op_clone_sheet,
}


def apply_ops(path: str, spec: dict, *, output: Optional[str] = None,
              in_place: bool = False) -> dict:
    """spec の ops を順に適用し、全成功＋内部整合 OK のときだけ書き出す。

    返り値は監査レポート dict（ok / output / ops / changed_entries / …）。
    1つでも失敗したら何も書かず ok=False（failed_op と理由つき）。
    """
    if in_place and output:
        raise ValueError("--in-place と -o は同時指定できません")
    if not isinstance(spec, dict):
        return {"ok": False, "written": False, "error": "spec はマッピング（sheet/ops）です"}
    ops = spec.get("ops")
    if not isinstance(ops, list) or not ops:
        return {"ok": False, "written": False, "error": "spec に ops のリストがありません"}

    entries = xlsx_zip._read_zip(path)
    base = dict(entries)
    default_sheet = spec.get("sheet")

    summaries: list[str] = []
    for i, op in enumerate(ops, start=1):
        if not isinstance(op, dict) or len(op) != 1:
            return {"ok": False, "written": False,
                    "failed_op": {"index": i, "op": repr(op)},
                    "error": "各 op は {op名: {…}} の1キー dict です"}
        (name, params), = op.items()
        if name not in _OPS:
            return {"ok": False, "written": False,
                    "failed_op": {"index": i, "op": name},
                    "error": f"未知の op: {name}（対応: {', '.join(sorted(_OPS))}）"}
        try:
            summaries.append(_OPS[name](entries, dict(params or {}), default_sheet))
        except OpError as e:
            return {"ok": False, "written": False,
                    "failed_op": {"index": i, "op": name}, "error": str(e)}
        except Exception as e:  # 予期しない失敗も「何も書かず報告」に倒す
            return {"ok": False, "written": False,
                    "failed_op": {"index": i, "op": name},
                    "error": f"{type(e).__name__}: {e}"}

    ver = xlsm_verify.verify(base, entries)
    if not ver["ok"]:
        return {"ok": False, "written": False,
                "error": "適用後の内部整合チェックで ERROR（書き出し中止）",
                "ops": summaries,
                "verify_issues": ver["issues"]}

    if in_place:
        dest = path
    elif output:
        dest = output
    else:
        stem, ext = os.path.splitext(path)
        dest = f"{stem}_edited{ext}"
    xlsx_zip._write_zip(entries, dest)
    return {
        "ok": True,
        "written": True,
        "output": dest,
        "ops": summaries,
        "changed_entries": ver["changed_entries"],
        "added_entries": ver["added_entries"],
        "verify_warnings": [i for i in ver["issues"] if i["level"] == "WARNING"],
    }
