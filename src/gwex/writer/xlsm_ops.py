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


_OPS: dict[str, Callable] = {
    "set_cell": _op_set_cell,
    "clear_cell": _op_clear_cell,
    "insert_rows": _op_insert_rows,
    "append_history_row": _op_append_history_row,
    "edit_shape_text": _op_edit_shape_text,
    "add_shape": _op_add_shape,
    "delete_shape": _op_delete_shape,
    "replace_image": _op_replace_image,
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
