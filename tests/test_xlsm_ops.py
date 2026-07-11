"""writer.xlsm_ops.apply_ops: 宣言的 ops の ZIP 手術適用。

検証の柱:
- 最小タッチ（変更された ZIP エントリ集合の assert・vba/sharedStrings 不変）
- insert_rows 後の下方 図形アンカー・結合セルのシフト
- edit_shape_text の NFC マッチ（NFD 入力でも一致）と 0件/複数件の明示エラー
- 失敗時は何も書かない（原子性）
"""

from __future__ import annotations

import datetime
import re
import unicodedata
import zipfile

import pytest
from openpyxl import load_workbook

from gwex.writer import xlsm_ops

from tests._xlsm_fixtures import entries_of, make_book, rewrite

PIL = pytest.importorskip("PIL.Image")


def _apply(p, tmp_path, ops, sheet="N1"):
    out = str(tmp_path / "out.xlsx")
    report = xlsm_ops.apply_ops(p, {"sheet": sheet, "ops": ops}, output=out)
    return out, report


_ROW11 = '<row r="11"><c r="B11" s="7"><v>1</v></c><c r="D11" s="8"><v>2</v></c></row>'


def _inject_self_closing_row(p):
    """sheet1 に実ファイル形状の自己終了行 r=10 と、直後のセル入り行 r=11 を注入する。"""
    entries = entries_of(p)
    sx = entries["xl/worksheets/sheet1.xml"].decode("utf-8")
    assert '<row r="12"' in sx
    sx = sx.replace('<row r="12"',
                    '<row r="10" ht="20" customHeight="1"/>' + _ROW11 + '<row r="12"', 1)
    entries["xl/worksheets/sheet1.xml"] = sx.encode("utf-8")
    rewrite(p, entries)


def _changed(p, out):
    before, after = entries_of(p), entries_of(out)
    changed = sorted(n for n in before if n in after and before[n] != after[n])
    added = sorted(n for n in after if n not in before)
    return changed, added


def test_set_cell_minimal_touch(tmp_path):
    """変更されるのは対象シート XML ただ1つ。vba・sharedStrings はバイト不変。"""
    p = make_book(tmp_path)
    out, report = _apply(p, tmp_path, [
        {"set_cell": {"cell": "B8", "text": "新しい説明", "mode": "replace"}},
    ])
    assert report["ok"], report
    changed, added = _changed(p, out)
    assert changed == ["xl/worksheets/sheet1.xml"]
    assert added == []
    assert report["changed_entries"] == ["xl/worksheets/sheet1.xml"]

    sx = entries_of(out)["xl/worksheets/sheet1.xml"].decode("utf-8")
    m = re.search(r'<c r="B8"[^>]*t="inlineStr"><is><t[^>]*>([^<]*)</t></is></c>', sx)
    assert m and m.group(1) == "新しい説明"


def test_set_cell_append_joins_with_newline(tmp_path):
    p = make_book(tmp_path)
    out, report = _apply(p, tmp_path, [
        {"set_cell": {"cell": "B8", "text": "追記行", "mode": "append"}},
    ])
    assert report["ok"], report
    sx = entries_of(out)["xl/worksheets/sheet1.xml"].decode("utf-8")
    m = re.search(r'<c r="B8"[^>]*t="inlineStr"><is><t[^>]*>(.*?)</t></is></c>', sx, re.DOTALL)
    assert m.group(1) == "ダッシュボード概要\n追記行"


def test_set_cell_into_self_closing_row_expands_without_touching_next_row(tmp_path):
    """自己終了 <row r="10"/> への書き込みは行を展開し、次行 r=11 を壊さない。

    旧regexは <row .../> の次の </row> まで飲み込み、正しい r 属性のまま
    B10 を row 11 の中に列順を壊して注入していた（過去実害の破損クラス）。
    """
    p = make_book(tmp_path)
    _inject_self_closing_row(p)
    out, report = _apply(p, tmp_path, [
        {"set_cell": {"cell": "B10", "text": "展開された行"}},
    ])
    assert report["ok"], report
    sx = entries_of(out)["xl/worksheets/sheet1.xml"].decode("utf-8")
    assert re.search(
        r'<row r="10" ht="20" customHeight="1"><c r="B10" t="inlineStr">'
        r"<is><t[^>]*>展開された行</t></is></c></row>", sx)
    assert _ROW11 in sx  # 次行はバイト同一のまま


def test_set_cell_insert_keeps_column_order(tmp_path):
    """既存セルがある行への挿入は列順（A,B,…）を保つ（先頭・間・末尾）。"""
    p = make_book(tmp_path)
    _inject_self_closing_row(p)
    out, report = _apply(p, tmp_path, [
        {"set_cell": {"cell": "C11", "text": "間に挿入"}},
        {"set_cell": {"cell": "A11", "number": 1}},
        {"set_cell": {"cell": "E11", "number": 2}},
    ])
    assert report["ok"], report
    sx = entries_of(out)["xl/worksheets/sheet1.xml"].decode("utf-8")
    row11 = re.search(r'<row r="11">.*?</row>', sx).group(0)
    assert re.findall(r'<c r="([A-Z]+11)"', row11) == ["A11", "B11", "C11", "D11", "E11"]


def test_set_cell_number(tmp_path):
    p = make_book(tmp_path)
    _inject_self_closing_row(p)
    out, report = _apply(p, tmp_path, [
        {"set_cell": {"cell": "B11", "number": 46071}},
        {"set_cell": {"cell": "F11", "number": 0.5}},
    ])
    assert report["ok"], report
    sx = entries_of(out)["xl/worksheets/sheet1.xml"].decode("utf-8")
    assert '<c r="B11" s="7"><v>46071</v></c>' in sx  # t なし＝数値・既存 s= 温存
    assert '<c r="F11"><v>0.5</v></c>' in sx

    for bad_params in ({"cell": "B11", "text": "x", "number": 1}, {"cell": "B11"}):
        bad = xlsm_ops.apply_ops(p, {"sheet": "N1", "ops": [{"set_cell": bad_params}]},
                                 output=str(tmp_path / "x.xlsx"))
        assert not bad["ok"] and "どちらか一方" in bad["error"]
    assert not (tmp_path / "x.xlsx").exists()


def test_clear_cell_keeps_style_and_is_noop_when_missing(tmp_path):
    p = make_book(tmp_path)
    _inject_self_closing_row(p)
    out, report = _apply(p, tmp_path, [
        {"clear_cell": {"cell": "B11"}},   # s="7" のセル → 値だけ除去
        {"clear_cell": {"cell": "B8"}},    # スタイル無しの inlineStr セル
        {"clear_cell": {"cell": "Z99"}},   # 存在しないセル → no-op で成功
    ])
    assert report["ok"], report
    assert any("no-op" in s for s in report["ops"])
    sx = entries_of(out)["xl/worksheets/sheet1.xml"].decode("utf-8")
    assert '<c r="B11" s="7"/>' in sx
    assert '<c r="D11" s="8"><v>2</v></c>' in sx  # 隣のセルは不変
    assert re.search(r'<c r="B8"[^>]*/>', sx)
    assert "ダッシュボード概要" not in sx


def test_insert_rows_shifts_merges_and_anchors(tmp_path):
    """挿入位置より下の結合セルと drawing アンカーだけが下へずれる。"""
    p = make_book(tmp_path)
    out, report = _apply(p, tmp_path, [
        {"insert_rows": {"at": 10, "count": 2, "copy_style_from": 12}},
    ])
    assert report["ok"], report

    entries = entries_of(out)
    sx = entries["xl/worksheets/sheet1.xml"].decode("utf-8")
    merges = re.findall(r'<mergeCell ref="([^"]+)"/>', sx)
    assert "B8:D9" in merges       # 挿入位置より上は不変
    assert "B24:C25" in merges     # B22:C23 が +2
    assert "B22:C23" not in merges
    assert re.search(r'<c r="B14"[^>]*t="s"', sx)  # B12(■イベント) が 14 へ
    assert re.search(r'<row r="10"[^>]*>', sx)     # 複製スタイルの空行が挿入された

    dx = entries["xl/drawings/drawing1.xml"].decode("utf-8")
    # pic の oneCellAnchor は row0=19 → 21 へ（0始まり）
    assert "<xdr:row>21</xdr:row>" in dx
    assert "<xdr:row>19</xdr:row>" not in dx

    changed, added = _changed(p, out)
    assert changed == ["xl/drawings/drawing1.xml", "xl/worksheets/sheet1.xml"]
    assert added == []


def test_append_history_row(tmp_path):
    p = make_book(tmp_path)
    out, report = _apply(p, tmp_path, [
        {"append_history_row": {"sheet": "改版履歴", "date": "2026-07-11",
                                "target": "N1", "status": "変更", "content": "文言修正"}},
    ], sheet=None)
    assert report["ok"], report
    serial = (datetime.date(2026, 7, 11) - datetime.date(1899, 12, 30)).days

    sx = entries_of(out)["xl/worksheets/sheet2.xml"].decode("utf-8")
    assert re.search(r'<c r="A6"[^>]*><v>%d</v></c>' % serial, sx)
    for ref, text in (("B6", "N1"), ("C6", "変更"), ("E6", "文言修正")):
        m = re.search(r'<c r="%s"[^>]*t="inlineStr"><is><t[^>]*>([^<]*)</t></is></c>' % ref, sx)
        assert m and m.group(1) == text, ref

    changed, _ = _changed(p, out)
    assert changed == ["xl/worksheets/sheet2.xml"]


def test_edit_shape_text_nfc_match(tmp_path):
    """NFD で渡された match_text でも NFC 正規化で一致する。"""
    p = make_book(tmp_path)
    nfd = unicodedata.normalize("NFD", "ボタン変更前")
    assert nfd != "ボタン変更前"  # 前提: 本当に分解されている
    out, report = _apply(p, tmp_path, [
        {"edit_shape_text": {"match_text": nfd, "new_text": "ボタン変更後"}},
    ])
    assert report["ok"], report
    dx = entries_of(out)["xl/drawings/drawing1.xml"].decode("utf-8")
    assert "<a:t>ボタン変更後</a:t>" in dx
    assert "ボタン変更前" not in dx
    # rPr（書式）は元の run から継承される
    assert re.search(r'<a:r><a:rPr lang="ja-JP" sz="1600" b="1"/><a:t>ボタン変更後</a:t></a:r>', dx)

    changed, _ = _changed(p, out)
    assert changed == ["xl/drawings/drawing1.xml"]


def test_edit_shape_text_zero_or_ambiguous_fails_without_write(tmp_path):
    p = make_book(tmp_path)
    out = str(tmp_path / "out.xlsx")

    report = xlsm_ops.apply_ops(
        p, {"sheet": "N1", "ops": [{"edit_shape_text": {"match_text": "存在しない文言", "new_text": "x"}}]},
        output=out)
    assert not report["ok"]
    assert "一致する図形がありません" in report["error"]
    assert not (tmp_path / "out.xlsx").exists()

    report = xlsm_ops.apply_ops(
        p, {"sheet": "N1", "ops": [{"edit_shape_text": {"match_text": "重複文言", "new_text": "x"}}]},
        output=out)
    assert not report["ok"]
    assert "2 個" in report["error"]
    assert not (tmp_path / "out.xlsx").exists()


def test_failure_mid_ops_writes_nothing(tmp_path):
    """途中の op が失敗したら、先行 op が成功していても何も書かない。"""
    p = make_book(tmp_path)
    out = str(tmp_path / "out.xlsx")
    report = xlsm_ops.apply_ops(p, {"sheet": "N1", "ops": [
        {"set_cell": {"cell": "B8", "text": "先行成功"}},
        {"delete_shape": {"match_text": "存在しない"}},
    ]}, output=out)
    assert not report["ok"]
    assert report["failed_op"] == {"index": 2, "op": "delete_shape"}
    assert not (tmp_path / "out.xlsx").exists()


def test_add_and_delete_shape(tmp_path):
    p = make_book(tmp_path)
    out, report = _apply(p, tmp_path, [
        {"add_shape": {"preset": "wedgeRectCallout", "text": "新ラベル",
                       "at_px": [120, 340], "size_px": [160, 40],
                       "style_from_text": "ラベルA"}},
        {"delete_shape": {"match_text": "ボタン変更前"}},
    ])
    assert report["ok"], report
    dx = entries_of(out)["xl/drawings/drawing1.xml"].decode("utf-8")
    assert "<a:t>新ラベル</a:t>" in dx
    assert "ボタン変更前" not in dx
    # 位置は px→EMU（×9525）。プリセットは複製元と違うので avLst 初期化
    new_sp = re.search(r"<xdr:absoluteAnchor>.*?</xdr:absoluteAnchor>", dx, re.DOTALL).group(0)
    assert f'<xdr:pos x="{120 * 9525}" y="{340 * 9525}"/>' in new_sp
    assert f'<a:ext cx="{160 * 9525}" cy="{40 * 9525}"/>' in new_sp
    assert '<a:prstGeom prst="wedgeRectCallout"><a:avLst/></a:prstGeom>' in new_sp
    # cNvPr id は重複しない
    ids = re.findall(r'<xdr:cNvPr [^>]*?id="(\d+)"', dx)
    assert len(ids) == len(set(ids))

    changed, added = _changed(p, out)
    assert changed == ["xl/drawings/drawing1.xml"]
    assert added == []


def test_replace_image(tmp_path):
    p = make_book(tmp_path)
    new_png = str(tmp_path / "new.png")
    PIL.new("RGB", (120, 260), (255, 0, 0)).save(new_png)
    out, report = _apply(p, tmp_path, [
        {"replace_image": {"index": 1, "image": new_png}},
    ])
    assert report["ok"], report
    changed, added = _changed(p, out)
    assert changed == ["xl/media/image1.png"]
    assert added == []
    assert entries_of(out)["xl/media/image1.png"] == open(new_png, "rb").read()

    # index 範囲外・非PNG は明示エラー
    bad = xlsm_ops.apply_ops(p, {"sheet": "N1", "ops": [
        {"replace_image": {"index": 9, "image": new_png}}]}, output=str(tmp_path / "x.xlsx"))
    assert not bad["ok"] and "範囲外" in bad["error"]


def test_clone_sheet_openpyxl_loadable(tmp_path):
    """複製シートが workbook/rels/Content_Types に配線され、Excel 互換で読める。"""
    p = make_book(tmp_path)
    out, report = _apply(p, tmp_path, [
        {"clone_sheet": {"template": "template", "new_name": "N2", "tab_after": "N1"}},
        {"set_cell": {"sheet": "N2", "cell": "B3", "text": "複製後の追記"}},
    ], sheet=None)
    assert report["ok"], report
    assert "xl/worksheets/sheet4.xml" in report["added_entries"]

    wb = load_workbook(out)
    assert wb.sheetnames == ["N1", "N2", "改版履歴", "template"]
    assert wb["N2"]["B2"].value == "TPLヘッダ"
    assert wb["N2"]["B3"].value == "複製後の追記"

    # 既存名との衝突は明示エラー
    bad = xlsm_ops.apply_ops(p, {"ops": [
        {"clone_sheet": {"template": "template", "new_name": "N1"}}]},
        output=str(tmp_path / "y.xlsx"))
    assert not bad["ok"] and "既に存在" in bad["error"]


def test_clone_sheet_duplicates_drawing(tmp_path):
    p = make_book(tmp_path)
    out, report = _apply(p, tmp_path, [
        {"clone_sheet": {"template": "N1", "new_name": "N1copy"}},
    ], sheet=None)
    assert report["ok"], report
    assert "xl/drawings/drawing2.xml" in report["added_entries"]
    assert "xl/drawings/_rels/drawing2.xml.rels" in report["added_entries"]
    wb = load_workbook(out)
    assert len(wb["N1copy"]._images) == 1  # 複製側にも画像が見える
    assert len(wb["N1"]._images) == 1


def test_vba_and_sharedstrings_always_preserved(tmp_path):
    """どの op を通しても vbaProject.bin と sharedStrings はバイト不変。"""
    p = make_book(tmp_path)
    out, report = _apply(p, tmp_path, [
        {"set_cell": {"cell": "B8", "text": "x"}},
        {"edit_shape_text": {"match_text": "ラベルA", "new_text": "ラベルB"}},
        {"append_history_row": {"sheet": "改版履歴", "date": "2026-07-11",
                                "target": "N1", "status": "変更", "content": "c"}},
    ])
    assert report["ok"], report
    before, after = entries_of(p), entries_of(out)
    assert after["xl/vbaProject.bin"] == before["xl/vbaProject.bin"]
    assert after["xl/sharedStrings.xml"] == before["xl/sharedStrings.xml"]


def test_unknown_op_and_bad_spec(tmp_path):
    p = make_book(tmp_path, with_image=False, with_shapes=False)
    out = str(tmp_path / "out.xlsx")
    r = xlsm_ops.apply_ops(p, {"ops": [{"nop": {}}]}, output=out)
    assert not r["ok"] and "未知の op" in r["error"]
    r = xlsm_ops.apply_ops(p, {"ops": []}, output=out)
    assert not r["ok"]
    r = xlsm_ops.apply_ops(p, {"ops": [{"set_cell": {"cell": "B8", "text": "x"}}]}, output=out)
    assert not r["ok"] and "sheet が未指定" in r["error"]
    assert not (tmp_path / "out.xlsx").exists()
