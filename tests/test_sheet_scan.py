"""domains.sheet_scan: 巨大 xlsm を開かずに1シート構造を JSON 化する「目」。

セクション（B列 ■）/ sparse・full セル / rPh（ふりがな）除去 / 画像・図形の
列挙と counts を、実ファイルと同じ構造のフィクスチャで検証する。
"""

from __future__ import annotations

import unicodedata

import pytest

from gwex.domains import sheet_scan

from tests._xlsm_fixtures import IW, IH, entries_of, make_book

pytest.importorskip("PIL.Image")


def test_sections_and_sparse_cells(tmp_path):
    p = make_book(tmp_path)
    result = sheet_scan.scan(entries_of(p), "N1")

    assert result["xml_path"] == "xl/worksheets/sheet1.xml"
    assert result["sections"] == [
        {"marker": "■画面概要", "row": 7},
        {"marker": "■イベント", "row": 12},
    ]
    refs = {c["ref"]: c for c in result["cells"]}
    assert refs["B8"]["text"] == "ダッシュボード概要"
    assert refs["B8"]["merged"] == "B8:D9"
    assert refs["B22"]["merged"] == "B22:C23"
    assert "C3" not in refs  # 非B列・非結合セルは sparse から除外
    assert "A4" not in refs
    assert result["dims"]["max_row"] >= 22
    assert result["counts"]["merges"] == 2


def test_full_cells_include_numbers(tmp_path):
    p = make_book(tmp_path, with_image=False, with_shapes=False)
    result = sheet_scan.scan(entries_of(p), "N1", cells_mode="full")
    refs = {c["ref"]: c for c in result["cells"]}
    assert refs["C3"]["text"] == "タイトル"
    assert refs["A4"]["text"] == 46000  # シリアル値はそのまま数値

    with pytest.raises(ValueError, match="sparse | full"):
        sheet_scan.scan(entries_of(p), "N1", cells_mode="bogus")
    with pytest.raises(ValueError, match="シートが見つかりません"):
        sheet_scan.scan(entries_of(p), "NOSUCH")


def test_rph_furigana_excluded(tmp_path):
    """sharedStrings の <rPh> ラン・<phoneticPr> はセル文字列に混ぜない。

    フィクスチャの B12 は実ファイルと同じ t="s" 参照で、si に
    <rPh>イベント</rPh> が混在している（_xlsm_fixtures._SST）。
    """
    p = make_book(tmp_path, with_image=False, with_shapes=False)
    entries = entries_of(p)
    assert "<rPh" in entries["xl/sharedStrings.xml"].decode("utf-8")
    assert '<c r="B12" t="s">' in entries["xl/worksheets/sheet1.xml"].decode("utf-8")

    result = sheet_scan.scan(entries, "N1")
    b12 = [c for c in result["cells"] if c["ref"] == "B12"]
    assert b12 and b12[0]["text"] == "■イベント"
    assert result["sections"][1] == {"marker": "■イベント", "row": 12}


def test_images_shapes_and_counts(tmp_path):
    p = make_book(tmp_path)
    result = sheet_scan.scan(entries_of(p), "N1")

    assert result["counts"]["pic"] == 1
    assert result["counts"]["sp"] == 4
    assert result["counts"]["cxnSp"] == 0
    assert result["counts"]["grpSp"] == 0

    (img,) = result["images"]
    assert img["name"] == "capture1"
    assert img["anchor"] == "oneCellAnchor"
    assert img["from"] == [19, 2]  # [row, col] 0始まり（C20 アンカー）
    assert img["media"] == "xl/media/image1.png"
    # ext_pt = EMU/12700（width/height px × 9525 EMU）
    assert img["ext_pt"] == [round(IW * 9525 / 12700), round(IH * 9525 / 12700)]

    texts = {s["text"] for s in result["shapes"]}
    assert {"ボタン変更前", "ラベルA", "重複文言"} <= texts
    label = [s for s in result["shapes"] if s["text"] == "ラベルA"][0]
    assert label["preset"] == "accentBorderCallout2"
    assert label["kind"] == "sp"
    assert label["anchor"] == "twoCellAnchor"
    assert label["from"] == [34, 5]

    # 図形テキストは NFC 正規化されている（NFD では保持しない）
    for s in result["shapes"]:
        assert s["text"] == unicodedata.normalize("NFC", s["text"])


def test_iter_cells_does_not_misattribute_after_self_closing_row():
    """自己終了 <row/> の次行のセルを、自己終了行の row 番号へ誤帰属しない。"""
    sx = ("<worksheet><sheetData>"
          '<row r="154" ht="20" customHeight="1"/>'
          '<row r="155"><c r="A155"><v>1</v></c><c r="B155"/></row>'
          "</sheetData></worksheet>")
    cells = [(row, ref) for row, ref, *_ in sheet_scan.iter_cells(sx)]
    assert cells == [(155, "A155"), (155, "B155")]


def test_summarize_mentions_sections(tmp_path):
    p = make_book(tmp_path)
    text = sheet_scan.summarize(sheet_scan.scan(entries_of(p), "N1"))
    assert "■画面概要@7" in text
    assert "pic=1" in text
