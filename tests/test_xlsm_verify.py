"""domains.xlsm_verify: 部分編集後の整合検査（pure・bytes dict 同士の突合せ）。

過去に実際に起きた「row 154 のセルが row 155 の中に紛れる」regex 編集破損
クラスの検出を最重要ケースとして固定する。
"""

from __future__ import annotations

from gwex.domains import xlsm_verify

_WS = "xl/worksheets/sheet1.xml"
_DX = "xl/drawings/drawing1.xml"
_DRELS = "xl/drawings/_rels/drawing1.xml.rels"


def _sheet(rows: str, merges: str = "") -> bytes:
    mc = f'<mergeCells count="{merges.count("<mergeCell")}">{merges}</mergeCells>' if merges else ""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f"<sheetData>{rows}</sheetData>{mc}</worksheet>"
    ).encode("utf-8")


def _kinds(result, level=None):
    return [i["kind"] for i in result["issues"] if level is None or i["level"] == level]


def test_clean_diff_is_ok():
    base = {_WS: _sheet('<row r="1"><c r="A1"><v>1</v></c></row>'), "xl/vbaProject.bin": b"VBA"}
    edited = dict(base)
    edited[_WS] = _sheet('<row r="1"><c r="A1"><v>2</v></c></row>')
    result = xlsm_verify.verify(base, edited)
    assert result["ok"]
    assert result["changed_entries"] == [_WS]
    assert result["issues"] == []


def test_detects_cell_in_wrong_row():
    """fix_sheet39 型破損: row 155 の中に r="A154" のセルが紛れている。"""
    base = {_WS: _sheet('<row r="154"><c r="A154"/></row><row r="155"><c r="A155"/></row>')}
    edited = {_WS: _sheet('<row r="155"><c r="A154"/><c r="A155"/></row>')}
    result = xlsm_verify.verify(base, edited)
    assert not result["ok"]
    assert "cell_row_mismatch" in _kinds(result, "ERROR")
    detail = [i for i in result["issues"] if i["kind"] == "cell_row_mismatch"][0]["detail"]
    assert "A154" in detail and "155" in detail


def test_self_closing_row_no_false_positive():
    """自己終了 <row/> の直後の行のセルを「紛れている」と誤検出しない。

    旧regexは <row .../> で貪欲な [^>]* が末尾 / を食い、次の行の </row> まで
    飲み込んでいた（素の実ファイルで 65〜66件/シートの偽陽性の真因）。
    """
    rows = ('<row r="153"><c r="A153"/></row>'
            '<row r="154" ht="20" customHeight="1"/>'
            '<row r="155"><c r="A155"/><c r="B155"/></row>')
    base = {_WS: _sheet('<row r="1"><c r="A1"><v>1</v></c></row>')}
    result = xlsm_verify.verify(base, {_WS: _sheet(rows)})
    assert result["ok"], result["issues"]
    assert result["issues"] == []


def test_detects_cell_col_order_violation():
    """バグ2の破損クラス: r 属性は正しいのに列順を壊して並ぶセルを ERROR にする。"""
    base = {_WS: _sheet('<row r="155"><c r="A155"/><c r="B155"/></row>')}
    swapped = {_WS: _sheet('<row r="155"><c r="B155"/><c r="A155"/></row>')}
    result = xlsm_verify.verify(base, swapped)
    assert not result["ok"]
    assert "cell_col_order" in _kinds(result, "ERROR")
    detail = [i for i in result["issues"] if i["kind"] == "cell_col_order"][0]["detail"]
    assert "B155" in detail and "A155" in detail

    dup = {_WS: _sheet('<row r="155"><c r="A155"/><c r="A155"/></row>')}
    assert "cell_col_order" in _kinds(xlsm_verify.verify(base, dup), "ERROR")

    # Z の後の AA は昇順（2文字列の列番号換算を誤らない）
    wide = {_WS: _sheet('<row r="155"><c r="Z155"/><c r="AA155"/></row>')}
    assert xlsm_verify.verify(base, wide)["ok"]


def test_detects_injected_cell_breaking_col_order():
    """過去実害の破損形状: 自己終了行が展開されず B154 が row 155 の列順を壊して並ぶ。"""
    broken = {_WS: _sheet('<row r="154"/><row r="155">'
                          '<c r="A155"/><c r="B155"/><c r="B154"/><c r="C155"/></row>')}
    base = {_WS: _sheet('<row r="1"><c r="A1"/></row>')}
    kinds = _kinds(xlsm_verify.verify(base, broken), "ERROR")
    assert "cell_row_mismatch" in kinds
    assert "cell_col_order" in kinds


def test_detects_row_order_and_duplicates():
    base = {_WS: _sheet('<row r="1"><c r="A1"/></row>')}
    out_of_order = {_WS: _sheet('<row r="5"><c r="A5"/></row><row r="3"><c r="A3"/></row>')}
    assert "row_order" in _kinds(xlsm_verify.verify(base, out_of_order), "ERROR")
    dup = {_WS: _sheet('<row r="3"><c r="A3"/></row><row r="3"><c r="B3"/></row>')}
    assert "row_duplicate" in _kinds(xlsm_verify.verify(base, dup), "ERROR")


def test_detects_malformed_xml():
    base = {_WS: _sheet('<row r="1"><c r="A1"/></row>')}
    edited = {_WS: b"<worksheet><sheetData><row r=1></sheetData>"}
    result = xlsm_verify.verify(base, edited)
    assert not result["ok"]
    assert "xml_malformed" in _kinds(result, "ERROR")


def test_detects_merge_problems():
    base = {_WS: _sheet('<row r="1"><c r="A1"/></row>')}
    dup = {_WS: _sheet('<row r="1"><c r="A1"/></row>',
                       merges='<mergeCell ref="A1:B2"/><mergeCell ref="A1:B2"/>')}
    assert "merge_duplicate" in _kinds(xlsm_verify.verify(base, dup), "ERROR")

    inverted = {_WS: _sheet('<row r="1"><c r="A1"/></row>', merges='<mergeCell ref="B5:A2"/>')}
    assert "merge_ref_invalid" in _kinds(xlsm_verify.verify(base, inverted), "ERROR")

    bad_count = dict(base)
    bad_count[_WS] = base[_WS].replace(
        b"</sheetData>",
        b'</sheetData><mergeCells count="3"><mergeCell ref="A1:B2"/></mergeCells>')
    assert "merge_count_mismatch" in _kinds(xlsm_verify.verify(base, bad_count), "ERROR")


def test_vba_must_be_byte_identical():
    base = {"xl/vbaProject.bin": b"VBA", _WS: _sheet("")}
    changed = {"xl/vbaProject.bin": b"vba!", _WS: _sheet("")}
    result = xlsm_verify.verify(base, changed)
    assert not result["ok"] and "vba_changed" in _kinds(result, "ERROR")

    removed = {_WS: _sheet("")}
    result = xlsm_verify.verify(base, removed)
    kinds = _kinds(result, "ERROR")
    assert "vba_removed" in kinds and "entry_removed" in kinds


def test_expect_entries_minimal_touch():
    base = {_WS: _sheet('<row r="1"><c r="A1"><v>1</v></c></row>'),
            "xl/styles.xml": b"<styleSheet/>"}
    edited = {_WS: _sheet('<row r="1"><c r="A1"><v>2</v></c></row>'),
              "xl/styles.xml": b"<styleSheet></styleSheet>"}
    ng = xlsm_verify.verify(base, edited, expect_entries=[_WS])
    assert not ng["ok"]
    assert [i["entry"] for i in ng["issues"] if i["kind"] == "unexpected_change"] == ["xl/styles.xml"]
    ok = xlsm_verify.verify(base, edited, expect_entries=[_WS, "xl/styles.xml"])
    assert ok["ok"]


def _drawing(*shapes: str) -> bytes:
    return (
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        + "".join(shapes) + "</xdr:wsDr>"
    ).encode("utf-8")


def _sp(id_: int) -> str:
    return (f'<xdr:oneCellAnchor><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="{id_}" name="s{id_}"/>'
            f"</xdr:nvSpPr></xdr:sp><xdr:clientData/></xdr:oneCellAnchor>")


def _pic(id_: int, rid: str) -> str:
    return (f'<xdr:oneCellAnchor><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{id_}" name="p{id_}"/>'
            f'</xdr:nvPicPr><xdr:blipFill><a:blip r:embed="{rid}"/></xdr:blipFill></xdr:pic>'
            f"<xdr:clientData/></xdr:oneCellAnchor>")


def test_drawing_shape_decrease_is_warning():
    base = {_DX: _drawing(_sp(1), _sp(2))}
    edited = {_DX: _drawing(_sp(1))}
    result = xlsm_verify.verify(base, edited)
    assert result["ok"]  # WARNING のみ（意図的削除もある）
    assert "shape_count_decreased" in _kinds(result, "WARNING")


def test_drawing_duplicate_cnvpr_id_is_error():
    base = {_DX: _drawing(_sp(1))}
    edited = {_DX: _drawing(_sp(7), _sp(7))}
    result = xlsm_verify.verify(base, edited)
    assert not result["ok"]
    assert "cnvpr_id_duplicate" in _kinds(result, "ERROR")


def test_drawing_embed_resolution():
    rels = (b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            b'<Relationship Id="rId1" Type="t" Target="../media/image1.png"/></Relationships>')
    base = {_DX: _drawing(_sp(1))}

    unresolved = {_DX: _drawing(_pic(2, "rId9")), _DRELS: rels,
                  "xl/media/image1.png": b"PNG"}
    result = xlsm_verify.verify(base, unresolved)
    assert "embed_unresolved" in _kinds(result, "ERROR")

    missing_target = {_DX: _drawing(_pic(2, "rId1")), _DRELS: rels}
    result = xlsm_verify.verify(base, missing_target)
    assert "embed_target_missing" in _kinds(result, "ERROR")

    ok = {_DX: _drawing(_pic(2, "rId1")), _DRELS: rels, "xl/media/image1.png": b"PNG"}
    result = xlsm_verify.verify(base, ok)
    assert not [k for k in _kinds(result, "ERROR") if k.startswith("embed")]


def test_shared_strings_change_is_warning():
    base = {"xl/sharedStrings.xml": b"<sst/>"}
    edited = {"xl/sharedStrings.xml": b"<sst></sst>"}
    result = xlsm_verify.verify(base, edited)
    assert result["ok"]
    assert "shared_strings_changed" in _kinds(result, "WARNING")


def test_summarize_lines():
    base = {_WS: _sheet('<row r="1"><c r="A1"><v>1</v></c></row>')}
    edited = {_WS: _sheet('<row r="2"><c r="A1"><v>1</v></c></row>')}
    text = xlsm_verify.summarize(xlsm_verify.verify(base, edited))
    assert text.startswith("NG")
    assert "cell_row_mismatch" in text
