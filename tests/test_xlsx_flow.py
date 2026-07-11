"""xlsx_flow.draw_flow: 画面遷移図（スクショ+赤枠+矢印+ラベル+処理ボックス）の生成。

すべて absoluteAnchor + spPr の a:xfrm 併記（Excel のアンカークランプ回避と
Quick Look 互換の両立）で描かれることと、レイアウト座標の厳密性を検証する。
"""

from __future__ import annotations

import re
import zipfile
from xml.dom import minidom

import pytest
from openpyxl import load_workbook

from gwex.writer import xlsx_flow

PIL = pytest.importorskip("PIL.Image")

EMU = 9525

# 検算しやすい格子: A(col0,row0)=(10,50), B(col1,row0)=(160,50), 表示 100x200px
GRID = {"node_width": 100, "col_gap": 50, "row_gap": 40,
        "origin_x": 10, "origin_y": 20, "label_gap": 30}


def _spec(tmp_path, edge_extra=None):
    a = tmp_path / "a.png"
    b = tmp_path / "b.png"
    PIL.new("RGB", (100, 200), (0, 0, 255)).save(a)   # 表示スケール 1.0
    PIL.new("RGB", (100, 200), (0, 128, 0)).save(b)
    return {
        "sheet": "Flow",
        "grid": GRID,
        "nodes": [
            {"id": "N1", "name": "画面A", "image": str(a), "col": 0, "row": 0},
            {"id": "N2", "name": "画面B", "image": str(b), "col": 1, "row": 0},
        ],
        "edges": [{"from": "N1", "to": "N2", **(edge_extra or {})}],
    }


def _drawing(path) -> str:
    with zipfile.ZipFile(path) as z:
        names = [n for n in z.namelist() if re.fullmatch(r"xl/drawings/drawing\d+\.xml", n)]
        assert len(names) == 1
        return z.read(names[0]).decode("utf-8")


def test_nodes_pics_and_labels(tmp_path):
    out = str(tmp_path / "flow.xlsx")
    xlsx_flow.draw_flow(_spec(tmp_path), out)

    dx = _drawing(out)
    minidom.parseString(dx)  # well-formed
    load_workbook(out)       # openpyxl 互換

    # 画像2枚が absoluteAnchor + xfrm で正確な位置にある（A=(10,50), B=(160,50)）
    for name, x in [("node_N1", 10), ("node_N2", 160)]:
        m = re.search(
            rf'<xdr:absoluteAnchor><xdr:pos x="(\d+)" y="(\d+)"/><xdr:ext cx="(\d+)" cy="(\d+)"/>'
            rf'<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="\d+" name="{name}"/>', dx)
        assert m, f"{name} が無い"
        assert (int(m.group(1)), int(m.group(2))) == (x * EMU, 50 * EMU)
        assert (int(m.group(3)), int(m.group(4))) == (100 * EMU, 200 * EMU)
        # xfrm（Quick Look 用の絶対座標）も同値
        assert f'<a:off x="{x * EMU}" y="{50 * EMU}"/>' in dx

    # 画面ID/画面名ラベル
    assert "<a:t>画面ID：　N1</a:t>" in dx
    assert "<a:t>画面名：　画面A</a:t>" in dx
    with zipfile.ZipFile(out) as z:
        assert len([n for n in z.namelist() if n.startswith("xl/media/")]) == 2


def test_edge_straight_arrow(tmp_path):
    """trigger 無し: A右端中央(110,150) → B左端中央(160,150)。水平なので直線矢印。"""
    out = str(tmp_path / "flow.xlsx")
    xlsx_flow.draw_flow(_spec(tmp_path), out)
    dx = _drawing(out)
    m = re.search(
        r'<xdr:absoluteAnchor><xdr:pos x="(\d+)" y="(\d+)"/><xdr:ext cx="(\d+)" cy="(\d+)"/>'
        r'<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="\d+" name="edge1"/>', dx)
    assert m
    assert (int(m.group(1)), int(m.group(2))) == (110 * EMU, 150 * EMU)
    assert int(m.group(3)) == 50 * EMU
    assert 'prst="straightConnector1"' in dx
    assert '<a:tailEnd type="triangle"/>' in dx


def test_edge_trigger_rect_and_bent(tmp_path):
    """trigger_rect [10,20,30,40]: 赤枠=(20,70,30,40)、矢印起点=枠右端中央(50,90)→(160,150)でカギ線。"""
    out = str(tmp_path / "flow.xlsx")
    xlsx_flow.draw_flow(_spec(tmp_path, {"trigger_rect": [10, 20, 30, 40]}), out)
    dx = _drawing(out)

    f = re.search(
        r'<xdr:pos x="(\d+)" y="(\d+)"/><xdr:ext cx="(\d+)" cy="(\d+)"/>'
        r'<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="\d+" name="trigger1"/>', dx)
    assert f
    assert [int(g) for g in f.groups()] == [20 * EMU, 70 * EMU, 30 * EMU, 40 * EMU]
    assert re.search(r'name="trigger1"/>.*?<a:srgbClr val="FF0000"/>', dx, re.DOTALL)

    e = re.search(
        r'<xdr:pos x="(\d+)" y="(\d+)"/><xdr:ext cx="(\d+)" cy="(\d+)"/>'
        r'<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="\d+" name="edge1"/>', dx)
    assert e
    assert [int(g) for g in e.groups()] == [50 * EMU, 90 * EMU, 110 * EMU, 60 * EMU]
    assert 'prst="bentConnector3"' in dx


def test_edge_via_box_and_label(tmp_path):
    """via: 中点に処理ボックス+矢印2本。label: 条件ラベルのテキスト図形。"""
    out = str(tmp_path / "flow.xlsx")
    xlsx_flow.draw_flow(_spec(tmp_path, {"via": "実行処理", "label": "成功時"}), out)
    dx = _drawing(out)

    assert "<a:t>実行処理</a:t>" in dx
    assert "<a:t>成功時</a:t>" in dx
    assert re.search(r'name="process1"/>.*?<a:srgbClr val="9DC3E6"/>', dx, re.DOTALL)
    assert 'name="edge1a"' in dx and 'name="edge1b"' in dx
    # ボックス中心 = 起点(110,150)と終点(160,150)の中点(135,150)
    m = re.search(
        r'<xdr:pos x="(\d+)" y="(\d+)"/><xdr:ext cx="(\d+)" cy="(\d+)"/>'
        r'<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="\d+" name="process1"/>', dx)
    px, py, cw, ch = (int(g) for g in m.groups())
    assert px + cw // 2 == 135 * EMU
    assert py + ch // 2 == 150 * EMU


def test_dash_edge_styles(tmp_path):
    """detour エッジは既定で点線、dash: false で実線に戻せる。通常エッジは実線。"""
    out = str(tmp_path / "flow.xlsx")
    xlsx_flow.draw_flow(_spec(tmp_path, {"detour": "閉じる"}), out)
    dx = _drawing(out)
    assert re.search(r'name="edge1"/>.*?<a:prstDash val="dash"/>', dx, re.DOTALL)

    xlsx_flow.draw_flow(_spec(tmp_path, {"detour": "閉じる", "dash": False}), out)
    assert "prstDash" not in _drawing(out)

    xlsx_flow.draw_flow(_spec(tmp_path), out)  # 通常エッジ
    assert "prstDash" not in _drawing(out)


def test_shared_image_media_dedup(tmp_path):
    """同じ画像を複数ノードで使うと media は1エントリに共有される。"""
    spec = _spec(tmp_path)
    spec["nodes"][1]["image"] = spec["nodes"][0]["image"]
    out = str(tmp_path / "flow.xlsx")
    xlsx_flow.draw_flow(spec, out)
    with zipfile.ZipFile(out) as z:
        assert len([n for n in z.namelist() if n.startswith("xl/media/")]) == 1
    assert _drawing(out).count("<xdr:pic>") == 2


def test_empty_nodes_raises(tmp_path):
    with pytest.raises(ValueError, match="nodes"):
        xlsx_flow.draw_flow({"nodes": []}, str(tmp_path / "x.xlsx"))
