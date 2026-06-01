"""Google parsers の単体テスト（API スキーマに沿った合成 fixture）。

NOTE: 実 API レスポンスでの検証は OAuth 設定後に差し替える。
"""

from __future__ import annotations

from gwex.ir import Divider, Heading, Image, ListBlock, Note, Paragraph, Table
from gwex.parsers import gdocs, gsheet, gslide
from gwex.serializers import markdown as md


# --- Docs -----------------------------------------------------------------


def _para(text, **style):
    return {
        "paragraph": {
            "paragraphStyle": {"namedStyleType": style.get("style", "NORMAL_TEXT")},
            "elements": [{"textRun": {"content": text, "textStyle": style.get("ts", {})}}],
        }
    }


def test_gdocs_heading_paragraph_link_bold():
    doc = {
        "title": "設計メモ",
        "body": {
            "content": [
                _para("概要\n", style="HEADING_1"),
                {
                    "paragraph": {
                        "paragraphStyle": {"namedStyleType": "NORMAL_TEXT"},
                        "elements": [
                            {"textRun": {"content": "詳細は ", "textStyle": {}}},
                            {"textRun": {"content": "こちら", "textStyle": {"link": {"url": "https://x"}}}},
                            {"textRun": {"content": " を参照", "textStyle": {"bold": True}}},
                        ],
                    }
                },
            ]
        },
    }
    out = md.render(gdocs.parse(doc))
    assert out.startswith("# 概要")
    assert "[こちら](https://x)" in out
    assert "** を参照**" in out  # 先頭スペース込みで太字


def test_gdocs_nested_list():
    def bullet(text, level, list_id="L", ordered=False):
        return {
            "paragraph": {
                "bullet": {"listId": list_id, "nestingLevel": level},
                "elements": [{"textRun": {"content": text + "\n", "textStyle": {}}}],
            }
        }

    doc = {
        "title": "t",
        "lists": {"L": {"listProperties": {"nestingLevels": [{}, {}]}}},
        "body": {"content": [bullet("親", 0), bullet("子", 1), bullet("親2", 0)]},
    }
    parsed = gdocs.parse(doc)
    assert isinstance(parsed.blocks[0], ListBlock)
    out = md.render(parsed)
    assert "- 親" in out
    assert "  - 子" in out
    assert "- 親2" in out


def test_gdocs_table_and_image():
    def cell(text):
        return {"content": [_para(text + "\n")]}

    doc = {
        "title": "t",
        "inlineObjects": {
            "io1": {"inlineObjectProperties": {"embeddedObject": {"title": "図", "imageProperties": {"contentUri": "https://img/1"}}}}
        },
        "body": {
            "content": [
                {"table": {"tableRows": [
                    {"tableCells": [cell("項目"), cell("値")]},
                    {"tableCells": [cell("A"), cell("1")]},
                ]}},
                {"paragraph": {"elements": [{"inlineObjectElement": {"inlineObjectId": "io1"}}]}},
            ]
        },
    }
    out = md.render(gdocs.parse(doc))
    assert "| 項目 | 値 |" in out
    assert "| A | 1 |" in out
    assert "![図](https://img/1)" in out


# --- Sheets ----------------------------------------------------------------


def test_gsheet_basic_and_collapse():
    def cv(v):
        return {"formattedValue": v} if v else {}

    ss = {
        "properties": {"title": "Book"},
        "sheets": [
            {
                "properties": {"title": "S1"},
                "data": [{"rowData": [
                    {"values": [cv("名前"), cv(""), cv("得点")]},
                    {"values": [cv(""), cv(""), cv("")]},
                    {"values": [cv("太郎"), cv(""), cv("80")]},
                ]}],
            }
        ],
    }
    out = md.render(gsheet.parse(ss))
    assert "## S1" in out
    assert "| 名前 | 得点 |" in out  # 空列・空行が畳まれている
    assert "| 太郎 | 80 |" in out


def test_gsheet_hyperlink():
    ss = {
        "properties": {"title": "B"},
        "sheets": [{"properties": {"title": "S"}, "data": [{"rowData": [
            {"values": [{"formattedValue": "site", "hyperlink": "https://e"}]},
        ]}]}],
    }
    out = md.render(gsheet.parse(ss))
    assert "[site](https://e)" in out


# --- Slides ----------------------------------------------------------------


def test_gslide_text_table_notes():
    def shape(text, x=0.0, y=0.0):
        return {
            "transform": {"translateX": x, "translateY": y},
            "shape": {"text": {"textElements": [
                {"paragraphMarker": {}},
                {"textRun": {"content": text, "style": {}}},
            ]}},
        }

    pres = {
        "title": "Deck",
        "slides": [
            {
                "pageElements": [shape("本文B", y=200), shape("タイトルA", y=10)],
                "slideProperties": {"notesPage": {"pageElements": [shape("ノート", y=0)]}},
            }
        ],
    }
    parsed = gslide.parse(pres)
    assert isinstance(parsed.blocks[0], Divider)
    assert isinstance(parsed.blocks[1], Heading)
    out = md.render(parsed)
    # y座標でソートされ タイトルA が 本文B より先
    assert out.index("タイトルA") < out.index("本文B")
    assert "> **Note:**" in out
    assert "ノート" in out
