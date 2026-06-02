"""Google スプレッドシートに before/after セクションを作る（罫線/結合/青背景/中央/画像）。privileged。

Excel の writer/xlsx_format.create_before_after_section と同じ体裁を Google でも再現する:
見出し＋修正前/修正後ラベル＋薄青背景＋箱罫線＋中央仕切り、画像は枠の scale(0.8) で
比率維持・中央配置（Apps Script 経由）。列割りは Excel と同じ 修正前=left..split-1 /
修正後=split..right（隙間なし）。
"""

from __future__ import annotations

from typing import Optional

from openpyxl.utils import column_index_from_string, get_column_letter

HEADER_FILL = "9CC2E5"
AREA_FILL = "DEEAF6"
_ROW_PX = 21


def _rgb(hexstr: str) -> dict:
    return {
        "red": int(hexstr[0:2], 16) / 255,
        "green": int(hexstr[2:4], 16) / 255,
        "blue": int(hexstr[4:6], 16) / 255,
    }


def create_before_after_section(
    target: str,
    sheet: str,
    top_row: int,
    title: str,
    before_image: Optional[str] = None,
    after_image: Optional[str] = None,
    *,
    left_cols: tuple[str, str] = ("C", "L"),
    split_col: str = "H",
    scale: float = 0.8,
) -> str:
    """Google シートに枠付き before/after セクションを構築して画像を配置する。"""
    from PIL import Image as PILImage

    from gwex.writer import gsheet_writer

    svc, _ = gsheet_writer._service()
    ss_id = gsheet_writer._spreadsheet_id(target)
    meta = svc.spreadsheets().get(
        spreadsheetId=ss_id, ranges=[f"'{sheet}'"],
        fields="sheets(properties(sheetId,title),data(columnMetadata))",
    ).execute()
    sh = next(s for s in meta["sheets"] if s["properties"]["title"] == sheet)
    gid = sh["properties"]["sheetId"]
    col_px = [cm.get("pixelSize", 100) for cm in sh["data"][0].get("columnMetadata", [])]

    lc = column_index_from_string(left_cols[0])
    rc = column_index_from_string(left_cols[1])
    sc = column_index_from_string(split_col)
    r_label = top_row + 1
    r_img = top_row + 2

    # 修正前カラム群(lc..sc-1)の px 幅で箱アスペクト=画像アスペクトにし、必要行数を出す
    iw, ih = PILImage.open(before_image or after_image).size
    box_w = sum(col_px[c - 1] for c in range(lc, sc))
    box_h = int(box_w * ih / iw)
    n_rows = max(1, round(box_h / _ROW_PX))
    r_bottom = r_img + n_rows - 1

    def gr(r0, r1, c0, c1):
        return {"sheetId": gid, "startRowIndex": r0 - 1, "endRowIndex": r1,
                "startColumnIndex": c0 - 1, "endColumnIndex": c1}

    def cell_text(r, c, s):
        return {"updateCells": {"rows": [{"values": [{
            "userEnteredValue": {"stringValue": s},
            "userEnteredFormat": {"horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
                                  "textFormat": {"bold": True}}}]}],
            "fields": "userEnteredValue,userEnteredFormat",
            "start": {"sheetId": gid, "rowIndex": r - 1, "columnIndex": c - 1}}}

    solid = {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}}
    reqs = [
        {"updateDimensionProperties": {
            "range": {"sheetId": gid, "dimension": "ROWS", "startIndex": r_img - 1, "endIndex": r_bottom},
            "properties": {"pixelSize": _ROW_PX}, "fields": "pixelSize"}},
        {"mergeCells": {"range": gr(top_row, top_row, lc, rc), "mergeType": "MERGE_ALL"}},
        {"mergeCells": {"range": gr(r_label, r_label, lc, sc - 1), "mergeType": "MERGE_ALL"}},
        {"mergeCells": {"range": gr(r_label, r_label, sc, rc), "mergeType": "MERGE_ALL"}},
        # テキスト＋中央揃え（先に）
        cell_text(top_row, lc, title),
        cell_text(r_label, lc, "修正前"),
        cell_text(r_label, sc, "修正後"),
        # 背景（fields=backgroundColor のみ＝テキスト/配置を壊さない）
        {"repeatCell": {"range": gr(top_row, top_row, lc, rc),
                        "cell": {"userEnteredFormat": {"backgroundColor": _rgb(HEADER_FILL)}},
                        "fields": "userEnteredFormat.backgroundColor"}},
        {"repeatCell": {"range": gr(r_label, r_bottom, lc, rc),
                        "cell": {"userEnteredFormat": {"backgroundColor": _rgb(AREA_FILL)}},
                        "fields": "userEnteredFormat.backgroundColor"}},
        # 箱罫線＋中央仕切り（split の左）
        {"updateBorders": {"range": gr(top_row, r_bottom, lc, rc),
                           "top": solid, "bottom": solid, "left": solid, "right": solid}},
        {"updateBorders": {"range": gr(top_row, r_bottom, sc, sc), "left": solid}},
    ]
    svc.spreadsheets().batchUpdate(spreadsheetId=ss_id, body={"requests": reqs}).execute()

    # 画像（Apps Script で枠フィット×scale・中央寄せ）。範囲は画像エリアそのもの。
    if before_image:
        rng = f"{left_cols[0]}{r_img}:{get_column_letter(sc - 1)}{r_bottom}"
        gsheet_writer.set_cell_image(target, sheet, f"{left_cols[0]}{r_img}", before_image,
                                     cell_range=rng, scale=scale, max_dim=1280)
    if after_image:
        rng = f"{split_col}{r_img}:{left_cols[1]}{r_bottom}"
        gsheet_writer.set_cell_image(target, sheet, f"{split_col}{r_img}", after_image,
                                     cell_range=rng, scale=scale, max_dim=1280)
    return target
