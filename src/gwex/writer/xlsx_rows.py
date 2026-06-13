"""ローカル .xlsx の行削除（openpyxl）。privileged。

openpyxl の `Worksheet.delete_rows` はセル値しかシフトせず、結合セル・行の高さ・
条件付き書式・data validation を据え置いたままにする。本モジュールはそれらも
正しく -count 行シフトしながら行削除する。

主用途:
- テンプレ複製時の「指示行（行1）」削除（結合/行高/条件付き書式/dv の -1 シフト込み）
- 未使用サンプル行・ブロックの削除
"""

from __future__ import annotations

from copy import copy
from pathlib import Path
from typing import Optional

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter, range_boundaries


def _shift_range(min_c, min_r, max_c, max_r, start, count):
    """行 [start, start+count-1] を削除したときの結合/書式範囲の新しい行境界を返す。
    完全に削除範囲へ収まる場合は None（範囲消滅）。"""
    end = start + count - 1
    if max_r < start:                      # 完全に上 → 据置
        new_min_r, new_max_r = min_r, max_r
    elif min_r > end:                      # 完全に下 → -count
        new_min_r, new_max_r = min_r - count, max_r - count
    else:                                  # 削除範囲と交差
        upper = start - min_r if min_r < start else 0      # 削除範囲より上に残る行数
        lower = max_r - end if max_r > end else 0          # 削除範囲より下に残る行数
        if upper == 0 and lower == 0:
            return None                    # 全消滅
        new_min_r = min_r if min_r < start else start
        new_max_r = (max_r - count) if max_r > end else (start - 1)
    return new_min_r, new_max_r


def _shift_sqref(sqref, start, count):
    """空白区切りの sqref（複数範囲可）を行削除に合わせてシフト。残った範囲文字列を返す。"""
    parts = []
    for p in str(sqref).split():
        min_c, min_r, max_c, max_r = range_boundaries(p)
        shifted = _shift_range(min_c, min_r, max_c, max_r, start, count)
        if shifted is None:
            continue
        nmin_r, nmax_r = shifted
        parts.append(
            f"{get_column_letter(min_c)}{nmin_r}:{get_column_letter(max_c)}{nmax_r}"
        )
    return " ".join(parts)


def _delete_rows_full(ws, start: int, count: int) -> None:
    """結合/行高/条件付き書式/dv を保ったまま行を削除する。"""
    # 1) 結合セルを退避して全 unmerge
    merged = [(range_boundaries(str(m)), get_column_letter(range_boundaries(str(m))[0]),
               get_column_letter(range_boundaries(str(m))[2])) for m in list(ws.merged_cells.ranges)]
    for m in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(m))

    # 2) 条件付き書式を退避してクリア
    cf_rules = [(rng.sqref, [copy(r) for r in rules])
                for rng, rules in list(ws.conditional_formatting._cf_rules.items())]
    ws.conditional_formatting._cf_rules.clear()

    # 3) data validation を退避してクリア
    dvs = [(str(dv.sqref), dv) for dv in list(ws.data_validations.dataValidation)]
    ws.data_validations.dataValidation.clear()

    # 4) 行の高さを退避
    heights = {r: ws.row_dimensions[r].height
               for r in ws.row_dimensions if ws.row_dimensions[r].height is not None}

    # 5) 値のシフト（openpyxl 標準）
    ws.delete_rows(start, count)

    # 6) 行高を再設定
    end = start + count - 1
    for r in sorted(heights):
        if r < start:
            ws.row_dimensions[r].height = heights[r]
        elif r > end:
            ws.row_dimensions[r - count].height = heights[r]

    # 7) 結合を再 merge
    for (min_c, min_r, max_c, max_r), col_l, col_r in merged:
        shifted = _shift_range(min_c, min_r, max_c, max_r, start, count)
        if shifted is None:
            continue
        nmin_r, nmax_r = shifted
        ws.merge_cells(f"{col_l}{nmin_r}:{col_r}{nmax_r}")

    # 8) 条件付き書式を再登録
    for sqref, rules in cf_rules:
        new_sqref = _shift_sqref(sqref, start, count)
        if not new_sqref:
            continue
        for rule in rules:
            ws.conditional_formatting.add(new_sqref, rule)

    # 9) data validation を再登録
    for sqref, dv in dvs:
        new_sqref = _shift_sqref(sqref, start, count)
        if not new_sqref:
            continue
        dv.sqref = new_sqref
        ws.add_data_validation(dv)


def delete_rows(path: str, sheet: str, start: int, count: int,
                *, output: Optional[str] = None) -> str:
    """指定シートの行 [start, start+count-1] を削除（結合/行高/条件付き書式/dv をシフト）。"""
    if Path(path).suffix.lower() != ".xlsx":
        raise ValueError("delete_rows はローカル .xlsx のみ対応です（Google スプレッドシート不可）。")
    wb = load_workbook(path)
    ws = wb[sheet]
    _delete_rows_full(ws, start, count)
    dest = output or path
    wb.save(dest)
    return dest


def clear_data_region(path: str, sheet: str, start: int, count: int,
                      *, left_col: str = "B", right_col: str = "S",
                      output: Optional[str] = None) -> str:
    """データ領域 [start, start+count-1] × [left_col, right_col] の結合を解除し値を消す。

    write-testspec の前処理用。テンプレのサンプル行に残る結合（E11:E12 等）と
    プレースホルダ（{C11}…）が記入を壊す（結合セルへの書込が左上にリダイレクトされる）ため、
    記入前にこの領域だけを更地化する。罫線・行高・書式は保持する（値と結合のみ操作）。
    """
    if Path(path).suffix.lower() != ".xlsx":
        raise ValueError("clear_data_region はローカル .xlsx のみ対応です。")
    wb = load_workbook(path)
    ws = wb[sheet]
    lc = column_index_from_string(left_col)
    rc = column_index_from_string(right_col)
    end = start + count - 1
    for m in list(ws.merged_cells.ranges):
        mc1, mr1, mc2, mr2 = range_boundaries(str(m))
        if mr2 >= start and mr1 <= end and mc2 >= lc and mc1 <= rc:
            ws.unmerge_cells(str(m))
    for r in range(start, end + 1):
        for c in range(lc, rc + 1):
            ws.cell(r, c).value = None
    dest = output or path
    wb.save(dest)
    return dest


def clear_images(path: str, sheet: str, *, output: Optional[str] = None) -> tuple[str, int]:
    """シートの埋め込み画像を全削除する（set-image の逆。スプシ変換前の画像なし版作成用）。
    返り値: (出力先, 削除した画像数)。"""
    if Path(path).suffix.lower() != ".xlsx":
        raise ValueError("clear_images はローカル .xlsx のみ対応です。")
    wb = load_workbook(path)
    ws = wb[sheet]
    n = len(ws._images)
    ws._images = []
    dest = output or path
    wb.save(dest)
    return dest, n


def extract_sheet(path: str, sheet: str, output: str) -> str:
    """対象シートだけを残した .xlsx を別ファイルに書き出す（Quick Look 目視用。
    qlmanage は先頭シートしか描画しないため、2枚目以降のシートを目視するときに使う）。"""
    if Path(path).suffix.lower() != ".xlsx":
        raise ValueError("extract_sheet はローカル .xlsx のみ対応です。")
    wb = load_workbook(path)
    if sheet not in wb.sheetnames:
        raise ValueError(f"シートがありません: {sheet}")
    for name in [s for s in wb.sheetnames if s != sheet]:
        del wb[name]
    wb.save(output)
    return output


def delete_instruction_row(path: str, sheet: Optional[str] = None,
                           *, marker: str = "{A$1}", output: Optional[str] = None) -> str:
    """行1が指示行（A1 が marker で始まる）のシートで行1を削除する。
    sheet 省略時は指示行を持つ全シートを自動検出して処理する。"""
    if Path(path).suffix.lower() != ".xlsx":
        raise ValueError("delete_instruction_row はローカル .xlsx のみ対応です。")
    wb = load_workbook(path)
    targets = [sheet] if sheet else list(wb.sheetnames)
    done = []
    for name in targets:
        ws = wb[name]
        a1 = ws.cell(1, 1).value
        if isinstance(a1, str) and a1.startswith(marker):
            _delete_rows_full(ws, 1, 1)
            done.append(name)
        elif sheet:  # 明示指定なのに指示行が無い → エラーにせず警告的にスキップ
            raise ValueError(f"シート '{name}' の A1 は指示行（{marker}…）ではありません。")
    dest = output or path
    wb.save(dest)
    return dest, done
