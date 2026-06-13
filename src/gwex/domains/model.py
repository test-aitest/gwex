"""テスト仕様の階層データモデル（pydantic）と列マッピング設定。

階層: TestSpec → Screen → Group(medium+small) → Case。
結合セル（画面名・中項目名など）は親として一度だけ表現し、配下に Case をぶら下げる。
"""

from __future__ import annotations

from typing import Optional, Union

from openpyxl.utils import column_index_from_string
from pydantic import BaseModel, Field

# --------------------------------------------------------------------------
# 階層モデル
# --------------------------------------------------------------------------


class Case(BaseModel):
    # test_no 専用列を持たないテンプレもあるため Optional（既定 None）。
    test_no: Optional[Union[int, str]] = None
    verification_content: str
    execution_steps: list[str] = Field(default_factory=list)
    # 抽出時に元シートの行番号(1始まり)を付与。生成・書込時は None（更新先の特定に使う）。
    source_row: Optional[int] = None


class Group(BaseModel):
    medium_category: str = ""
    small_category: str = ""
    cases: list[Case] = Field(default_factory=list)


class Screen(BaseModel):
    screen_name: str = ""
    groups: list[Group] = Field(default_factory=list)


class TestSpec(BaseModel):
    __test__ = False  # pytest にテストクラスとして収集させない

    sheet: str
    screens: list[Screen] = Field(default_factory=list)


# --------------------------------------------------------------------------
# 列マッピング設定
# --------------------------------------------------------------------------

ROLES = (
    "screen_name",
    "medium_category",
    "small_category",
    "test_no",
    "verification_content",
    "execution_steps",
)


class MappingConfig(BaseModel):
    """シートのどの列がどのロールかを表す設定。

    columns の値は「列レター(例 'C')」「列インデックス(1始まり int)」
    「ヘッダラベル(文字列)」のいずれか。ラベルは header_row から解決する。
    """

    header_row: int = 7          # 1 始まり。ヘッダ行（ラベル解決と検証に使用）
    data_start_row: int = 10     # 1 始まり。データ開始行
    columns: dict[str, Union[str, int]] = Field(default_factory=dict)
    # 整形（罫線/結合/番号）用。No 列（画面/中項目/小項目の番号列）と枠の左右端。
    # 未設定なら整形は最小限（番号なし）。box_right_col 未設定時は値列の最右で代用。
    number_columns: dict[str, str] = Field(default_factory=dict)
    box_left_col: Optional[str] = None
    box_right_col: Optional[str] = None
    fill: list[str] = Field(
        default_factory=lambda: [
            "screen_name",
            "medium_category",
            "small_category",
            "execution_steps",
        ]
    )
    steps_split: str = "newline+marker"  # or "newline"

    def column_index(self, role: str, header: Optional[list] = None) -> Optional[int]:
        """ロールの 0 始まり列インデックスを返す。未定義なら None。

        - 列レター 'C' → 2
        - int(1始まり) 3 → 2
        - ラベル '確認内容' → header 行から探す
        """
        spec = self.columns.get(role)
        if spec is None:
            return None
        if isinstance(spec, int):
            return spec - 1
        s = str(spec)
        if s.isalpha():
            return column_index_from_string(s.upper()) - 1
        if s.isdigit():
            return int(s) - 1
        if header is not None:  # ラベル一致
            for i, cell in enumerate(header):
                if cell is not None and str(cell).strip() == s:
                    return i
        return None


# 実ファイル「3.結合テスト項目」テンプレに合わせた既定マッピング。
# 各論理列は「No 列 + 値列」の 2 列構成のため、値列（レター）を採用する。
DEFAULT_MAPPING = MappingConfig(
    header_row=7,
    data_start_row=10,
    # 実テンプレは test_no 専用列を持たず、H=確認内容 / I=実施手順。
    columns={
        "screen_name": "C",        # B=大項目No, C=大項目名
        "medium_category": "E",    # D=中項目No, E=中項目名
        "small_category": "G",     # F=小項目No, G=小項目名
        "verification_content": "H",
        "execution_steps": "I",
    },
    # 画面/中項目/小項目の No 列（値列の左隣）。整形時に親番号を採番・縦結合する。
    number_columns={
        "screen_name": "B",
        "medium_category": "D",
        "small_category": "F",
    },
    # 枠（罫線）の左右端。B〜S が 1 ケース行の箱（実施日/結果/備考まで含む）。
    box_left_col="B",
    box_right_col="S",
)
