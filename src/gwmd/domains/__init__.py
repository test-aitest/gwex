"""テスト仕様シートのドメイン特化抽出（pure）。

汎用の表抽出とは別に、結合テスト項目シートを「画面名→(中項目,小項目)→テストケース」の
階層モデル(TestSpec)に写像する。この層は pure: openpyxl をバイト列に対して使うのみで、
ネットワーク I/O も fetcher も import しない（import-linter で強制）。
"""

from gwmd.domains.model import (
    Case,
    Group,
    MappingConfig,
    Screen,
    TestSpec,
    DEFAULT_MAPPING,
)

__all__ = [
    "TestSpec",
    "Screen",
    "Group",
    "Case",
    "MappingConfig",
    "DEFAULT_MAPPING",
]
