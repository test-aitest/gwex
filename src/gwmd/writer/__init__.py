"""書き戻し層（privileged: ファイル/シートを破壊的に変更する）。

pure 層（parsers/serializers/ir/domains）とは別レイヤ。confined parser とは分離し、
ユーザーの成果物を変更する外向き操作をここに閉じ込める。
"""
