"""gwex の CLI。

    gwex convert <source> --to md|json
"""

from __future__ import annotations

from typing import Optional

import typer

from gwex import core

app = typer.Typer(add_completion=False, help="文書を Markdown / typed JSON に変換する")


@app.callback()
def _root() -> None:
    """gwex: Google Workspace + Microsoft 文書 → Markdown / typed JSON。"""


@app.command()
def convert(
    source: str = typer.Argument(..., help="Google URL またはローカル OOXML パス"),
    to: str = typer.Option("md", "--to", help="出力形式: md | json"),
    output: Optional[str] = typer.Option(None, "--output", "-o", help="出力先ファイル（省略時は標準出力）"),
    no_sandbox: bool = typer.Option(False, "--no-sandbox", help="解析を sandbox 隔離せず実行（開発用）"),
) -> None:
    result = core.convert(source, to=to, use_sandbox=not no_sandbox)
    if output:
        with open(output, "w", encoding="utf-8") as f:
            f.write(result)
        typer.echo(f"書き出しました: {output}")
    else:
        typer.echo(result)


@app.command()
def share(
    path: str = typer.Argument(..., help="共有したいローカルファイル(.xlsx/.docx/.pptx/.pdf)"),
    convert: bool = typer.Option(False, "--convert", help="Google ネイティブ形式(スプレッドシート/ドキュメント/スライド)に変換して共有"),
    private: bool = typer.Option(False, "--private", help="anyone-with-link 共有をしない（既定は anyone reader）"),
    name: Optional[str] = typer.Option(None, "--name", help="Drive 上の表示名（2026-000xxx_概要設計書_結合テスト項目 等の正式名を渡す）"),
) -> None:
    """Office ファイルを Google Drive にアップロードし、共有 URL を出力する（要 drive.file スコープ）。"""
    from gwex import share as share_mod

    url = share_mod.share_file(path, convert=convert, public=not private, name=name)
    typer.echo(url)


@app.command()
def rename(
    target: str = typer.Argument(..., help="Drive ファイルの URL または ID"),
    name: str = typer.Option(..., "--name", help="新しい表示名"),
) -> None:
    """既存の Drive ファイル(スプシ等)の表示名を変更する（share 後に名前がズレたときの修正）。"""
    from gwex import share as share_mod

    new_name, url = share_mod.rename_file(target, name)
    typer.echo(f"表示名を変更しました: 「{new_name}」 {url}")


@app.command()
def testspec(
    source: str = typer.Argument(..., help="Google URL またはローカル .xlsx パス"),
    sheet: str = typer.Option(..., "--sheet", help="対象シート名（例: 3.結合テスト項目）"),
    to: str = typer.Option("md", "--to", help="出力形式: md | json"),
    mapping: Optional[str] = typer.Option(None, "--mapping", help="列マッピング TOML（省略時は既定）"),
    output: Optional[str] = typer.Option(None, "--output", "-o", help="出力先ファイル（省略時は標準出力）"),
) -> None:
    cfg = core.load_mapping(mapping) if mapping else None
    spec = core.to_testspec(source, sheet, cfg)
    result = core.serialize_testspec(spec, to)
    if output:
        with open(output, "w", encoding="utf-8") as f:
            f.write(result)
        typer.echo(f"書き出しました: {output}")
    else:
        typer.echo(result)


@app.command()
def auth() -> None:
    """OAuth 認証（不足スコープがあれば再同意）。token を ~/.config/gwex に保存。"""
    from gwex.fetcher import auth as auth_mod

    creds = auth_mod.get_credentials()
    typer.echo("認証 OK。付与スコープ:")
    for s in creds.scopes or []:
        typer.echo(f"  - {s}")


@app.command(name="write-testspec")
def write_testspec(
    target: str = typer.Argument(..., help="書き込み先 .xlsx（テンプレ）"),
    sheet: str = typer.Option(..., "--sheet", help="対象シート名"),
    json_path: str = typer.Option(..., "--json", help="TestSpec JSON ファイル"),
    mapping: Optional[str] = typer.Option(None, "--mapping", help="列マッピング TOML（省略時は既定）"),
    append: bool = typer.Option(False, "--append", help="既存データの下に追記（採番は既存グループ最大Noから継続）"),
    dedup: bool = typer.Option(False, "--dedup", help="既存と完全一致するケースを除外（冪等）"),
    fmt: bool = typer.Option(False, "--format", help="追記ブロックに既存体裁（罫線/縦結合/親番号）を後付け（.xlsx）"),
    clear_rows: Optional[int] = typer.Option(None, "--clear-rows", help="記入前にデータ開始行から N 行ぶんの領域（box列）を更地化（テンプレのサンプル結合/プレースホルダ除去）"),
    start_row: Optional[int] = typer.Option(None, "--start-row", help="データ書き込みを開始する行番号（省略時は mapping の data_start_row=10）"),
    output: Optional[str] = typer.Option(None, "--output", "-o", help="出力先 .xlsx（既定 in-place）"),
) -> None:
    cfg = core.load_mapping(mapping) if mapping else None
    with open(json_path, encoding="utf-8") as f:
        spec_json = f.read()
    dest = core.write_testspec(spec_json, target, sheet, cfg, append=append, dedup=dedup, apply_format=fmt, clear_rows=clear_rows, start_row=start_row, output=output)
    typer.echo(f"テスト仕様を書き込みました: {dest}")


@app.command(name="update-case")
def update_case(
    target: str = typer.Argument(..., help="ローカル .xlsx パス または Google スプレッドシート URL"),
    sheet: str = typer.Option(..., "--sheet", help="対象シート名"),
    row: int = typer.Option(..., "--row", help="更新対象の行番号（source_row）"),
    verification: str = typer.Option(..., "--verification", help="確認内容"),
    steps: str = typer.Option("", "--steps", help="実施手順（改行区切り）"),
    mapping: Optional[str] = typer.Option(None, "--mapping", help="列マッピング TOML"),
) -> None:
    cfg = core.load_mapping(mapping) if mapping else None
    step_list = [s for s in steps.split("\n") if s.strip()]
    dest = core.update_case(target, sheet, row, verification, step_list, cfg)
    typer.echo(f"ケースを更新しました（行 {row}）: {dest}")


@app.command(name="strip-instruction-row")
def strip_instruction_row(
    target: str = typer.Argument(..., help="ローカル .xlsx パス"),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="対象シート名（省略時は指示行 {A$1}… を持つ全シートを自動処理）"),
    output: Optional[str] = typer.Option(None, "--output", "-o", help="出力先 .xlsx（既定 in-place）"),
) -> None:
    """テンプレ複製用: 行1（指示行）を削除し、結合/行高/条件付き書式/dv を -1 行シフトする。"""
    from gwex.writer import xlsx_rows

    dest, done = xlsx_rows.delete_instruction_row(target, sheet, output=output)
    typer.echo(f"指示行を削除しました（{', '.join(done) or '対象なし'}）: {dest}")


@app.command(name="copy-row-format")
def copy_row_format(
    src: str = typer.Argument(..., help="書式コピー元 .xlsx パス"),
    dst: str = typer.Argument(..., help="書式コピー先 .xlsx パス"),
    src_sheet: str = typer.Option(..., "--src-sheet", help="コピー元シート名"),
    dst_sheet: str = typer.Option(..., "--dst-sheet", help="コピー先シート名"),
    src_rows: str = typer.Option(..., "--src-rows", help="コピー元行番号（カンマ区切り。例: 6,7,8,9,10）"),
    dst_rows: str = typer.Option(..., "--dst-rows", help="コピー先行番号（カンマ区切り。例: 6,7,8,9,10）"),
    output: Optional[str] = typer.Option(None, "--output", "-o", help="出力先 .xlsx（既定 in-place で dst を上書き）"),
) -> None:
    """ソース xlsx の指定行の書式（行高・セル書式）をターゲット xlsx の指定行にコピーする（値はコピーしない）。"""
    from gwex.writer import xlsx_rows

    s_rows = [int(r.strip()) for r in src_rows.split(",")]
    d_rows = [int(r.strip()) for r in dst_rows.split(",")]
    dest = xlsx_rows.copy_row_format(src, src_sheet, s_rows, dst, dst_sheet, d_rows, output=output)
    typer.echo(f"行書式をコピーしました: {dest}")


@app.command(name="insert-rows")
def insert_rows_cmd(
    target: str = typer.Argument(..., help="ローカル .xlsx パス"),
    sheet: str = typer.Option(..., "--sheet", help="対象シート名"),
    at: int = typer.Option(..., "--at", help="挿入位置（この行の直前に挿入。1始まり）"),
    count: int = typer.Option(1, "--count", help="挿入行数（既定1）"),
    template_row: Optional[int] = typer.Option(None, "--template-row", help="書式コピー元の行番号（セル書式・行高をコピー）"),
    output: Optional[str] = typer.Option(None, "--output", "-o", help="出力先 .xlsx（既定 in-place）"),
) -> None:
    """指定位置に行を挿入し、結合/行高/条件付き書式/dv をシフトする（修正内容の依頼行追加用）。"""
    from gwex.writer import xlsx_rows

    dest = xlsx_rows.insert_rows(target, sheet, at, count, template_row=template_row, output=output)
    typer.echo(f"行を挿入しました（{at} の直前に {count} 行）: {dest}")


@app.command(name="delete-rows")
def delete_rows(
    target: str = typer.Argument(..., help="ローカル .xlsx パス"),
    sheet: str = typer.Option(..., "--sheet", help="対象シート名"),
    start: int = typer.Option(..., "--start", help="削除開始行（1始まり）"),
    count: int = typer.Option(..., "--count", help="削除行数"),
    output: Optional[str] = typer.Option(None, "--output", "-o", help="出力先 .xlsx（既定 in-place）"),
) -> None:
    """指定シートの行を削除し、結合/行高/条件付き書式/dv をシフトする（未使用サンプル行・ブロック削除用）。"""
    from gwex.writer import xlsx_rows

    dest = xlsx_rows.delete_rows(target, sheet, start, count, output=output)
    typer.echo(f"行を削除しました（{start} から {count} 行）: {dest}")


@app.command(name="lint")
def lint(
    target: str = typer.Argument(..., help="走査する .xlsx"),
    sheet: Optional[list[str]] = typer.Option(None, "--sheet", help="対象シート（複数可。省略時は全シート）"),
    ignore: Optional[list[str]] = typer.Option(None, "--ignore", help="除外語（複数可）"),
    to: str = typer.Option("summary", "--to", help="出力: summary | json"),
    output: Optional[str] = typer.Option(None, "--output", "-o", help="出力先ファイル（省略時は標準出力）"),
) -> None:
    """記入し残し（プレースホルダ/サンプル/指示文）を機械検出する。完成前の必須ゲート（ERROR 0 が条件）。"""
    import json as _json

    from gwex.domains import doc_lint

    result = doc_lint.lint(target, sheets=list(sheet) if sheet else None,
                           extra_ignore=list(ignore) if ignore else None)
    text = _json.dumps(result, ensure_ascii=False, indent=1) if to == "json" else doc_lint.summarize(result)
    if output:
        with open(output, "w", encoding="utf-8") as f:
            f.write(text)
        typer.echo(f"書き出しました: {output}（ERROR {len(result['errors'])} / WARN {len(result['warnings'])}）")
    else:
        typer.echo(text)
    if result["errors"]:
        raise typer.Exit(code=1)


@app.command(name="check-images")
def check_images(
    target: str = typer.Argument(..., help="検査する .xlsx"),
    sheet: Optional[list[str]] = typer.Option(None, "--sheet", help="対象シート（複数可。省略時は全シート）"),
    tolerance_px: float = typer.Option(4.0, "--tolerance-px", help="はみ出し判定の許容誤差（px）"),
    aspect_tolerance: float = typer.Option(0.01, "--aspect-tolerance", help="縦横比ずれの許容割合"),
    pair_tolerance_px: float = typer.Option(2.0, "--pair-tolerance-px", help="修正前後の寸法差の許容（px）"),
    to: str = typer.Option("summary", "--to", help="出力: summary | json"),
    output: Optional[str] = typer.Option(None, "--output", "-o", help="出力先ファイル（省略時は標準出力）"),
) -> None:
    """画像の はみ出し/浮遊/二重/前後サイズ違い/縦横比崩れ を数値検出する（目視の置き換え・破綻 0 が条件）。"""
    import json as _json

    from gwex.domains import image_check

    result = image_check.check(target, sheets=list(sheet) if sheet else None,
                               tolerance_px=tolerance_px, aspect_tolerance=aspect_tolerance,
                               pair_tolerance_px=pair_tolerance_px)
    text = _json.dumps(result, ensure_ascii=False, indent=1) if to == "json" else image_check.summarize(result)
    if output:
        with open(output, "w", encoding="utf-8") as f:
            f.write(text)
        typer.echo(f"書き出しました: {output}（破綻 {len(result['issues'])} 件）")
    else:
        typer.echo(text)
    if result["issues"]:
        raise typer.Exit(code=1)


@app.command(name="diff")
def diff(
    a: str = typer.Argument(..., help="比較元 .xlsx（自作）"),
    b: str = typer.Argument(..., help="比較先 .xlsx（answer 等）"),
    sheet: Optional[str] = typer.Option(None, "--sheet", help="対象シート（省略時は共通シート全て）"),
    to: str = typer.Option("summary", "--to", help="出力: summary | json"),
    max_row: Optional[int] = typer.Option(None, "--max-row", help="比較する最大行（省略時は両者の最大行）"),
    max_col: Optional[int] = typer.Option(None, "--max-col", help="比較する最大列（省略時は両者の最大列）"),
    output: Optional[str] = typer.Option(None, "--output", "-o", help="出力先ファイル（省略時は標準出力）"),
) -> None:
    """2つの .xlsx を5観点（値/整列/結合/画像数/行高）で比較する（anken-verify 手順1）。"""
    import json as _json

    from gwex.domains import xlsx_compare

    result = xlsx_compare.compare(a, b, sheet=sheet, max_row=max_row, max_col=max_col)
    text = _json.dumps(result, ensure_ascii=False, indent=1) if to == "json" else xlsx_compare.summarize(result)
    if output:
        with open(output, "w", encoding="utf-8") as f:
            f.write(text)
        typer.echo(f"書き出しました: {output}")
    else:
        typer.echo(text)


@app.command(name="clear-images")
def clear_images(
    target: str = typer.Argument(..., help="ローカル .xlsx パス"),
    sheet: str = typer.Option(..., "--sheet", help="対象シート名"),
    output: Optional[str] = typer.Option(None, "--output", "-o", help="出力先 .xlsx（既定 in-place）"),
) -> None:
    """シートの埋め込み画像を全削除する（set-image の逆。スプシ変換前の画像なし版作成用）。"""
    from gwex.writer import xlsx_rows

    dest, n = xlsx_rows.clear_images(target, sheet, output=output)
    typer.echo(f"画像を {n} 件削除しました: {dest}")


@app.command(name="extract-sheet")
def extract_sheet(
    target: str = typer.Argument(..., help="ローカル .xlsx パス"),
    sheet: str = typer.Option(..., "--sheet", help="残すシート名"),
    output: str = typer.Option(..., "--output", "-o", help="出力先 .xlsx"),
) -> None:
    """対象シートだけの .xlsx を書き出す（qlmanage は先頭シートしか描画しないため、2枚目以降の Quick Look 目視用）。"""
    from gwex.writer import xlsx_rows

    dest = xlsx_rows.extract_sheet(target, sheet, output)
    typer.echo(f"書き出しました: {dest}")


@app.command(name="set-text")
def set_text(
    target: str = typer.Argument(..., help="ローカル .xlsx パス または Google スプレッドシート URL"),
    sheet: str = typer.Option(..., "--sheet", help="対象シート名"),
    cell: str = typer.Option(..., "--cell", help="セル（例: B2）"),
    text: str = typer.Option(..., "--text", help="書き込むテキスト"),
    value_type: Optional[str] = typer.Option(None, "--type", help="値の型: int | float（省略時は文字列）"),
    output: Optional[str] = typer.Option(None, "--output", "-o", help="出力先 .xlsx（xlsx のみ・既定 in-place）"),
) -> None:
    from gwex.writer import base

    dest = base.set_cell_text(target, sheet, cell, text, value_type=value_type, output=output)
    typer.echo(f"書き込みました: {dest}")


@app.command(name="set-image")
def set_image(
    target: str = typer.Argument(..., help="ローカル .xlsx パス または Google スプレッドシート URL"),
    sheet: str = typer.Option(..., "--sheet", help="対象シート名"),
    cell: str = typer.Option(..., "--cell", help="アンカーセル（例: B2）。--range 指定時は左上として扱う"),
    image: str = typer.Option(..., "--image", help="画像ファイルパス"),
    cell_range: Optional[str] = typer.Option(None, "--range", help="画像枠の結合範囲（例: C7:F20）。指定時は結合＋オーバーレイ"),
    insert_rows: bool = typer.Option(False, "--insert-rows", help="枠ぶんの行を挿入して確実に空枠を作る（既存を押し下げ）"),
    width: Optional[int] = typer.Option(None, "--width", help="幅(px)"),
    height: Optional[int] = typer.Option(None, "--height", help="高さ(px)"),
    max_dim: Optional[int] = typer.Option(None, "--max-dim", help="長辺上限px（超過時に自動縮小）"),
    scale: float = typer.Option(1.0, "--scale", help="枠フィット後の倍率（Google: 0.8で枠の80%＝余白付き・中央寄せ）"),
    output: Optional[str] = typer.Option(None, "--output", "-o", help="出力先 .xlsx（xlsx のみ・既定 in-place）"),
) -> None:
    from gwex.writer import base

    dest = base.set_cell_image(
        target, sheet, cell, image,
        cell_range=cell_range, insert_rows=insert_rows, width=width, height=height,
        max_dim=max_dim, scale=scale, output=output,
    )
    typer.echo(f"画像を挿入しました: {dest}")


@app.command(name="set-section")
def set_section(
    target: str = typer.Argument(..., help="ローカル .xlsx パス または Google スプレッドシート URL"),
    sheet: str = typer.Option(..., "--sheet", help="対象シート名（例: 2.画面イメージ(iOS)）"),
    top_row: int = typer.Option(..., "--top-row", help="セクション見出しの行（1始まり）"),
    title: str = typer.Option(..., "--title", help="セクション見出し（画面名など）"),
    before: Optional[str] = typer.Option(None, "--before", help="修正前 画像パス"),
    after: Optional[str] = typer.Option(None, "--after", help="修正後 画像パス"),
    left_cols: str = typer.Option("C,L", "--cols", help="セクション左右端の列（既定 C,L）"),
    split_col: str = typer.Option("H", "--split", help="修正後の開始列（既定 H）"),
    scale: float = typer.Option(0.8, "--scale", help="画像を枠の何倍にするか（既定0.8＝80%・中央）"),
    n_rows: Optional[int] = typer.Option(None, "--n-rows", help="画像エリアの行数を固定（省略時は画像比率から自動計算）"),
    output: Optional[str] = typer.Option(None, "--output", "-o", help="出力先 .xlsx（xlsx のみ・既定 in-place）"),
) -> None:
    """枠付き before/after セクションを作成して画像を配置する（Excel / Google 両対応）。

    見出し＋修正前/修正後ラベル＋青背景＋箱罫線、画像は枠の scale(80%)・比率維持・中央配置。
    --n-rows で行数を固定すると複数ブロックで行数を統一できる。
    --before/--after 両方省略時は枠のみ作成（画像なし・n-rows 省略時は34行）。
    """
    lc = tuple(left_cols.split(","))
    dest = core.create_section(
        target, sheet, top_row, title, before, after,
        left_cols=lc, split_col=split_col, scale=scale, n_rows=n_rows, output=output,
    )
    typer.echo(f"枠付きセクションを作成しました: {dest}")


@app.command(name="set-row-height")
def set_row_height_cmd(
    xlsx: str = typer.Argument(..., help="ローカル .xlsx パス"),
    sheet: str = typer.Option(..., "--sheet", help="対象シート名"),
    row: int = typer.Option(..., "--row", help="行番号（1始まり）"),
    height: float = typer.Option(..., "--height", help="行高（pt）"),
    output: Optional[str] = typer.Option(None, "--output", "-o", help="出力先 .xlsx（既定 in-place）"),
) -> None:
    """指定行の行高を設定する（xlsx のみ対応）。"""
    from gwex.writer import base

    dest = base.set_row_height(xlsx, sheet, row, height, output=output)
    typer.echo(f"行高を設定しました: {dest}")


@app.command(name="set-col-width")
def set_col_width_cmd(
    target: str = typer.Argument(..., help="ローカル .xlsx パス または Google スプレッドシート URL"),
    sheet: str = typer.Option(..., "--sheet", help="対象シート名"),
    col: str = typer.Option(..., "--col", help="列文字（例: C）"),
    width: float = typer.Option(..., "--width", help="列幅。Google スプシは pixel、xlsx は Excel 列幅単位"),
    output: Optional[str] = typer.Option(None, "--output", "-o", help="出力先 .xlsx（xlsx のみ・既定 in-place）"),
) -> None:
    """指定列の幅を設定する（xlsx=Excel列幅単位 / Google=pixel）。画像枠の左右非対称是正用。"""
    from gwex.writer import base

    dest = base.set_col_width(target, sheet, col, width, output=output)
    typer.echo(f"列幅を設定しました ({col}): {dest}")


@app.command(name="set-alignment")
def set_alignment_cmd(
    xlsx: str = typer.Argument(..., help="ローカル .xlsx パス"),
    sheet: str = typer.Option(..., "--sheet", help="対象シート名"),
    cells: list[str] = typer.Option(..., "--cell", help="対象セルまたは範囲（例: C6 / C6:C11）。複数指定可"),
    horizontal: Optional[str] = typer.Option(None, "--horizontal", help="水平位置（left/center/right）"),
    vertical: Optional[str] = typer.Option(None, "--vertical", help="垂直位置（top/center/bottom）"),
    wrap_text: Optional[bool] = typer.Option(None, "--wrap-text", help="折り返し（true/false）"),
    output: Optional[str] = typer.Option(None, "--output", "-o", help="出力先 .xlsx（既定 in-place）"),
) -> None:
    """指定セルのアライメントを部分的に上書きする（未指定の属性は現状維持）。xlsx のみ対応。"""
    from gwex.writer import xlsx_rows

    dest = xlsx_rows.set_alignment(xlsx, sheet, list(cells), horizontal=horizontal, vertical=vertical, wrap_text=wrap_text, output=output)
    typer.echo(f"アライメントを設定しました: {dest}")


@app.command(name="set-merge")
def set_merge(
    xlsx: str = typer.Argument(..., help="ローカル .xlsx パス"),
    sheet: str = typer.Option(..., "--sheet", help="対象シート名"),
    merge_range: list[str] = typer.Option(..., "--range", help="結合範囲（例: B9:B15）。複数指定可"),
    output: Optional[str] = typer.Option(None, "--output", "-o", help="出力先 .xlsx（既定 in-place）"),
) -> None:
    """セル範囲を結合する（既存結合は事前解除してから再結合）。xlsx のみ対応。"""
    from gwex.writer import base

    dest = base.set_merge_cells(xlsx, sheet, list(merge_range), output=output)
    typer.echo(f"結合しました: {dest}")


@app.command(name="autofit-rows")
def autofit_rows_cmd(
    target: str = typer.Argument(..., help="Google スプレッドシート URL"),
    sheet: str = typer.Option(..., "--sheet", help="対象シート名"),
    start_row: int = typer.Option(..., "--start-row", help="開始行（1始まり）"),
    end_row: int = typer.Option(..., "--end-row", help="終了行（1始まり・含む）"),
) -> None:
    """Google スプシの行高をコンテンツに合わせて自動調整する（AutoResizeDimensionsRequest）。

    セルの折り返し（wrap_text）が True の行はテキスト全体が見えるよう自動フィットされる。
    xlsx は Excel がファイルを開くときに customHeight=False の行を自動フィットするので本コマンドは不要。
    """
    from gwex.writer import base

    dest = base.autofit_rows(target, sheet, start_row, end_row)
    typer.echo(f"自動調整しました (row {start_row}〜{end_row}): {dest}")


@app.command(name="export-pdf")
def export_pdf_cmd(
    target: str = typer.Argument(..., help="Google スプレッドシート URL または ID"),
    sheet: list[str] = typer.Option(..., "--sheet", help="対象シート名（複数指定可: --sheet A --sheet B）"),
    out_dir: str = typer.Option("/tmp", "--out-dir", "-o", help="出力ディレクトリ（既定 /tmp）"),
    png: bool = typer.Option(False, "--png", help="PDF を PNG にも変換（sips・Read 目視用）"),
    landscape: bool = typer.Option(False, "--landscape", help="横向きで書き出す（既定は縦向き）"),
) -> None:
    """スプシの指定シートを PDF（任意で PNG）に書き出す（レンダリング目視用・アドホック python 不要）。"""
    from gwex.export_pdf import export_pdf as _export_pdf

    results = _export_pdf(target, list(sheet), out_dir, png=png, portrait=not landscape)
    for r in results:
        if r.get("error"):
            typer.echo(f"[NG] {r['sheet']}: {r['error']}")
        else:
            line = f"PDF: {r['pdf']} ({r['bytes']} bytes)"
            if r.get("png"):
                line += f" / PNG: {r['png']}"
            typer.echo(f"[OK] {r['sheet']} -> {line}")


def main() -> None:
    app()


if __name__ == "__main__":
    main()
