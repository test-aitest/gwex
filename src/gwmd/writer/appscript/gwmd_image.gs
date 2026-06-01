/**
 * gwmd セル内画像挿入 Web アプリ。
 *
 * Google Sheets REST API には画像挿入が無いため、Apps Script の CellImage を使って
 * ネイティブのセル内画像を base64 データURL で埋め込む（公開リンク不要・PII 露出なし）。
 *
 * gwmd（Python）からは POST で {spreadsheetId, sheet, cell, dataUrl, width?, height?} を送る。
 *
 * 注意: in-cell 画像（CellImage）はセルサイズに自動フィットするため、width/height は
 * 行高・列幅の調整用途（必要なら下のコメント解除）。
 */
function doPost(e) {
  try {
    var p = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.openById(p.spreadsheetId);
    var sheet = ss.getSheetByName(p.sheet);
    if (!sheet) throw new Error("sheet not found: " + p.sheet);

    var image = SpreadsheetApp.newCellImage().setSourceUrl(p.dataUrl).build();
    var range = sheet.getRange(p.cell);
    range.setValue(image);

    // 任意: セルの大きさを画像に合わせて広げたい場合
    // if (p.width)  sheet.setColumnWidth(range.getColumn(), p.width);
    // if (p.height) sheet.setRowHeight(range.getRow(), p.height);

    return _json({ ok: true });
  } catch (err) {
    return _json({ ok: false, error: String(err) });
  }
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}
