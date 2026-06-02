/**
 * gwex 画像オーバーレイ Web アプリ（正本）。
 *
 * Google Sheets のセルの「上に」画像を乗せる（in-cell ではなく over-grid 画像）。
 * range 指定時はその範囲を結合して画像を範囲サイズにフィットさせる。
 * insertRows=true で枠ぶんの空行を先に挿入し、既存内容を壊さず確実に枠を作る。
 *
 * gwex（Python）から POST: {spreadsheetId, sheet, cell, range?, insertRows?, dataUrl}
 */
function doPost(e) {
  try {
    var p = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.openById(p.spreadsheetId);
    var sheet = ss.getSheetByName(p.sheet);
    if (!sheet) throw new Error("sheet not found: " + p.sheet);

    var blob = _blobFromDataUrl(p.dataUrl);

    if (p.range) {
      var r = sheet.getRange(p.range);
      var startRow = r.getRow(), startCol = r.getColumn();
      var numRows = r.getNumRows(), numCols = r.getNumColumns();
      if (p.insertRows) {
        sheet.insertRowsBefore(startRow, numRows);
        r = sheet.getRange(startRow, startCol, numRows, numCols);
      }
      _removeImagesInRange(sheet, startRow, startCol, numRows, numCols);
      r.breakApart();
      r.merge();
      var w = 0, h = 0;
      for (var c = 0; c < numCols; c++) w += sheet.getColumnWidth(startCol + c);
      for (var rr = 0; rr < numRows; rr++) h += sheet.getRowHeight(startRow + rr);
      var img = sheet.insertImage(blob, startCol, startRow);
      var ow = img.getWidth(), oh = img.getHeight();
      var fit = Math.min(w / ow, h / oh);
      var scale = fit * (p.scale ? p.scale : 1);   // p.scale=0.8 で枠の80%に縮小
      var nw = Math.max(1, Math.round(ow * scale));
      var nh = Math.max(1, Math.round(oh * scale));
      img.setWidth(nw).setHeight(nh);
      // 枠内で中央寄せ（アンカーセル左上からのオフセット）
      img.setAnchorCellXOffset(Math.max(0, Math.round((w - nw) / 2)));
      img.setAnchorCellYOffset(Math.max(0, Math.round((h - nh) / 2)));
    } else {
      var cell = sheet.getRange(p.cell);
      sheet.insertImage(blob, cell.getColumn(), cell.getRow());
    }
    return _json({ ok: true });
  } catch (err) {
    return _json({ ok: false, error: String(err) });
  }
}

function _removeImagesInRange(sheet, startRow, startCol, numRows, numCols) {
  var imgs = sheet.getImages();
  for (var i = 0; i < imgs.length; i++) {
    var a = imgs[i].getAnchorCell();
    if (a.getRow() >= startRow && a.getRow() <= startRow + numRows - 1 &&
        a.getColumn() >= startCol && a.getColumn() <= startCol + numCols - 1) {
      imgs[i].remove();
    }
  }
}

function _blobFromDataUrl(dataUrl) {
  var m = dataUrl.match(/^data:([^;]+);base64,(.*)$/);
  var mime = m ? m[1] : "image/png";
  var b64 = m ? m[2] : dataUrl;
  return Utilities.newBlob(Utilities.base64Decode(b64), mime, "capture.png");
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}
