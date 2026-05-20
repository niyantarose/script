function fixRowHeightOnEdit(e) {
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  const sheetName = sh.getName().trim();

  const targetSheets = [
    '台湾まんが',
    '台湾書籍その他',
    '台湾グッズ'
  ];

  if (!targetSheets.includes(sheetName)) return;

  const row = e.range.getRow();
  const numRows = e.range.getNumRows();
  const lastCol = sh.getLastColumn();

  const fullRowRange = sh.getRange(row, 1, numRows, lastCol);

  // 折り返しを「切り詰める」にする
  fullRowRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  SpreadsheetApp.flush();

  // ほかのonEdit処理の反映を少し待つ
  Utilities.sleep(500);

  // 行高さを強制固定
  sh.setRowHeightsForced(row, numRows, 29);
  SpreadsheetApp.flush();

  console.log(row + "行目から" + numRows + "行を21pxに強制固定しました");
}

function 台湾まんが_行高さを固定する() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('台湾まんが');
  if (!sh) return;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  sh.getRange(2, 1, lastRow - 1, sh.getLastColumn())
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  sh.setRowHeightsForced(2, lastRow - 1, 21);

  console.log('台湾まんがの行高さを21pxに強制固定しました');
}