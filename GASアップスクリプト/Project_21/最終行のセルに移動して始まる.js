function onOpen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  // データの入っている最終行を取得
  const lastRow = sheet.getLastRow();
  
  // 最終行が取得できている場合、次の入力行を自動で選択する
  if (lastRow > 0) {
    // 最終行の「次の行」の「B列（商品コードの列）」を選択する場合
    sheet.getRange(lastRow + 1, 2).activate();
  }
}