function 選択確認() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const selection = ss.getActiveRange();
  SpreadsheetApp.getUi().alert(
    '開始行：' + selection.getRow() + '\n' +
    '終了行：' + selection.getLastRow()
  );
}