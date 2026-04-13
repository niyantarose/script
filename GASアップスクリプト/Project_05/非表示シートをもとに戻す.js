function 非表示シートを表示する() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  // 全てのシートを確認し、非表示なら表示する
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].isSheetHidden()) {
      sheets[i].showSheet();
    }
  }
}