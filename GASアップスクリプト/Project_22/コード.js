function onOpen() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // 全データを取得して、下から順にスキャン
  const data = sheet.getDataRange().getValues();
  
  let lastRow = 1;
  for (let i = data.length - 1; i >= 0; i--) {
    // その行のどこかに1つでもデータがあればそこが最終行
    if (data[i].some(cell => cell !== "")) {
      lastRow = i + 1;
      break;
    }
  }
  
  sheet.getRange(lastRow, 1).activate();
}