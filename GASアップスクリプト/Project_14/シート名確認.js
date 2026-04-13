function checkSourceSheetNames() {
  const srcSpreadsheetId = '1Hga-KqhARMZANTlhfZXtmJIAO5DbzoCTY1nYKSllQz8';
  const srcSs = SpreadsheetApp.openById(srcSpreadsheetId);
  const sheets = srcSs.getSheets();
  
  Logger.log('元のスプレッドシートのシート一覧:');
  sheets.forEach((sheet, index) => {
    Logger.log((index + 1) + '. ' + sheet.getName());
  });
}