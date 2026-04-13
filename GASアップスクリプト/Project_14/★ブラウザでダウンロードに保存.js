function downloadQoo10SheetAuto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Qoo10');
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Qoo10シートが見つかりません');
    return;
  }
  
  const filename = 'Qoo10_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss') + '.xlsx';
  
  // スプレッドシートをExcel形式でエクスポート
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx&gid=' + sheet.getSheetId();
  
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token }
  });
  
  const blob = response.getBlob().setName(filename);
  const base64 = Utilities.base64Encode(blob.getBytes());
  
  const html = HtmlService.createHtmlOutput(`
    <script>
      const link = document.createElement('a');
      link.href = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${base64}';
      link.download = '${filename}';
      link.click();
      google.script.host.close();
    </script>
    <p>ダウンロードを開始しています...</p>
  `)
  .setWidth(300)
  .setHeight(100);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'ダウンロード中');
}