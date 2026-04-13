function downloadQoo10SheetAsExcel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Qoo10価格在庫変更';
  const targetSheet = ss.getSheetByName(sheetName);

  if (!targetSheet) {
    Browser.msgBox('シート「' + sheetName + '」が見つかりません');
    return;
  }

  // ファイル名につける日時フォーマット
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
  const fileName = `Qoo10価格在庫変更_${dateStr}.xlsx`;

  // 1. 一時的なスプレッドシートを新規作成し、対象シートだけをコピー
  const tempSpreadsheet = SpreadsheetApp.create(fileName);
  targetSheet.copyTo(tempSpreadsheet).setName(sheetName);
  
  // 新規作成時にデフォルトで作られる「シート1」を削除
  tempSpreadsheet.getSheetByName('シート1').activate();
  tempSpreadsheet.deleteActiveSheet();

  // 2. 作成した一時ファイルをExcel形式(Blob)として取得
  const tempFileId = tempSpreadsheet.getId();
  const url = "https://docs.google.com/spreadsheets/d/" + tempFileId + "/export?format=xlsx";
  
  const options = {
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  
  // 3. 取得したら、Googleドライブ上の一時ファイルはゴミ箱へ（ゴミを残さない）
  DriveApp.getFileById(tempFileId).setTrashed(true);

  if (response.getResponseCode() !== 200) {
    Browser.msgBox('ファイルの変換に失敗しました。');
    return;
  }

  const blob = response.getBlob();
  const base64 = Utilities.base64Encode(blob.getBytes());

  // 4. ブラウザ側でダウンロードを発動させるHTMLを表示
  const htmlStr = `
    <html>
      <body onload="downloadFile()">
        <div style="text-align: center; font-family: sans-serif; padding: 20px;">
          <p>ダウンロードを開始します...</p>
          <p>自動で開始されない場合は<a id="downloadLink" href="#">こちら</a>をクリックしてください。</p>
        </div>
        <script>
          function downloadFile() {
            const data = "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${base64}";
            const filename = "${fileName}";
            
            const link = document.getElementById('downloadLink');
            link.href = data;
            link.download = filename;
            
            // リンクを自動クリックしてダウンロード開始
            link.click();
            
            // 少し待ってからダイアログを閉じる
            setTimeout(function() {
              google.script.host.close();
            }, 3000);
          }
        </script>
      </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlStr)
    .setWidth(350)
    .setHeight(150);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Excelダウンロード');
}