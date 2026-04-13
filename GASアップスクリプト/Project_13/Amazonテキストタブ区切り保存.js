function downloadAmazonSheetAsTsv() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Amazon価格在庫変更';
  const targetSheet = ss.getSheetByName(sheetName);

  if (!targetSheet) {
    Browser.msgBox('シート「' + sheetName + '」が見つかりません');
    return;
  }

  // 1. シートの全データを「見た目通り」に取得
  // getDisplayValuesを使うことで、日付や数値の表示形式を維持します
  const range = targetSheet.getDataRange();
  const values = range.getDisplayValues();

  // 2. タブ区切りテキスト(TSV)を作成
  let tsvString = "";
  
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    // 各セル内のタブを削除（構造破壊防止）し、タブで結合
    const rowText = row.map(cell => {
      return String(cell).replace(/\t/g, " "); 
    }).join("\t");
    
    tsvString += rowText + "\r\n";
  }

  // 3. ファイル名の作成
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
  const fileName = `Amazon価格在庫変更_${dateStr}.txt`;

  // 4. Blob作成とBase64エンコード
  // UTF-8で作成します
  const blob = Utilities.newBlob(tsvString, 'text/tab-separated-values', fileName);
  const base64 = Utilities.base64Encode(blob.getBytes());

  // 5. ダウンロード用ダイアログを表示
  const htmlStr = `
    <html>
      <body onload="downloadFile()">
        <div style="text-align: center; font-family: sans-serif; padding: 20px;">
          <p>Amazon用ファイル（タブ区切り）を作成しました。</p>
          <p>自動で開始されない場合は<a id="downloadLink" href="#">こちら</a>をクリックしてください。</p>
        </div>
        <script>
          function downloadFile() {
            // テキストファイルとしてダウンロード
            const data = "data:text/plain;charset=utf-8;base64,${base64}";
            const filename = "${fileName}";
            
            const link = document.getElementById('downloadLink');
            link.href = data;
            link.download = filename;
            
            // ダウンロード発動
            link.click();
            
            // ダイアログを閉じる
            setTimeout(function() {
              google.script.host.close();
            }, 3000);
          }
        </script>
      </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlStr)
    .setWidth(400)
    .setHeight(150);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'テキスト（タブ区切り）ダウンロード');
}