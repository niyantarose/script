/******************************************************
 * ★Yahoo在庫変更シートをCSVでダウンロード（Shift-JIS）
 * 
 * 使い方：
 * 1. メニューから実行、またはスクリプトエディタから直接実行
 * 2. ダイアログにダウンロードリンクが表示される
 * 3. リンクをクリックしてCSVをダウンロード
 ******************************************************/

const STOCK_CSV_SHEET_NAME = '★Yahoo在庫変更';

/**
 * メニュー追加（既存のonOpenから呼び出してください）
 */
function STOCKCSV_onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📥 在庫CSV')
    .addItem('★Yahoo在庫変更をCSV保存', 'Yahoo在庫変更シートCSVダウンロード')
    .addToUi();
}

/**
 * ★Yahoo在庫変更シートをCSVでダウンロード
 */
function Yahoo在庫変更シートCSVダウンロード() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(STOCK_CSV_SHEET_NAME);
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('エラー', `「${STOCK_CSV_SHEET_NAME}」が見つかりません`, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // データ取得
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < 1 || lastCol < 1) {
    SpreadsheetApp.getUi().alert('エラー', 'シートにデータがありません', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  
  // CSV形式に変換
  const csv = convertToCSV_Stock_(data);
  
  // ファイル名（日時付き）
  const now = new Date();
  const timestamp = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  const fileName = `Yahoo在庫変更_${timestamp}.csv`;
  
  // ダウンロードダイアログを表示
  showDownloadDialog_Stock_(csv, fileName);
}

/**
 * 2次元配列をCSV文字列に変換
 */
function convertToCSV_Stock_(data) {
  return data.map(row => {
    return row.map(cell => {
      // nullやundefinedは空文字に
      if (cell === null || cell === undefined) {
        return '';
      }
      
      // 日付オブジェクトは文字列に変換
      if (cell instanceof Date) {
        return Utilities.formatDate(cell, 'Asia/Tokyo', 'yyyy/MM/dd');
      }
      
      // 文字列に変換
      let str = String(cell);
      
      // カンマ、改行、ダブルクォートを含む場合はダブルクォートで囲む
      if (str.includes(',') || str.includes('\n') || str.includes('"')) {
        str = '"' + str.replace(/"/g, '""') + '"';
      }
      
      return str;
    }).join(',');
  }).join('\r\n');
}

/**
 * ダウンロードダイアログを表示（Shift-JIS版）
 */
function showDownloadDialog_Stock_(csvContent, fileName) {
  // Shift_JISでエンコード（Yahoo在庫管理はSJIS必須）
  const sjisBlob = Utilities.newBlob('').setDataFromString(csvContent, 'Shift_JIS');
  const base64 = Utilities.base64Encode(sjisBlob.getBytes());
  
  // HTMLでダウンロードリンクを作成
  const html = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      padding: 20px;
      text-align: center;
    }
    .download-btn {
      display: inline-block;
      padding: 15px 30px;
      background-color: #ff6b35;
      color: white;
      text-decoration: none;
      border-radius: 5px;
      font-size: 16px;
      margin: 20px 0;
      cursor: pointer;
    }
    .download-btn:hover {
      background-color: #e55a2b;
    }
    .info {
      color: #666;
      font-size: 14px;
      margin-top: 15px;
    }
    .filename {
      font-weight: bold;
      color: #333;
    }
    .encoding {
      color: #007bff;
      font-size: 12px;
      margin-top: 5px;
    }
  </style>
</head>
<body>
  <h3>📥 在庫CSV ダウンロード準備完了</h3>
  <p class="filename">${fileName}</p>
  <p class="encoding">文字コード: Shift_JIS</p>
  <a class="download-btn" 
     href="data:text/csv;charset=shift_jis;base64,${base64}" 
     download="${fileName}">
    ダウンロード
  </a>
  <p class="info">ボタンをクリックするとダウンロードが始まります<br>在庫管理ページでアップロードしてください</p>
</body>
</html>
`;
  
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(380)
    .setHeight(280);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '在庫CSVダウンロード');
}

/*
 * ★ 既存の onOpen() に以下の1行を追加してください：
 *    STOCKCSV_onOpen();
 */