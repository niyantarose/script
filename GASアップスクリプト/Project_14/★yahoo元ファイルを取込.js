function YahooのCSVをインポート() {
  const targetSpreadsheetId = '1B8_o1s7UoAQSyMtm4LVLcm4gF2avGpRTNRD4FyGLaA4';
  const targetSheetName = 'yahoo変換元ファイル';
  const folderId = '1CldavMF4BMmTr4JnYViFWwCqhCux_Xov';
  const janColumnIndex = 33; // AH列は34列目だが、0-indexedで33
  
  Logger.log('=== スクリプト開始 ===');
  try {
    const ss = SpreadsheetApp.openById(targetSpreadsheetId);
    const targetSheet = ss.getSheetByName(targetSheetName);
    if (!targetSheet) {
      throw new Error('シート "' + targetSheetName + '" が見つかりません');
    }
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    let latestFile = null;
    while (files.hasNext()) {
      const f = files.next();
      if (!latestFile || f.getLastUpdated() > latestFile.getLastUpdated()) {
        latestFile = f;
      }
    }
    if (!latestFile) {
      throw new Error('フォルダ内にファイルがありません');
    }
    const mime = latestFile.getMimeType();
    const fileName = latestFile.getName();
    let data = [];
    Logger.log('最新ファイル: ' + fileName + ' (' + mime + ')');
    
    if (mime === MimeType.GOOGLE_SHEETS) {
      // ◆ スプレッドシートの場合
      const srcSs = SpreadsheetApp.openById(latestFile.getId());
      const srcSheet = srcSs.getSheets()[0];
      const lastRow = srcSheet.getLastRow();
      const lastCol = srcSheet.getLastColumn();
      if (lastRow > 0 && lastCol > 0) {
        data = srcSheet.getRange(1, 1, lastRow, lastCol).getValues();
        // JAN列を文字列に変換
        for (let i = 0; i < data.length; i++) {
          if (data[i].length > janColumnIndex && data[i][janColumnIndex]) {
            data[i][janColumnIndex] = String(data[i][janColumnIndex]);
          }
        }
      }
    } else if (mime === MimeType.CSV || mime === 'text/csv') {
      // ◆ CSVの場合：手動パース
      Logger.log('CSVとして読み込み開始');
      let csvContent = latestFile.getBlob().getDataAsString('UTF-8');
      // 文字化けしてたらShift_JISでもう一回
      if (csvContent.indexOf('�') !== -1 || csvContent.indexOf('\ufffd') !== -1) {
        Logger.log('UTF-8で文字化け → Shift_JISで再読込');
        csvContent = latestFile.getBlob().getDataAsString('Shift_JIS');
      }
      Logger.log('CSVパース開始（手動パース）');
      data = parseCSVWithStringColumns(csvContent, [janColumnIndex]);
      Logger.log('CSVパース完了 行数: ' + data.length);
    } else {
      throw new Error('対応していないファイル形式です: ' + mime);
    }
    
    if (data.length === 0) {
      throw new Error('ファイルにデータがありません');
    }
    
    // 転記
    targetSheet.clear();
    targetSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    
    // AH列全体を文字列フォーマットに設定
    if (data.length > 0 && data[0].length > janColumnIndex) {
      Logger.log('JAN列に文字列フォーマットを適用中...');
      targetSheet.getRange(1, janColumnIndex + 1, data.length, 1).setNumberFormat('@STRING@');
    }
    
    targetSheet.setFrozenRows(1);
    Logger.log('書き込み完了: ' + data.length + '行');
    ss.toast('インポート完了：' + data.length + '行（' + fileName + '）', '完了', 5);
  } catch (e) {
    Logger.log('=== エラー発生 ===');
    Logger.log(e.stack);
    SpreadsheetApp.getActiveSpreadsheet().toast('エラー: ' + e.message, 'エラー', 10);
  }
  Logger.log('=== スクリプト終了 ===');
}

/**
 * CSVを手動でパースし、指定された列は文字列として保持する
 * @param {string} csvText - CSVテキスト
 * @param {number[]} stringColumns - 文字列として扱う列のインデックス（0-indexed）
 * @return {Array<Array>} パースされたデータ
 */
function parseCSVWithStringColumns(csvText, stringColumns) {
  const rows = [];
  let currentRow = [];
  let currentField = '';
  let inQuotes = false;
  let columnIndex = 0;
  
  const stringColumnSet = new Set(stringColumns);
  
  for (let i = 0; i < csvText.length; i++) {
    const char = csvText[i];
    const nextChar = csvText[i + 1];
    
    if (char === '"') {
      if (inQuotes && nextChar === '"') {
        currentField += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === ',' && !inQuotes) {
      // フィールドの区切り
      // この列が文字列として保持する列なら、そのまま文字列として保存
      if (stringColumnSet.has(columnIndex)) {
        currentRow.push(currentField);
      } else {
        // 通常の処理（数値変換など）
        currentRow.push(currentField);
      }
      currentField = '';
      columnIndex++;
    } else if ((char === '\n' || (char === '\r' && nextChar === '\n')) && !inQuotes) {
      // 行の区切り
      if (stringColumnSet.has(columnIndex)) {
        currentRow.push(currentField);
      } else {
        currentRow.push(currentField);
      }
      
      if (currentRow.length > 0) {
        rows.push(currentRow);
      }
      currentRow = [];
      currentField = '';
      columnIndex = 0;
      if (char === '\r') i++;
    } else if (char !== '\r') {
      currentField += char;
    }
  }
  
  // 最後のフィールドと行を追加
  if (currentField !== '' || currentRow.length > 0) {
    if (stringColumnSet.has(columnIndex)) {
      currentRow.push(currentField);
    } else {
      currentRow.push(currentField);
    }
    rows.push(currentRow);
  }
  
  return rows;
}