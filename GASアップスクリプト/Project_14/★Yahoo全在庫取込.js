function importYahooStock_超高速版() {
  const srcFolderId = '1oHU2lhOFQP_YcALRnBOq1Ac4ScjNzo4y'; // ここにフォルダIDを指定
  const srcSheetName = 'yahoo全在庫';
  const dst = SpreadsheetApp.getActiveSpreadsheet();
  const dstSheetName = 'Yahoo全在庫';

  Logger.log('インポート開始...');
  
  try {
    // フォルダ内の最新ファイルを取得
    const folder = DriveApp.getFolderById(srcFolderId);
    const files = folder.getFiles();
    
    let latestFile = null;
    let latestDate = new Date(0);
    
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      
      // スプレッドシートまたはCSVファイルのみ対象
      if (fileName.match(/\.(csv|xlsx|xls)$/i) || file.getMimeType() === MimeType.GOOGLE_SHEETS) {
        const modifiedDate = file.getLastUpdated();
        if (modifiedDate > latestDate) {
          latestDate = modifiedDate;
          latestFile = file;
        }
      }
    }
    
    if (!latestFile) {
      throw new Error('フォルダ内に有効なファイルが見つかりません');
    }
    
    Logger.log('取込元ファイル: ' + latestFile.getName());
    Logger.log('最終更新日時: ' + latestDate);
    
    let srcSs;
    
    // CSVの場合は一時的にスプレッドシートに変換
    if (latestFile.getName().match(/\.csv$/i)) {
      const csvData = latestFile.getBlob().getDataAsString('Shift_JIS');
      const tempSs = SpreadsheetApp.create('temp_' + latestFile.getName());
      const tempSheet = tempSs.getSheets()[0];
      
      const rows = Utilities.parseCsv(csvData);
      if (rows.length > 0) {
        tempSheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
      }
      
      srcSs = tempSs;
      
    } else {
      // スプレッドシートの場合
      srcSs = SpreadsheetApp.open(latestFile);
    }
    
    const srcSheet = srcSs.getSheetByName(srcSheetName) || srcSs.getSheets()[0];
    Logger.log('取込元シート: ' + srcSheet.getName());
    
    const lastRow = srcSheet.getLastRow();
    const lastCol = srcSheet.getLastColumn();
    
    if (lastRow === 0 || lastCol === 0) {
      throw new Error('元シートにデータがありません');
    }
    
    // A～D列を取得（E列の前まで）
    const colAtoD = srcSheet.getRange(1, 1, lastRow, 4).getValues();
    
    // G列以降を取得（F列の後から）
    let colGtoEnd = [];
    if (lastCol > 6) {
      colGtoEnd = srcSheet.getRange(1, 7, lastRow, lastCol - 6).getValues();
    }
    
    // 結合
    const filteredValues = [];
    for (let i = 0; i < lastRow; i++) {
      const newRow = colAtoD[i].concat(colGtoEnd[i] || []);
      filteredValues.push(newRow);
    }
    
    Logger.log('データ処理完了: ' + filteredValues.length + '行');

    let dstSheet = dst.getSheetByName(dstSheetName);
    if (!dstSheet) {
      dstSheet = dst.insertSheet(dstSheetName);
    }

    dstSheet.clear();
    
    if (filteredValues.length > 0 && filteredValues[0].length > 0) {
      dstSheet.getRange(1, 1, filteredValues.length, filteredValues[0].length).setValues(filteredValues);
    }
    
    // CSV用の一時ファイルを削除
    if (latestFile.getName().match(/\.csv$/i)) {
      DriveApp.getFileById(srcSs.getId()).setTrashed(true);
    }
    
    Logger.log('インポート完了: ' + filteredValues.length + '行');
    dst.toast('Yahoo全在庫のインポート完了: ' + filteredValues.length + '行\nファイル: ' + latestFile.getName(), '完了', 5);
    
  } catch (e) {
    Logger.log('エラー: ' + e.message);
    dst.toast('エラー: ' + e.message, 'エラー', 5);
  }
}