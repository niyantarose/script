function Qoo10最新ファイルをインポート_データのみ() {
  const folderId = '1gjwYy17LiJ2CIDEOoIyoj8-h6naG933I';
  const startRow = 5;

  Logger.log('=== Qoo10最新ファイル データインポート開始 ===');

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const destSheet = ss.getSheets()[0];

    // ===== 1) フォルダ内の「最終更新が最新」のファイルを探す（拡張子/MIME問わず） =====
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    let latestFile = null;

    while (files.hasNext()) {
      const f = files.next();
      if (!latestFile || f.getLastUpdated().getTime() > latestFile.getLastUpdated().getTime()) {
        latestFile = f;
      }
    }
    if (!latestFile) throw new Error('フォルダ内にファイルがありません');

    const latestName = latestFile.getName();
    const latestMime = latestFile.getMimeType();
    Logger.log(`最新ファイル: ${latestName} / mime=${latestMime}`);

    // ===== 2) 最新ファイルからデータを取得（GoogleSheets / Excel / CSV に対応） =====
    let sourceValues = null;

    if (latestMime === MimeType.GOOGLE_SHEETS) {
      // ---- Google Sheets ----
      sourceValues = readFromGoogleSheets_(latestFile.getId(), startRow);

    } else if (latestMime === MimeType.MICROSOFT_EXCEL) {
      // ---- Excel (.xlsx/.xls) → 一時的にGoogle Sheetsへ変換して読む ----
      const tempSheetFileId = convertExcelToGoogleSheets_(latestFile.getId(), folderId, latestName);
      try {
        sourceValues = readFromGoogleSheets_(tempSheetFileId, startRow);
      } finally {
        // 変換した一時ファイルを消したくない場合は、この行をコメントアウトしてOK
        DriveApp.getFileById(tempSheetFileId).setTrashed(true);
      }

    } else if (latestMime === MimeType.CSV || latestName.toLowerCase().endsWith('.csv')) {
      // ---- CSV ----
      sourceValues = readFromCsv_(latestFile.getId(), startRow);

    } else {
      throw new Error(`未対応のファイル形式です: ${latestName} / ${latestMime}`);
    }

    if (!sourceValues || sourceValues.length === 0) {
      ss.toast('インポート元にデータがありません（5行目以降）。', '注意', 5);
      return;
    }

    // ===== 3) 貼り付け先の既存データをクリア（書式は残す） =====
    const destLastRow = destSheet.getLastRow();
    const destLastCol = destSheet.getLastColumn();

    if (destLastRow >= startRow) {
      destSheet.getRange(startRow, 1, destLastRow - startRow + 1, destLastCol).clearContent();
    }

    // ===== 4) 値のみ貼り付け =====
    destSheet.getRange(startRow, 1, sourceValues.length, sourceValues[0].length).setValues(sourceValues);

    Logger.log(`インポート完了: ${sourceValues.length}行`);
    ss.toast(`${sourceValues.length}行のデータを更新しました（最新: ${latestName}）`, '完了', 5);

  } catch (e) {
    Logger.log('エラー: ' + e.message);
    SpreadsheetApp.getActiveSpreadsheet().toast('エラー: ' + e.message, '失敗', 10);
  }
}

/**
 * Googleスプレッドシートの1枚目から startRow 以降の値を読む
 */
function readFromGoogleSheets_(fileId, startRow) {
  const sourceSs = SpreadsheetApp.openById(fileId);
  const sourceSheet = sourceSs.getSheets()[0];

  const lastRow = sourceSheet.getLastRow();
  if (lastRow < startRow) return [];

  const numRows = lastRow - startRow + 1;
  const numCols = sourceSheet.getLastColumn();

  return sourceSheet.getRange(startRow, 1, numRows, numCols).getValues();
}

/**
 * ExcelファイルをGoogleスプレッドシートへ変換して返す
 * ※Advanced Drive Service（Drive API）を使う版
 */
function convertExcelToGoogleSheets_(excelFileId, parentFolderId, originalName) {
  // ★重要：GASの「サービス」から「Drive API」をONにして使う
  //（名前が "Drive" の高度なGoogleサービス）

  const resource = {
    title: `[TEMP] ${originalName}`,
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [{ id: parentFolderId }]
  };

  const newFile = Drive.Files.copy(resource, excelFileId);
  return newFile.id;
}

/**
 * CSVを読み込んで配列（2次元）にする
 * startRow は「CSV上の行番号」扱い（1始まり）
 */
function readFromCsv_(fileId, startRow) {
  const blob = DriveApp.getFileById(fileId).getBlob();

  // まず Shift_JIS（Windows-932）を試し、ダメなら UTF-8
  const candidates = ['MS932', 'UTF-8'];

  let text = null;
  for (const enc of candidates) {
    try {
      text = blob.getDataAsString(enc);
      if (text && text.length > 0) break;
    } catch (e) {
      // 次の候補へ
    }
  }
  if (!text) throw new Error('CSVの文字コード判定に失敗しました');

  // Utilities.parseCsv は改行やダブルクォートCSVに対応
  const all = Utilities.parseCsv(text);

  // startRow=5 なら、index は 4 から
  const startIndex = Math.max(0, startRow - 1);
  const sliced = all.slice(startIndex);

  // 空行だけの末尾を削る（軽く掃除）
  while (sliced.length > 0 && sliced[sliced.length - 1].join('').trim() === '') {
    sliced.pop();
  }
  return sliced;
}
