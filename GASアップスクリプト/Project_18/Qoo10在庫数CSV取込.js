/**
 * Qoo10 在庫ファイル（CSV/XLSX）を指定フォルダから最新ファイルを取得し、
 * Google スプレッドシートの指定シートへ全データをまるごと貼り付けるスクリプト
 *
 * 必要：Advanced Drive Service（Drive API v2）を有効化
 */

// ===== 設定 =====
const FOLDER_ID    = '1vcRWHZ3TaDgZaFNsxs4efC_2hsCcXNGG';  // Qoo10在庫ファイルを保存するフォルダID
const TARGET_SHEET = 'Qoo10在庫_raw';                     // データ貼り付け先シート名

/**
 * フォルダ内の最新CSV/XLSXをGoogleシートへ一時変換し、データを全取得して貼り付け
 */
function importLatestQoo10Inventory() {
  const ss    = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(TARGET_SHEET);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`シート「${TARGET_SHEET}」が見つかりません`);
    return;
  }

  // 1) 最新ファイルを取得
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files  = folder.getFiles();
  if (!files.hasNext()) {
    ss.toast('フォルダにファイルがありません');
    return;
  }
  let latest = files.next();
  while (files.hasNext()) {
    const f = files.next();
    if (f.getLastUpdated() > latest.getLastUpdated()) {
      latest = f;
    }
  }

  // 2) Drive API v2 でスプレッドシートへ変換コピー
  //    → convert:true はデフォルト引数なので省略可
  const tmpFile = Drive.Files.copy(
    { title: 'tmpQoo10Conv', mimeType: MimeType.GOOGLE_SHEETS },
    latest.getId()
  );

  try {
    // 3) 変換後シートから全データを取得
    const tmpSS   = SpreadsheetApp.openById(tmpFile.id);
    const srcRows = tmpSS.getSheets()[0].getDataRange().getValues();

    // 4) 一時ファイルはゴミ箱へ
    DriveApp.getFileById(tmpFile.id).setTrashed(true);

    // 5) 指定シートへまるごと貼り付け
    sheet.clearContents();
    if (srcRows.length) {
      sheet
        .getRange(1, 1, srcRows.length, srcRows[0].length)
        .setValues(srcRows);
    }
    ss.toast(`${TARGET_SHEET} を更新しました`);
  } catch (e) {
    // 何か問題が起きたらトースト＆ログ
    ss.toast('データ取得中にエラーが発生しました');
    console.error(e);
  }
}


/**
 * ドライバートリガー：古いトリガーを削除して、
 * importLatestQoo10Inventory を30分ごとに実行するタイムドリガーを作成
 */
function driverTrigger() {
  // 既存の同名トリガーを削除
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'importLatestQoo10Inventory')
    .forEach(t => ScriptApp.deleteTrigger(t));

  // 30分ごとのトリガーを新規作成
  ScriptApp.newTrigger('importLatestQoo10Inventory')
    .timeBased()
    .everyMinutes(30)
    .create();

  SpreadsheetApp.getActive().toast('30分ごとトリガーを設定しました');
}
