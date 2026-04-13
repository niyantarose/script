/**
 * テンプレートシートをコピー（UI不要版）
 */
function テンプレートシートをコピー() {
  try {
    // ★ コピー元のスプレッドシート（Googleスプレッドシート版）
    const sourceSpreadsheetId = '1mKBJFFyUrp0BVrXZSXYLv5LiMPBqUdisTQvuu2o4biU';
    const sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
    const sourceSheet = sourceSpreadsheet.getSheetByName('テンプレート');  // ←シート名これでOKならそのまま

    if (!sourceSheet) {
      Logger.log('❌ コピー元に「テンプレート」シートが見つかりません');
      return;
    }

    // コピー先のスプレッドシート（Amazon商品登録）
    const targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // 既存の「テンプレート」シートを削除（存在する場合）
    const existingSheet = targetSpreadsheet.getSheetByName('テンプレート');
    if (existingSheet) {
      Logger.log('既存のテンプレートシートを削除中...');
      targetSpreadsheet.deleteSheet(existingSheet);
    }

    // シートをコピー
    Logger.log('シートをコピー中...');
    const copiedSheet = sourceSheet.copyTo(targetSpreadsheet);
    copiedSheet.setName('テンプレート');

    // 最初のシートとして移動
    targetSpreadsheet.setActiveSheet(copiedSheet);
    targetSpreadsheet.moveActiveSheet(1);

    Logger.log('✅ テンプレートシートのコピーが完了しました');

  } catch (error) {
    Logger.log('❌ エラー: ' + error.toString());
  }
}
