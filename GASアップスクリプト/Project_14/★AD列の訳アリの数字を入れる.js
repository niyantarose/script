function 訳アリ中古チェック() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  // データ範囲を取得(2行目から最終行まで)
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  // E列とAD列のデータを取得
  const eColumn = sheet.getRange(2, 5, lastRow - 1, 1).getValues(); // E列(商品名)
  const adColumn = sheet.getRange(2, 30, lastRow - 1, 1).getValues(); // AD列
  
  // チェックするキーワード
  const keywords = ['訳アリ', '訳あり', '訳有', '中古', '難あり', '難アリ', 'わけあり'];
  
  // 更新用の配列
  const updates = [];
  
  for (let i = 0; i < eColumn.length; i++) {
    const productName = eColumn[i][0] ? String(eColumn[i][0]) : '';
    let shouldUpdate = false;
    
    // キーワードチェック
    for (let keyword of keywords) {
      if (productName.includes(keyword)) {
        shouldUpdate = true;
        break;
      }
    }
    
    // 該当する場合は5を設定
    updates.push([shouldUpdate ? 5 : adColumn[i][0]]);
  }
  
  // AD列に一括更新
  sheet.getRange(2, 30, updates.length, 1).setValues(updates);
  
  SpreadsheetApp.getActiveSpreadsheet().toast('訳アリ・中古チェックが完了しました', '処理完了', 3);
}