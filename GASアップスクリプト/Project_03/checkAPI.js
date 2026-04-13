//――――― デバッグ用の詳細ログ関数 ―――――
function runAllProcessesDebug() {
  Logger.log('runAllProcesses: start');
  
  // 認証状態確認
  try {
    const token = getValidAccessToken();
    Logger.log('認証OK: トークン取得成功');
  } catch (e) {
    Logger.log('認証エラー: ' + e.message);
    Logger.log('スタックトレース: ' + e.stack);
    return;
  }
  
  // 商品データ更新
  try { 
    Logger.log('商品データ更新開始...');
    updateProductSheet(); 
    Logger.log('商品データ更新完了');
  } catch (e) { 
    Logger.log('商品更新エラー: ' + e.message);
    Logger.log('スタックトレース: ' + e.stack);
  }
  
  // 在庫データ更新
  try { 
    Logger.log('在庫データ更新開始...');
    updateStockSheet(); 
    Logger.log('在庫データ更新完了');
  } catch (e) { 
    Logger.log('在庫更新エラー: ' + e.message); 
    Logger.log('スタックトレース: ' + e.stack);
  }
  
  // 商品リスト更新
  try { 
    Logger.log('商品リスト更新開始...');
    updateProductList(); 
    Logger.log('商品リスト更新完了');
  } catch (e) { 
    Logger.log('一覧更新エラー: ' + e.message);
    Logger.log('スタックトレース: ' + e.stack);
  }
  
  Logger.log('runAllProcesses: end');
}

//――――― CSV取得のテスト関数 ―――――
function testFetchCsv() {
  Logger.log('=== CSV取得テスト開始 ===');
  
  try {
    Logger.log('商品CSVテスト開始...');
    const productCsv = fetchCsv(PRODUCT_CSV_TYPE);
    Logger.log('商品CSV取得成功 - データ長: ' + productCsv.length);
    Logger.log('商品CSV先頭100文字: ' + productCsv.substring(0, 100));
    
    Logger.log('在庫CSVテスト開始...');
    const stockCsv = fetchCsv(STOCK_CSV_TYPE);
    Logger.log('在庫CSV取得成功 - データ長: ' + stockCsv.length);
    Logger.log('在庫CSV先頭100文字: ' + stockCsv.substring(0, 100));
    
  } catch (e) {
    Logger.log('CSV取得エラー: ' + e.message);
    Logger.log('スタックトレース: ' + e.stack);
  }
  
  Logger.log('=== CSV取得テスト終了 ===');
}

//――――― シート存在確認関数 ―――――
function checkSheets() {
  Logger.log('=== シート存在確認 ===');
  
  const ss = SpreadsheetApp.getActive();
  const productSheet = ss.getSheetByName('CSV商品データ');
  const stockSheet = ss.getSheetByName('CSV在庫データ');
  
  Logger.log('商品シート存在: ' + !!productSheet);
  Logger.log('在庫シート存在: ' + !!stockSheet);
  
  if (productSheet) {
    Logger.log('商品シート行数: ' + productSheet.getLastRow());
    Logger.log('商品シート列数: ' + productSheet.getLastColumn());
  }
  
  if (stockSheet) {
    Logger.log('在庫シート行数: ' + stockSheet.getLastRow());
    Logger.log('在庫シート列数: ' + stockSheet.getLastColumn());
  }
}