function Yahoo在庫F列をAmazon変更シートC列に反映() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const srcSheet = ss.getSheetByName('Yahoo在庫ビュー');
  const dstSheet = ss.getSheetByName('Amazon価格在庫変更');

  if (!srcSheet) throw new Error('「Yahoo在庫ビュー」が見つかりません');
  if (!dstSheet) throw new Error('「Amazon価格在庫変更」が見つかりません');

  // ==========================================
  // 設定: Amazonシートの列位置 (画像に合わせて設定)
  // ==========================================
  const DST_START_ROW = 4; // データは4行目から
  const DST_SKU_COL   = 1; // A列 (SKU)
  const DST_QTY_COL   = 3; // C列 (Quantity)

  // ==========================================
  // 1. AmazonシートからSKU一覧を取得 (A列)
  // ==========================================
  const dstLastRow = dstSheet.getLastRow();
  if (dstLastRow < DST_START_ROW) {
    Browser.msgBox('Amazon価格在庫変更シートに更新対象のSKU(A列)がありません。');
    return;
  }

  // A列のSKUを全て取得
  const targetSkus = dstSheet.getRange(DST_START_ROW, DST_SKU_COL, dstLastRow - DST_START_ROW + 1, 1).getValues();

  // ==========================================
  // 2. Yahoo在庫ビューを読み込み (F列が在庫数)
  // ==========================================
  const srcLastRow = srcSheet.getLastRow();
  // 必要な列: E列(統合キー/SKU) と F列(在庫数)
  // E列は5番目, F列は6番目
  const srcValues = srcSheet.getRange(3, 1, srcLastRow - 2, 6).getValues();
  
  // マップ作成: Key(E列) -> Quantity(F列)
  const stockMap = new Map();

  for (let i = 0; i < srcValues.length; i++) {
    const row = srcValues[i];
    const key = String(row[4] || '').trim(); // E列: 統合キー(SKU)
    const qty = Number(row[5]);              // F列: 在庫数 ★ここをF列に変更
    
    // 在庫数を安全な数値に変換
    const safeQty = isNaN(qty) ? 0 : qty;

    if (key) {
      stockMap.set(key, safeQty);
    }
  }

  // ==========================================
  // 3. AmazonのSKUと照合して更新用データを作成
  // ==========================================
  const updateValues = []; // 書き込み用（在庫数のみ）

  for (let i = 0; i < targetSkus.length; i++) {
    const sku = String(targetSkus[i][0] || '').trim();
    let newQty = 0; // 見つからない場合は0

    if (sku) {
      // 検索ロジック:
      // 1. そのまま検索 (完全一致)
      if (stockMap.has(sku)) {
        newQty = stockMap.get(sku);
      } 
      // 2. 「SKU + a」で検索 (AmazonにサフィックスがなくてもYahooのa付きを拾う)
      else if (stockMap.has(sku + 'a')) {
        newQty = stockMap.get(sku + 'a');
      } 
      // 3. 「SKU + A」で検索
      else if (stockMap.has(sku + 'A')) {
        newQty = stockMap.get(sku + 'A');
      }
    }

    // 更新用配列に追加
    updateValues.push([newQty]);
  }

  // ==========================================
  // 4. C列（在庫数）だけを一括更新
  // ==========================================
  if (updateValues.length > 0) {
    // C列の対象範囲に書き込み (1000万セル制限回避のため列を絞る)
    dstSheet.getRange(DST_START_ROW, DST_QTY_COL, updateValues.length, 1).setValues(updateValues);
    
    const msg = `完了しました。\nAmazonシートのA列にある ${updateValues.length} 件に対し、Yahoo在庫ビュー(F列)の在庫を反映しました。`;
    ss.toast(msg, '更新完了', 5);
  }
}