function Yahoo在庫数をQoo10変更シートに反映() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const srcSheet = ss.getSheetByName('Yahoo在庫ビュー');
  const dstSheet = ss.getSheetByName('Qoo10価格在庫変更');

  if (!srcSheet) throw new Error('「Yahoo在庫ビュー」が見つかりません');
  if (!dstSheet) throw new Error('「Qoo10価格在庫変更」が見つかりません');

  // ==========================================
  // 1. Yahoo在庫ビュー（元データ）を読み込み
  // ==========================================
  const srcLastRow = srcSheet.getLastRow();
  if (srcLastRow < 3) {
    Browser.msgBox('Yahoo在庫ビューにデータがありません');
    return;
  }

  // A列:code, B:name, C:sub, D:quantity, E:統合キー
  // 必要なのは A列(code), D列(qty), E列(key)
  const srcValues = srcSheet.getRange(3, 1, srcLastRow - 2, 5).getValues();
  
  // Yahoo在庫をMap化 (Key -> Quantity)
  const yahooStockMap = new Map();

  for (let i = 0; i < srcValues.length; i++) {
    const row = srcValues[i];
    const code = String(row[0] || '').trim(); // A列
    const qty  = Number(row[3]);              // D列 (在庫数)
    const key  = String(row[4] || '').trim(); // E列 (統合キー)
    
    // 確実な数値にする
    const safeQty = isNaN(qty) ? 0 : qty;

    // 統合キーがあればそれを優先キーにする
    if (key) {
      yahooStockMap.set(key, safeQty);
    }
    // code単体もキーとして登録しておく（統合キーがない場合や親コードマッチ用）
    if (code) {
      // 統合キーですでに登録されていない場合のみ、あるいはバックアップとして登録
      if (!yahooStockMap.has(code)) {
        yahooStockMap.set(code, safeQty);
      }
    }
  }

  // ==========================================
  // 2. Qoo10価格在庫変更（更新先）を読み込み
  // ==========================================
  const DST_START_ROW = 5; // データ開始行
  const dstLastRow = dstSheet.getLastRow();

  if (dstLastRow < DST_START_ROW) {
    Browser.msgBox('Qoo10価格在庫変更に更新対象のデータがありません');
    return;
  }

  // B列:seller_unique_item_id, D列:seller_unique_option_id
  const numRows = dstLastRow - DST_START_ROW + 1;
  const dstKeyRange = dstSheet.getRange(DST_START_ROW, 1, numRows, 4); // A〜D列を取得
  const dstKeyValues = dstKeyRange.getValues();

  // 更新用在庫数リスト
  const updateQtyValues = [];

  // ==========================================
  // 3. マッチングとデータ作成
  // ==========================================
  let updateCount = 0;

  for (let i = 0; i < dstKeyValues.length; i++) {
    const itemId   = String(dstKeyValues[i][1] || '').trim(); // B列
    const optionId = String(dstKeyValues[i][3] || '').trim(); // D列

    // Qoo10側のキーを決定（オプションID優先）
    let targetKey = optionId || itemId;

    // Yahoo在庫Mapから検索
    let newQty = null;

    // パターンA: 完全一致検索
    if (yahooStockMap.has(targetKey)) {
      newQty = yahooStockMap.get(targetKey);
    } 
    // パターンB: Qoo10側が 'a' 付きだが Yahoo側にない場合 -> 'a'なしで再検索
    else if (targetKey.toLowerCase().endsWith('a')) {
       const keyWithoutA = targetKey.slice(0, -1); // 末尾のaを削除
       if (yahooStockMap.has(keyWithoutA)) {
         newQty = yahooStockMap.get(keyWithoutA);
       }
    }
    // パターンC: Qoo10側が 'a' なしだが Yahoo側に 'a' 付きがある場合 -> 'a'をつけて再検索
    else {
       const keyWithA = targetKey + 'a';
       if (yahooStockMap.has(keyWithA)) {
         newQty = yahooStockMap.get(keyWithA);
       }
    }

    // 値が見つかった場合はその数値を、見つからなければ「変更なし(null)」または「0」
    // ※ここでは「見つからなければ0にする」運用にします（安全のため）
    // もし「見つからない場合は空欄」が良い場合は `newQty !== null ? newQty : ''` にしてください
    const finalQty = (newQty !== null) ? newQty : 0;
    
    updateQtyValues.push([finalQty]);
    updateCount++;
  }

  // ==========================================
  // 4. 書き込み実行（F列）
  // ==========================================
  if (updateQtyValues.length > 0) {
    // F列 (6列目) に書き込み
    dstSheet.getRange(DST_START_ROW, 6, updateQtyValues.length, 1).setValues(updateQtyValues);
    
    const msg = `完了しました。\n${updateCount} 件の在庫数をYahoo在庫ビューに合わせて更新しました。`;
    ss.toast(msg, '更新完了', 5);
  }
}