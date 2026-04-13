function Yahoo画像URLをQoo10に転記() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const yahooSheet = ss.getSheetByName('yahoo変換元ファイル');
  const qoo10Sheet = ss.getSheetByName('Qoo10');
  
  if (!yahooSheet || !qoo10Sheet) {
    SpreadsheetApp.getUi().alert('シートが見つかりません');
    return;
  }
  
  // Yahoo変換元ファイルのデータ取得
  const yahooLastRow = yahooSheet.getLastRow();
  if (yahooLastRow < 2) {
    SpreadsheetApp.getUi().alert('Yahoo変換元ファイルにデータがありません');
    return;
  }
  
  // C列(code)を取得
  const yahooCodeColumn = yahooSheet.getRange(1, 3, yahooLastRow, 1).getValues();
  
  // CL列とCM列のヘッダー確認と取得
  const yahooHeaders = yahooSheet.getRange(1, 1, 1, yahooSheet.getLastColumn()).getValues()[0];
  
  let imageUrlColumn = -1;
  for (let i = 0; i < yahooHeaders.length; i++) {
    if (yahooHeaders[i] === 'item-image-urls') {
      imageUrlColumn = i + 1; // 列番号(1始まり)
      break;
    }
  }
  
  if (imageUrlColumn === -1) {
    SpreadsheetApp.getUi().alert('item-image-urls列が見つかりません');
    return;
  }
  
  const yahooImageUrls = yahooSheet.getRange(2, imageUrlColumn, yahooLastRow - 1, 1).getValues();
  const yahooCodes = yahooCodeColumn.slice(1); // ヘッダー除く
  
  // Yahoo側のデータをマップ化(code → 変換後URL)
  const yahooMap = new Map();
  for (let i = 0; i < yahooCodes.length; i++) {
    const code = yahooCodes[i][0];
    const imageUrl = yahooImageUrls[i][0];
    
    if (code && imageUrl) {
      // セミコロンを$$に変換（split/joinで確実に変換）
      const convertedUrl = String(imageUrl).split(';').join('$$');
      yahooMap.set(String(code), convertedUrl);
    }
  }
  
  // Qoo10シートのデータ取得
  const qoo10LastRow = qoo10Sheet.getLastRow();
  if (qoo10LastRow < 2) {
    SpreadsheetApp.getUi().alert('Qoo10シートにデータがありません');
    return;
  }
  
  // B列(seller_unique_item_id)とR列(image_other_url)を取得
  const qoo10SellerIds = qoo10Sheet.getRange(2, 2, qoo10LastRow - 1, 1).getValues();
  const qoo10ImageOtherUrls = qoo10Sheet.getRange(2, 18, qoo10LastRow - 1, 1).getValues();
  
  // 更新用配列（R列とQ列）
  const updatesR = [];
  const updatesQ = [];
  let updateCount = 0;
  
  for (let i = 0; i < qoo10SellerIds.length; i++) {
    const sellerId = String(qoo10SellerIds[i][0]);
    
    if (yahooMap.has(sellerId)) {
      const fullUrl = yahooMap.get(sellerId);
      updatesR.push([fullUrl]);
      
      // R列の一番前のURLをQ列に転記
      // $$で分割して最初の要素を取得
      const firstUrl = fullUrl.split('$$')[0];
      updatesQ.push([firstUrl]);
      
      updateCount++;
    } else {
      updatesR.push([qoo10ImageOtherUrls[i][0]]); // 元の値を保持
      updatesQ.push(['']); // マッチしない場合はQ列を空にしない（元の値を保持したい場合は別途取得が必要）
    }
  }
  
  // マッチしない行のQ列は元の値を保持する場合
  const qoo10MainImages = qoo10Sheet.getRange(2, 17, qoo10LastRow - 1, 1).getValues();
  for (let i = 0; i < updatesQ.length; i++) {
    if (updatesQ[i][0] === '') {
      updatesQ[i][0] = qoo10MainImages[i][0]; // 元の値を保持
    }
  }
  
  // R列（18列目）に一括更新
  qoo10Sheet.getRange(2, 18, updatesR.length, 1).setValues(updatesR);
  
  // Q列（17列目）に一括更新
  qoo10Sheet.getRange(2, 17, updatesQ.length, 1).setValues(updatesQ);
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `${updateCount}件の画像URLを更新しました（R列・Q列）`, 
    '処理完了', 
    5
  );
}
