function 在庫検索_診断モード() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shSearch = ss.getSheetByName('在庫検索と在庫取込');
  const shY = ss.getSheetByName('Yahoo在庫ビュー');

  if (!shSearch || !shY) {
    Browser.msgBox("シートが見つかりません。シート名を確認してください。");
    return;
  }

  // --- 診断1: データ行数の確認 ---
  const lastRowY = shY.getLastRow();
  const msg1 = `【診断1】Yahoo在庫ビューの最終行は「${lastRowY}行目」です。`;
  
  // データが少なすぎる場合
  if (lastRowY < 3) {
    Browser.msgBox(msg1 + "\\n\\nデータが空、もしくはヘッダーしかありません。\\n先にYahoo在庫取込を実行してください。");
    return;
  }

  // --- 診断2: 列のデータ確認 ---
  // 3行目のデータを試し読みして表示します
  const sampleRow = 3; 
  const sampleVals = shY.getRange(sampleRow, 1, 1, 10).getValues()[0]; // A~J列を取得

  const msg2 = `【診断2】Yahoo在庫ビュー(${sampleRow}行目)のデータ確認:\n` +
               `A列(商品コード): [${sampleVals[0]}]\n` +
               `C列(Sub): [${sampleVals[2]}]\n` +
               `F列(Full): [${sampleVals[5]}]\n` +
               `G列: [${sampleVals[6]}]\n` +
               `H列: [${sampleVals[7]}]\n` +
               `I列: [${sampleVals[8]}]\n` +
               `J列: [${sampleVals[9]}]`;

  const userCheck = Browser.msgBox(msg1 + "\\n\\n" + msg2 + "\\n\\nこの列の並びで合っていますか？\\nAmazonやQoo10のデータが入っている列を確認してください。", Browser.Buttons.YES_NO);

  if (userCheck == "no") {
    ss.toast("処理を中断しました。列の設定を見直してください。", "中断");
    return;
  }

  // --- ここから本処理（強制実行） ---
  // データ開始行を判定（データがある行にあわせる）
  const DATA_START_Y = (lastRowY >= 3) ? 3 : 2; 

  // ★重要：実際のシートに合わせて列番号を指定してください
  // 例：Amazon商品コードが G列(7番目) なら 7
  const AMAZON_COL_SKU = 7; 
  const AMAZON_COL_QTY = 8;
  const QOO10_COL_SKU  = 9;
  const QOO10_COL_QTY  = 10;

  ss.toast("処理を開始します...", "在庫検索");

  // 以下、検索ロジック
  const SEARCH_START_ROW = 3;
  const lastKeyRow = shSearch.getLastRow();
  
  // 検索キー取得
  const keyRawVals = shSearch.getRange(SEARCH_START_ROW, 1, lastKeyRow - SEARCH_START_ROW + 1, 1).getValues();
  const keys = [];
  const seen = new Set();
  for (let i = 0; i < keyRawVals.length; i++) {
    const k = String(keyRawVals[i][0] || '').trim();
    if (k && !seen.has(k)) {
      seen.add(k);
      keys.push(k);
    }
  }

  // Yahooデータ読み込み
  const yRowsCount = lastRowY - DATA_START_Y + 1;
  // 読み込み範囲を広めに取る（最大列まで）
  const maxCol = Math.max(AMAZON_COL_QTY, QOO10_COL_QTY, 10);
  const yVals = shY.getRange(DATA_START_Y, 1, yRowsCount, maxCol).getValues();

  // マップ作成
  const yListByParent = new Map();
  const yListByVariant = new Map();
  const yListByFull = new Map();
  const aByMallKey = new Map();
  const qFromYahoo = new Map();
  const yAggByMallKey = new Map();

  // ヘルパー関数
  const splitCode = (raw) => {
    let s = String(raw||'').trim();
    if(!s) return {full:'', parent:'', variant:'', suffix:''};
    let parent=s, variant=s, suffix='';
    const m = s.match(/^(.*\d)([abAB])(?:-\d+)?$/);
    if(m){ parent=m[1]; suffix=m[2].toLowerCase(); variant=parent+suffix; }
    else if(/[abAB]$/.test(s) && !/-[abAB]$/.test(s)){ suffix=s.slice(-1).toLowerCase(); parent=s.slice(0,-1); variant=parent+suffix; }
    return {full:s, parent, variant, suffix};
  };

  const getParent = (c) => splitCode(c).parent; // 簡易版

  for(let i=0; i<yVals.length; i++){
    const row = yVals[i];
    // A列, C列, D列(即納), E列(お取寄), F列(Full) を想定
    const fullKey = String(row[5]||row[0]||'').trim();
    if(!fullKey) continue;

    const info = splitCode(fullKey);
    const sokuno = Number(row[3]||0);
    const otori = Number(row[4]||0);
    const mallKey = getParent(fullKey);

    const rowObj = { full: info.full, parent: info.parent, variant: info.variant, sokuno, otori, mallKey };
    
    if(!yListByFull.has(info.full)) yListByFull.set(info.full, []);
    yListByFull.get(info.full).push(rowObj);

    if(!yListByParent.has(info.parent)) yListByParent.set(info.parent, []);
    yListByParent.get(info.parent).push(rowObj);

    if(!yListByVariant.has(info.variant)) yListByVariant.set(info.variant, []);
    yListByVariant.get(info.variant).push(rowObj);

    // 集計
    let agg = yAggByMallKey.get(mallKey) || {s:0, o:0};
    agg.s += sokuno; agg.o += otori;
    yAggByMallKey.set(mallKey, agg);

    // Amazon
    const aSku = String(row[AMAZON_COL_SKU-1]||'').trim();
    if(aSku){
      const aKey = getParent(aSku);
      let aRec = aByMallKey.get(aKey) || {code:aSku, qty:0};
      aRec.qty += Number(row[AMAZON_COL_QTY-1]||0);
      aByMallKey.set(aKey, aRec);
    }
    // Qoo10
    const qSku = String(row[QOO10_COL_SKU-1]||'').trim();
    if(qSku){
      const qRec = qFromYahoo.get(info.full) || {sku:qSku, qty:0};
      qRec.qty += Number(row[QOO10_COL_QTY-1]||0);
      qFromYahoo.set(info.full, qRec);
    }
  }

  // 書き出しデータ作成
  const result = [];
  for(let k of keys){
    const kInfo = splitCode(k);
    let targets = [];
    if(yListByVariant.has(kInfo.variant)) targets = yListByVariant.get(kInfo.variant);
    else if(yListByParent.has(kInfo.parent)) targets = yListByParent.get(kInfo.parent);
    
    // 見つからない場合
    if(targets.length === 0){
       const row = new Array(12).fill('');
       row[0] = k; row[3] = '在庫なし'; row[11] = '未登録';
       result.push(row);
       continue;
    }

    // 見つかった場合
    let isFirst = true;
    for(let yDat of targets){
      const row = new Array(12).fill('');
      row[0] = isFirst ? k : '';
      row[1] = yDat.full;
      row[2] = yDat.sokuno + yDat.otori;
      row[3] = (row[2]>0) ? 'あり' : 'なし'; // 簡易表示
      row[11] = '登録済';
      
      const pAgg = yAggByMallKey.get(yDat.mallKey);
      const aDat = aByMallKey.get(yDat.mallKey);
      if(aDat && isFirst){ // Amazonは親単位で1回
         row[4] = aDat.code; row[5] = aDat.qty; row[7] = '登録済';
      }
      const qDat = qFromYahoo.get(yDat.full);
      if(qDat){
         row[8] = qDat.sku; row[9] = qDat.qty; row[10] = '登録済';
      }

      result.push(row);
      isFirst = false;
    }
  }

  // 書き込み
  if(result.length > 0){
    const clearRows = Math.max(0, shSearch.getMaxRows() - SEARCH_START_ROW + 1);
    if(clearRows>0) shSearch.getRange(SEARCH_START_ROW, 1, clearRows, 12).clearContent();
    shSearch.getRange(SEARCH_START_ROW, 1, result.length, 12).setValues(result);
    ss.toast("更新完了しました", "完了");
  } else {
    ss.toast("検索結果が0件でした", "確認");
  }
}