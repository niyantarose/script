/******************************************************
 * 在庫検索＋在庫取込まとめ
 * （F列ベースSKU対応・表示は行在庫/判定は親在庫・「-数字」落とし保険検索付き）
 * 2026/01/14 修正：
 * - 検索ヒットしない場合、末尾の「-数字」を削除して再検索する保険ロジックを追加
 * - これにより DCWEB-OL-02-1 がヒットしなくても DCWEB-OL-02 で親を拾いに行く
 ******************************************************/
function 在庫検索_更新() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const shSearch = ss.getSheetByName('在庫検索と在庫取込');
  const shY      = ss.getSheetByName('Yahoo在庫ビュー');

  if (!shSearch || !shY) {
    ss.toast('シートが見つかりません', 'エラー', 5);
    return;
  }

  // ▼▼▼ 設定エリア（Yahoo在庫ビューの列番号） ▼▼▼
  const Y_MAP = {
    FULL:   6,  // F列: Yahoo統合キー
    SOKUNO: 4,  // D列: 即納在庫
    OTORI:  5,  // E列: お取り寄せ在庫

    // Amazon
    AMZ_SKU: 9,   // I列: 商品コード
    AMZ_QTY: 10,  // J列: 在庫数
    AMZ_ST:  13,  // M列: 登録状況

    // Qoo10
    QOO_SKU: 15,  // O列: 商品コード
    QOO_QTY: 16,  // P列: 在庫数
    QOO_ST:  18   // R列: 登録状況
  };
  // ▲▲▲ 設定エリア終了 ▲▲▲

  // --- 0. 検索キー確認 ---
  const SEARCH_START_ROW = 3;
  const lastKeyRow = getLastDataRowByCol_AMZ_(shSearch, 1, SEARCH_START_ROW);
  if (lastKeyRow < SEARCH_START_ROW) {
    const maxRow = shSearch.getMaxRows();
    const clearRows = Math.max(0, maxRow - SEARCH_START_ROW + 1);
    if (clearRows > 0) shSearch.getRange(SEARCH_START_ROW, 1, clearRows, 14).clearContent();
    ss.toast('A列に検索コードがありません', '確認', 4);
    return;
  }

  // --- 1. Yahoo在庫ビュー読み込み ---
  const yLast = shY.getLastRow();
  let DATA_START_Y = 3;
  if (yLast < 3 && yLast >= 2) DATA_START_Y = 2;

  if (yLast < DATA_START_Y) {
    ss.toast('「Yahoo在庫ビュー」にデータがありません。', 'データ不足', 8);
    return;
  }

  const yRowsCount = yLast - DATA_START_Y + 1;
  const lastColNum = Math.max(Y_MAP.QOO_ST, 18);
  const yVals = shY.getRange(DATA_START_Y, 1, yRowsCount, lastColNum).getValues();

  // マップ作成
  const yListByParent  = new Map();
  const yListByFull    = new Map();
  const yListByBase    = new Map();
  const yAggByMallKey  = new Map();
  const aByMallKey     = new Map();

  for (let i = 0; i < yVals.length; i++) {
    const row = yVals[i];

    const fullKeyRaw = String(row[Y_MAP.FULL - 1] || row[0] || '').trim();
    if (!fullKeyRaw) continue;

    // ベースSKU生成（a/b除去）
    const baseKeyRaw = getBaseSkuFromYahooFull_(fullKeyRaw);
    
    // キーの大文字化
    const fullKeyU   = fullKeyRaw.toUpperCase();
    const baseKeyU   = baseKeyRaw.toUpperCase();
    const codeColRaw = String(row[0] || '').trim();
    const parentKeyU = codeColRaw.toUpperCase();

    // マップの親キー(mallKey)は ベースSKU を最優先
    const mallKey = baseKeyU || fullKeyU; 

    const sokuno = Number(row[Y_MAP.SOKUNO - 1] || 0) || 0;
    const otori  = Number(row[Y_MAP.OTORI  - 1] || 0) || 0;
    const yTotal = sokuno + otori;

    const rowObj = {
      full:   fullKeyRaw,
      base:   baseKeyRaw,
      sokuno,
      otori,
      mallKey,

      // Amazon
      amzSku: String(row[Y_MAP.AMZ_SKU - 1] || '').trim(),
      amzQty: row[Y_MAP.AMZ_QTY - 1],
      amzSt:  String(row[Y_MAP.AMZ_ST - 1]  || '').trim(),

      // Qoo10
      qooSku: String(row[Y_MAP.QOO_SKU - 1] || '').trim(),
      qooQty: row[Y_MAP.QOO_QTY - 1],
      qooSt:  String(row[Y_MAP.QOO_ST - 1]  || '').trim()
    };

    // リスト登録
    if (!yListByFull.has(fullKeyU)) yListByFull.set(fullKeyU, []);
    yListByFull.get(fullKeyU).push(rowObj);

    if (baseKeyU) {
      if (!yListByBase.has(baseKeyU)) yListByBase.set(baseKeyU, []);
      yListByBase.get(baseKeyU).push(rowObj);
    }

    if (parentKeyU) {
      if (!yListByParent.has(parentKeyU)) yListByParent.set(parentKeyU, []);
      yListByParent.get(parentKeyU).push(rowObj);
    }

    // 在庫合計（mallKey単位）
    let agg = yAggByMallKey.get(mallKey);
    if (!agg) agg = { sokuno: 0, otori: 0, total: 0 };
    agg.sokuno += sokuno;
    agg.otori  += otori;
    agg.total  += yTotal;
    yAggByMallKey.set(mallKey, agg);

    // Amazon在庫合計
    if (rowObj.amzSku) {
      let aRec = aByMallKey.get(mallKey);
      if (!aRec) aRec = { qty: 0 };
      const aq = Number(rowObj.amzQty);
      if (!isNaN(aq)) aRec.qty += aq;
      aByMallKey.set(mallKey, aRec);
    }
  }

  // --- 2. 検索実行 ＆ 結果作成 ---
  const keyRawVals = shSearch.getRange(SEARCH_START_ROW, 1, lastKeyRow - SEARCH_START_ROW + 1, 1).getValues();
  const keys = keyRawVals.map(r => String(r[0] || '').trim()).filter(k => k);
  const result = [];

  for (let idx = 0; idx < keys.length; idx++) {
    const keyRaw  = keys[idx];
    
    const keyFullU   = keyRaw.toUpperCase();
    const keyBaseU   = getBaseSkuFromYahooFull_(keyRaw).toUpperCase();
    const keyInfo    = _splitCode3Patterns_(keyRaw);
    const keyParentU = (keyInfo.parent || '').toUpperCase();

    let targets = [];

    // 1) まずは完全一致（F列 = 検索キーそのもの）
    if (yListByFull.has(keyFullU)) {
      targets = yListByFull.get(keyFullU);

    // 2) ベースSKU一致（F列から a/b を取ったもの）
    } else if (keyBaseU && yListByBase.has(keyBaseU)) {
      targets = yListByBase.get(keyBaseU);

    // 3) A列の親コード一致
    } else if (keyParentU && yListByParent.has(keyParentU)) {
      targets = yListByParent.get(keyParentU);
    }

    // ★★ 4) それでもヒットしないときの「保険」
    // 末尾の「-数字」を一段落として親コードとして再検索
    if (targets.length === 0) {
      const keyNoIdxU = stripLastIndexForSearch_(keyFullU);  // 例: DCWEB-OL-02-1 → DCWEB-OL-02
      if (keyNoIdxU && keyNoIdxU !== keyFullU) {
        if (yListByBase.has(keyNoIdxU)) {
          targets = yListByBase.get(keyNoIdxU);
        } else if (yListByParent.has(keyNoIdxU)) {
          targets = yListByParent.get(keyNoIdxU);
        }
      }
    }

    // ヒットなし
    if (targets.length === 0) {
      const row = new Array(14).fill('');
      row[0] = keyRaw;
      row[3] = '在庫なし';
      row[7] = '未登録';    
      row[11] = '未登録';   
      row[12] = 0;
      row[13] = 0;
      result.push(row);
      continue;
    }

    const mallKey    = targets[0].mallKey;
    const parentAgg  = yAggByMallKey.get(mallKey) || { sokuno: 0, otori: 0, total: 0 };
    const aRecParent = aByMallKey.get(mallKey);

    // 親合計在庫（判定用）
    const yStockParent = parentAgg.total;
    const ySokunoParent = parentAgg.sokuno;
    const yOtoriParent  = parentAgg.otori;

    // Amazon発送区分（親単位）
    let shipModeForMall = '在庫なし';
    if (ySokunoParent > 0)      shipModeForMall = '即納';
    else if (yOtoriParent > 0)  shipModeForMall = 'お取り寄せ';

    let isFirst = true;
    for (const yDat of targets) {
      const row = new Array(14).fill('');

      row[0] = isFirst ? keyRaw : '';
      row[1] = yDat.full;

      // C列: 行ごとの在庫数を表示
      const yRowStock = (Number(yDat.sokuno) || 0) + (Number(yDat.otori) || 0);
      row[2] = yRowStock;

      // D列: 行ごとの在庫区分を表示
      let stockMode = '在庫なし';
      if (yDat.sokuno > 0 && yDat.otori > 0) stockMode = '即納+お取り寄せ';
      else if (yDat.sokuno > 0)             stockMode = '即納';
      else if (yDat.otori > 0)              stockMode = 'お取り寄せ';
      row[3] = stockMode;

      // ======== Amazon(E〜H) ========
      const amzSku = yDat.amzSku;
      const amzQty = yDat.amzQty;
      const amzSt  = (yDat.amzSt || '').trim();

      const hasAmzSku      = !!amzSku;
      const amzStatusHas未 = amzSt.indexOf('未') > -1;
      const amzRegistered  = hasAmzSku && !amzStatusHas未;

      // H列判定：【親合計在庫】を使用
      if (yStockParent === 0) {
        row[7] = '在庫なし';
      } else if (amzRegistered) {
        row[7] = '登録済';
      } else {
        row[7] = '未登録';
      }

      // E/F/G
      if (row[7] === '登録済') {
        row[4] = yDat.full;
        const aq = Number(amzQty);
        row[5] = isNaN(aq) ? 0 : aq;
        row[6] = shipModeForMall;
      }

      // ======== Qoo10(I〜L) ========
      const qooSku = yDat.qooSku;
      const qooQty = yDat.qooQty;
      const qooSt  = (yDat.qooSt || '').trim();

      const hasQooSku      = !!qooSku;
      const qooStatusHas未 = qooSt.indexOf('未') > -1;
      const qooRegistered  = hasQooSku && !qooStatusHas未;

      // L列判定：【親合計在庫】を使用
      if (yStockParent === 0) {
        row[11] = '在庫なし';
      } else if (qooRegistered) {
        row[11] = '登録済';
      } else {
        row[11] = '未登録';
      }

      if (row[11] === '登録済') {
        row[8] = qooSku;
        const qq = Number(qooQty);
        row[9] = isNaN(qq) ? 0 : qq;

        // 即納 / お取り寄せ (親合計ベース)
        let qModeBase = '在庫なし';
        if (ySokunoParent > 0)      qModeBase = '即納';
        else if (yOtoriParent > 0)  qModeBase = 'お取り寄せ';
        row[10] = qModeBase;
      }

      // 差分計算（親合計との比較）
      if (isFirst) {
        const aParentQty = aRecParent ? aRecParent.qty : 0;
        row[12] = aParentQty - yStockParent;
      } else {
        row[12] = '';
      }

      const qVal = Number(yDat.qooQty) || 0;
      row[13] = qVal - yStockParent;

      result.push(row);
      isFirst = false;
    }
  }

  // --- 3. 書き込み ---
  const maxRow = shSearch.getMaxRows();
  const clearRows = Math.max(0, maxRow - SEARCH_START_ROW + 1);
  if (clearRows > 0) shSearch.getRange(SEARCH_START_ROW, 1, clearRows, 14).clearContent();

  if (result.length > 0) {
    const neededRows = SEARCH_START_ROW + result.length - 1;
    if (maxRow < neededRows) shSearch.insertRowsAfter(maxRow, neededRows - maxRow);
    shSearch.getRange(SEARCH_START_ROW, 1, result.length, 14).setValues(result);
    ss.toast(`完了：${result.length}行のデータを更新しました`, '完了', 5);
  } else {
    ss.toast('検索結果は0件でした', '完了', 5);
  }
}

// === ヘルパー関数 ===

// ★検索用: 末尾の「-数字」を1段だけ落としたキーを作る
// 例）DCWEB-OL-02-1 → DCWEB-OL-02
function stripLastIndexForSearch_(s) {
  s = String(s || '').trim();
  if (!s) return '';
  // 大文字小文字は事前に揃えてから来る前提
  const m = s.match(/^(.+)-\d+$/);
  return m ? m[1] : s;
}

// （以下、既存ヘルパーも維持）
function getLastDataRowByCol_AMZ_(sheet, col, startRow) {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return startRow - 1;
  const values = sheet.getRange(startRow, col, lastRow - startRow + 1, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0] || '').trim() !== '') return startRow + i;
  }
  return startRow - 1;
}

function getBaseSkuFromYahooFull_(full) {
  var s = String(full || '').trim();
  if (!s) return '';
  // 貪欲マッチで途中区切り防止
  var m = s.match(/^(.+)([abAB])(-.*)?$/);
  if (m) {
    var base = m[1];
    var rest = m[3] || '';
    return base + rest;
  }
  return s;
}

function _splitCode3Patterns_(raw) {
  let s = String(raw || '').trim();
  if (!s) return { full: '', parent: '', variant: '', suffix: '' };
  const full = s;
  let parent = full;
  let suffix = '';

  const m = s.match(/^(.+)([abAB])(-.*)?$/);
  if (m) {
    parent = (m[1] || '') + (m[3] || '');
    suffix = m[2].toLowerCase();
  }
  const variant = full;
  return { full, parent, variant, suffix };
}

function _analyzeFullCode_(k) { return _splitCode3Patterns_(k); }
function _parseSearchKey_(k) {
  const o = _splitCode3Patterns_(k);
  let mode = 'parent';
  if (o.suffix) mode = 'suffix';
  return { mode, ...o };
}