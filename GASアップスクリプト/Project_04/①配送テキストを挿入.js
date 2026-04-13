/******************************************************
 * 10_shipping_master.gs - 配送マスタ読込 & 配送テキスト反映
 * v3.1: R列が数字でもテキストでも全処理を実行
 * 
 * ✅ v3.1の処理フロー:
 *   R列を読む（テキスト or 数字どちらでも対応）
 *   → テキストの場合: B列で照合して種別を特定
 *   → 数字の場合:     D列→B列逆引きで種別を特定
 *   → K列設定: 佐川→1000, ゆうパケ系→100
 *   → I列/Q列: 配送テキスト反映
 *   → R列: テキストなら管理番号(数字)に変換
 ******************************************************/

// ====================================================================
// ★ 安全なUI通知（UIが無くても落ちない）
// ====================================================================

function SHIP_uiSafeAlert_(msg) {
  try {
    SpreadsheetApp.getUi().alert(String(msg));
    return;
  } catch (e) {}

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(String(msg), '配送マスタ', 10);
  } catch (e2) {}

  Logger.log('[SHIP_UI_FALLBACK] ' + msg);
}


/**
 * 配送マスタ読み込み
 */
function 配送マスタを読み込み_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const groupSheet = ss.getSheetByName(SHEET_配送グループ設定);
  const textSheet  = ss.getSheetByName(SHEET_配送種別テキスト);

  if (!groupSheet || !textSheet) {
    SHIP_uiSafeAlert_('配送グループ設定 / 配送種別テキスト シートが見つかりません。');
    throw new Error('配送マスタシートが見つからない');
  }

  Logger.log('=== 配送マスタ読み込み 開始 ===');

  // ① 配送グループ設定: A(postage-set), B(配送種別), C(備考), D(管理番号)
  const groupLastRow = groupSheet.getLastRow();
  const groupValues = groupLastRow >= 2
    ? groupSheet.getRange(2, 1, groupLastRow - 1, 4).getValues()
    : [];

  // A列(postage-set) → B列(type): 従来互換
  const postageToType = {};
  groupValues.forEach(row => {
    const postageSet = String(row[0] || '').trim();
    const type       = String(row[1] || '').trim();
    if (postageSet && type) postageToType[postageSet] = type;
  });

  // B列(配送種別) → D列(管理番号)
  const typeToNum = {};
  // ✅ v3.1: D列(管理番号) → B列(配送種別) 逆引きマップ
  const numToType = {};
  groupValues.forEach(row => {
    const type = String(row[1] || '').trim();   // B列
    const num  = row[3];                         // D列
    if (type && num !== '' && num !== null && num !== undefined) {
      const n = Number(num);
      if (!(type in typeToNum)) typeToNum[type] = n;
      if (!(n in numToType)) numToType[n] = type;
    }
  });
  Logger.log('typeToNum: ' + JSON.stringify(typeToNum));
  Logger.log('numToType: ' + JSON.stringify(numToType));

  // ② type → text(Drive): 従来通り
  const textLastRow = textSheet.getLastRow();
  const textRange = textLastRow >= 2
    ? textSheet.getRange(2, 1, textLastRow - 1, 2)
    : null;

  const typeToText = {};
  if (textRange) {
    const values = textRange.getValues();
    const rich   = textRange.getRichTextValues();

    for (let i = 0; i < values.length; i++) {
      const type = String(values[i][0] || '').trim();
      if (!type) continue;

      let url = rich[i][1]?.getLinkUrl();
      if (!url) url = String(values[i][1] || '').trim();
      if (!url) continue;

      let fileId = url.match(/[-\w]{25,}/)?.[0];
      if (!fileId) {
        const files = DriveApp.getFilesByName(url);
        if (files.hasNext()) fileId = files.next().getId();
      }
      if (!fileId) continue;

      try {
        const txt = DriveApp.getFileById(fileId).getBlob().getDataAsString('UTF-8');
        typeToText[type] = txt;
      } catch (e) {
        Logger.log('テキスト取得失敗: type=' + type + ' error=' + e);
      }
    }
  }

  Logger.log('=== 配送マスタ読み込み 終了 ===');
  return { postageToType, typeToText, typeToNum, numToType };
}


/**
 * R列テキストから配送種別を抽出する
 * "ゆうパケ_1(450)" → "ゆうパケ_1"
 * "佐川"            → "佐川"
 */
function 配送種別を抽出_(rValue) {
  const s = String(rValue || '').trim();
  if (!s) return '';
  return s.replace(/[（(].*/g, '').trim();
}


/**
 * 配送種別テキスト→管理番号を解決
 *  佐川        → 2
 *  ゆうパケ_1  → 7
 *  ゆうパケ_2  → 6
 *  ゆうパケ_3  → 3
 */
function 管理番号を解決_(種別, typeToNum) {
  if (!種別) return null;
  if (種別 === '佐川') return 2;
  if (種別 in typeToNum) return typeToNum[種別];
  return null;
}


/**
 * ✅ v3.1: 配送種別から配送テキストを検索する
 * 検索順: 完全一致 → ベース名フォールバック（ゆうパケ_1→ゆうパケ）
 */
function 配送テキストを検索_(種別, typeToText) {
  if (!種別) return null;
  // 完全一致
  if (typeToText[種別]) return typeToText[種別];
  // ベース名フォールバック: "ゆうパケ_1" → "ゆうパケ"
  if (種別.includes('_')) {
    const ベース名 = 種別.split('_')[0];
    if (typeToText[ベース名]) return typeToText[ベース名];
  }
  return null;
}


// ====== エントリポイント ======
function 配送テキストを反映_商品入力シート対象() {
  配送テキストを反映_汎用_(SHEET_商品入力, HEADER_商品入力);
}
function 配送テキストを反映_Yahoo商品登録シート対象() {
  配送テキストを反映_汎用_(SHEET_Yahoo商品登録, HEADER_Yahoo商品登録);
}


/**
 * 配送テキスト反映 共通
 * 
 * ✅ v3.1: R列がテキストでも数字でも全処理を実行
 *   テキスト（佐川/ゆうパケ_1等）→ B列照合で種別特定 → R列を数字に変換
 *   数字（2/7等）              → D列→B列逆引きで種別特定 → R列はそのまま
 *   どちらの場合も K列設定 + I列/Q列テキスト反映 を行う
 */
function 配送テキストを反映_汎用_(sheetName, headerRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    SHIP_uiSafeAlert_('シートが見つかりません: ' + sheetName);
    return;
  }

  const { postageToType, typeToText, typeToNum, numToType } = 配送マスタを読み込み_();

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= headerRows) {
    SHIP_uiSafeAlert_('データ行がありません: ' + sheetName);
    return;
  }

  const range = sheet.getRange(headerRows + 1, 1, lastRow - headerRows, lastCol);
  const data = range.getValues();

  const idxI = COL_ABSTRACT - 1;
  const idxQ = COL_SP_ADDITIONAL - 1;
  const idxR = COL_POSTAGE_SET - 1;
  const idxK = COL_SHIP_WEIGHT - 1;

  let textCount = 0;
  let weightCount = 0;
  let numCount = 0;

  for (let r = 0; r < data.length; r++) {
    const rRaw = data[r][idxR];
    let 種別 = '';
    let R列変換する = false;

    if (typeof rRaw === 'number') {
      /* ✅ v3.1: R列が数字 → D列→B列逆引きで配送種別を特定
       * 例: 2 → numToType[2] → "佐川"
       *     7 → numToType[7] → "ゆうパケ_1" */
      種別 = numToType[rRaw] || '';
      R列変換する = false;  // 既に数字なので変換不要
    } else {
      const rText = String(rRaw || '').trim();
      if (!rText) continue;
      // "ゆうパケ_1(450)" → "ゆうパケ_1"
      種別 = 配送種別を抽出_(rText);
      R列変換する = true;   // テキスト→数字に変換する
    }

    if (!種別) continue;

    // ---- ① K列: 重量設定（先に実行） ----
    if (種別 === '佐川') {
      data[r][idxK] = 1000;
      weightCount++;
    } else if (種別.startsWith('ゆうパケ')) {
      data[r][idxK] = 100;
      weightCount++;
    }

    // ---- ② I列/Q列: 配送テキスト反映 ----
    let txt = 配送テキストを検索_(種別, typeToText);
    if (!txt) {
      // 旧方式フォールバック: rRawをpostage-set番号として扱う
      const rText = String(rRaw || '').trim();
      const type = postageToType[rText];
      if (type) txt = typeToText[type];
    }
    if (txt) {
      data[r][idxI] = txt;
      data[r][idxQ] = txt;
      textCount++;
    }

    // ---- ③ R列: 管理番号(数字)に変換（テキストの場合のみ） ----
    if (R列変換する) {
      const num = 管理番号を解決_(種別, typeToNum);
      if (num !== null) {
        data[r][idxR] = num;
        numCount++;
      }
    }
  }

  range.setValues(data);

  SHIP_uiSafeAlert_(
    '配送テキスト反映完了\n' +
    'テキスト更新: ' + textCount + '行\n' +
    '重量設定: ' + weightCount + '行\n' +
    '管理番号変換: ' + numCount + '行'
  );
}
