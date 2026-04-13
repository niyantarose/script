/******************************************************
 * 21_description_validation.gs - 商品説明 整形＆バリデーション（Yahoo計算式完全版）
 * v3: 整形機能追加（■前に空行、■行は1行にまとめる）
 * 
 * 【制限】HTML不可 / 全角500文字（1000バイト）以内
 ******************************************************/

/** ===================================================
 * 設定エリア
 * =================================================== */
var VAL_CONFIG = {
  LIMIT_BYTES: 1000, // 制限バイト数（全角500文字相当）
  TARGET_COL: 10,    // J列
  
  COLOR_ERROR: '#ff9999', // エラー色（赤）
  COLOR_OK: null          // OK色（クリア）
};

// ====================================================================
// ★ 安全なUI通知（UIが無くても落ちない）
// ====================================================================

function DESC_uiSafeAlert_(msg) {
  try {
    SpreadsheetApp.getUi().alert(String(msg));
    return;
  } catch (e) {}

  // フォールバック: toast → log
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(String(msg), '商品説明チェック', 10);
  } catch (e2) {}

  Logger.log('[DESC_UI_FALLBACK] ' + msg);
}

function DESC_uiSafeConfirm_(title, msg) {
  try {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert(title, msg, ui.ButtonSet.OK_CANCEL);
    return response === ui.Button.OK;
  } catch (e) {
    Logger.log('[DESC_UI_FALLBACK] 確認ダイアログをスキップ（UIなし）: ' + title);
    return true;
  }
}


/**
 * 【一括処理】商品説明整形 + チェック
 */
function 商品説明を一括整形してチェック() {
  var confirmed = DESC_uiSafeConfirm_(
    '一括処理の確認',
    '商品説明を整形し、Yahooと同じ計算式でチェックします。\n' +
    '【整形】■の前に空行 / ■行を1行にまとめる\n' +
    '【制限】HTML不可 / 全角500文字（1000バイト）以内\n\n実行しますか?'
  );
  
  if (!confirmed) return;
  
  const results = [];
  const startTime = new Date();
  
  const sheetName = (typeof SHEET_商品入力 !== 'undefined') ? SHEET_商品入力 : '①商品入力シート';
  const headerRows = (typeof HEADER_商品入力 !== 'undefined') ? HEADER_商品入力 : 2;

  try {
    // 1. 整形
    整形商品説明_汎用_(sheetName, headerRows);
    Utilities.sleep(200); 
    
    // 2. チェック
    const checkResult = チェック商品説明_Yahoo式_(sheetName, headerRows);
    
    results.push({
      sheet: sheetName,
      success: true,
      overCount: checkResult.overCount,
      totalRows: checkResult.totalRows
    });
    
  } catch (e) {
    results.push({ sheet: sheetName, success: false, error: e.toString() });
  }
  
  const elapsed = Math.round((new Date() - startTime) / 1000);
  表示一括処理結果_厳密_(results, elapsed);
}

/**
 * 商品説明の整形
 * - 改行統一
 * - 「、」の後に改行
 * - ■の前に空行
 * - ■〜：の後は改行しない（1行にまとめる）
 * - 連続空行をまとめる
 */
function 整形商品説明_汎用_(sheetName, headerRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('シートが見つかりません: ' + sheetName);

  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRows) return;

  const range = sheet.getRange(headerRows + 1, VAL_CONFIG.TARGET_COL, lastRow - headerRows, 1);
  const values = range.getValues();
  
  const newValues = values.map(row => {
    let txt = String(row[0] || '');
    
   　// 1. 改行を \n に統一
　　　txt = txt.replace(/\r\n/g, '\n').replace(/\r/g, '\n');

　　　// 2. 「、」を改行に置き換える（、は削除）★修正
　　　txt = txt.replace(/、/g, '\n');

　　// 3. ■の前に空行を入れる（行頭でない場合）
　　　txt = txt.replace(/([^\n])(■)/g, '$1\n\n$2');
    // 4. ■で始まる行の「：」の後の改行を削除（1行にまとめる）
    txt = txt.replace(/(■[^：\n]*：)\s*\n+/g, '$1');
    
    // 5. 連続した空行（3つ以上の改行）を空行1つにまとめる
    txt = txt.replace(/\n{3,}/g, '\n\n');
    
    // 6. 前後の空白・改行を削除
    txt = txt.trim();
    
    return [txt];
  });

  range.setValues(newValues);
}
/**
 * Yahoo!ショッピング仕様のバイト数計算
 * 全角=2, 半角=1, 改行=2
 */
function 計算バイト数_Yahoo式_(str) {
  if (!str) return 0;
  
  let len = 0;
  for (let i = 0; i < str.length; i++) {
    const c = str.charCodeAt(i);
    
    if (c === 10) {
      len += 2;
    }
    else if ((c >= 0x0 && c <= 0x7e) || (c >= 0xff61 && c <= 0xff9f)) {
      len += 1;
    }
    else {
      len += 2;
    }
  }
  return len;
}

/**
 * チェック実行
 */
function チェック商品説明_Yahoo式_(sheetName, headerRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('シートが見つかりません: ' + sheetName);

  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRows) return { overCount: 0, totalRows: 0, overRows: [] };

  const range = sheet.getRange(headerRows + 1, VAL_CONFIG.TARGET_COL, lastRow - headerRows, 1);
  const values = range.getValues();
  const backgrounds = [];
  const notes = []; 
  
  let overCount = 0;
  const overRows = [];

  const htmlTagRegex = /<\/?([a-zA-Z][a-zA-Z0-9]*)\b[^>]*\/?>/;

  for (let i = 0; i < values.length; i++) {
    const text = String(values[i][0] || '');
    
    const bytes = 計算バイト数_Yahoo式_(text);
    const hasHtml = htmlTagRegex.test(text);
    const isByteOver = bytes > VAL_CONFIG.LIMIT_BYTES;

    if (isByteOver || hasHtml) {
      backgrounds.push([VAL_CONFIG.COLOR_ERROR]);
      overCount++;
      
      let reason = [];
      if (hasHtml) reason.push('HTMLタグ不可');
      if (isByteOver) reason.push('バイト過多(' + bytes + '/' + VAL_CONFIG.LIMIT_BYTES + ')');

      notes.push([reason.join('\n')]); 

      overRows.push({
        row: headerRows + 1 + i,
        reason: reason.join(', ')
      });
    } else {
      backgrounds.push([VAL_CONFIG.COLOR_OK]); 
      notes.push([null]);
    }
  }

  range.setBackgrounds(backgrounds);
  range.setNotes(notes);
  
  return { overCount: overCount, totalRows: values.length, overRows: overRows };
}

/**
 * 結果表示
 */
function 表示一括処理結果_厳密_(results, elapsed) {
  let message = '=== 一括処理完了 (' + elapsed + '秒) ===\n\n';
  message += '【整形】■の前に空行 / ■行を1行にまとめる\n';
  message += '【制限】HTML不可 / 全角500文字（1000バイト）以内\n\n';
  let totalOver = 0;
  
  for (const res of results) {
    if (res.success) {
      message += '✓ ' + res.sheet + '\n';
      if (res.overCount > 0) {
        message += '   ⚠️ エラー: ' + res.overCount + '行\n';
        totalOver += res.overCount;
      } else {
        message += '   OK（色はクリアされました）\n';
      }
    } else {
      message += '✗ ' + res.sheet + ': ' + res.error + '\n';
    }
  }
  
  if (totalOver > 0) {
    message += '\n赤色のセルにカーソルを合わせると理由が出ます。';
  } else {
    message += '\nすべてのデータが正常です。';
  }
  DESC_uiSafeAlert_(message);
}

// 個別メニュー用
function 商品説明バイト数チェック_商品入力シート() {
  const sheetName = (typeof SHEET_商品入力 !== 'undefined') ? SHEET_商品入力 : '①商品入力シート';
  const headerRows = (typeof HEADER_商品入力 !== 'undefined') ? HEADER_商品入力 : 2;
  const res = チェック商品説明_Yahoo式_(sheetName, headerRows);
  const msg = res.overCount > 0 ? 
    '制限超過が ' + res.overCount + ' 行あります。\n【制限】HTML不可 / 全角500文字（1000バイト）以内' : 
    'すべて正常です（背景色はクリアされました）';
  DESC_uiSafeAlert_(msg);
}

function 商品説明を整形してチェック_商品入力シート() {
  const sheetName = (typeof SHEET_商品入力 !== 'undefined') ? SHEET_商品入力 : '①商品入力シート';
  const headerRows = (typeof HEADER_商品入力 !== 'undefined') ? HEADER_商品入力 : 2;
  整形商品説明_汎用_(sheetName, headerRows);
  Utilities.sleep(200);
  const res = チェック商品説明_Yahoo式_(sheetName, headerRows);
  DESC_uiSafeAlert_(
    res.overCount > 0 ? 
      '整形完了。エラー: ' + res.overCount + '件\n【制限】HTML不可 / 全角500文字（1000バイト）以内' : 
      '整形完了。すべて正常です'
  );
}

function 商品説明の色をクリア_商品入力シート() {
  const sheetName = (typeof SHEET_商品入力 !== 'undefined') ? SHEET_商品入力 : '①商品入力シート';
  const headerRows = (typeof HEADER_商品入力 !== 'undefined') ? HEADER_商品入力 : 2;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRows) return;
  
  const range = sheet.getRange(headerRows + 1, VAL_CONFIG.TARGET_COL, lastRow - headerRows, 1);
  range.setBackground(null);
  range.clearNote(); 
  
  DESC_uiSafeAlert_('背景色とメモをクリアしました');
}