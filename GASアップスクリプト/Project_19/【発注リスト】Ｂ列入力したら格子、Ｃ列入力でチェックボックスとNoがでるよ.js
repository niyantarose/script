// =====================================================
// 発注シート：チェックボックス・罫線 設定
// =====================================================

const HATCHU_CFG = {
  SHEET_NAME: '発注',

  START_ROW: 7,

  COL_CHK: 1,           // A列：チェックボックス
  COL_NO: 2,            // B列：No.
  COL_DATE: 3,          // C列：発注日

  BORDER_START_COL: 1,  // A列から
  BORDER_COLS: 28,      // A列〜AB列まで
  MAX_COL: 28           // A列〜AB列まで
};

// =====================================================
// onEdit 統合版
// ※ function onEdit(e) はこれ1つだけにする
// =====================================================

function HATCHU_onEdit_legacy_(e) {
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  const sheetName = sh.getName();
  const range = e.range;

  try {
    // =====================================================
    // 発注シート側
    // =====================================================
    if (sheetName === HATCHU_CFG.SHEET_NAME) {
      SpreadsheetApp.flush();

      const editedStartCol = range.getColumn();
      const editedEndCol = range.getLastColumn();
      const onlyCheckboxCol =
        editedStartCol === HATCHU_CFG.COL_CHK &&
        editedEndCol === HATCHU_CFG.COL_CHK;

      // チェックボックス列Aだけの編集では同期しない
      // チェックを入れた瞬間にリセットされるのを防ぐため
      if (!onlyCheckboxCol) {
        syncCheckboxes_inner_(sh);
      }

      // 編集行の前後も含めて罫線更新
      HATCHU_updateBordersForRows_(
        sh,
        range.getRow(),
        range.getNumRows()
      );

// 消込判定の色更新
// 数式でAA列の判定が変わることがあるので、発注シート全体を更新
if (typeof colorKeshikomiAllRows_ === 'function') {
  SpreadsheetApp.flush();
  colorKeshikomiAllRows_(sh);
}
      return;
    }

    // =====================================================
    // EMSリスト側
    // =====================================================
    if (typeof EMS_isTargetSheet_ !== 'function') return;
    if (typeof EMS_CFG === 'undefined') return;
    if (!EMS_isTargetSheet_(sh)) return;

    const hitDate = EMS_rangeHitsColumn_(range, EMS_CFG.COL_EMS_DATE);
    const hitEms = EMS_rangeHitsColumn_(range, EMS_CFG.COL_EMS_NO);

    const isEmsList =
      typeof KESHIKOMI_COLOR_CFG !== 'undefined' &&
      sheetName === KESHIKOMI_COLOR_CFG.EMS_SHEET_NAME;

    // EMS発送数に関係する列を編集した場合も、発注シートの色を更新
    if (!hitDate && !hitEms) {
      if (isEmsList && typeof colorKeshikomiAllRows_ === 'function') {
        SpreadsheetApp.flush();

        const orderSheet = e.source.getSheetByName(KESHIKOMI_COLOR_CFG.SHEET_NAME);
        if (orderSheet) {
          colorKeshikomiAllRows_(orderSheet);
        }
      }
      return;
    }

    const lock = LockService.getDocumentLock();
    if (!lock.tryLock(1000)) return;

    try {
      if (hitDate) {
        EMS_updateStatusOnlyForEditedRows_(sh, range);
      }

      if (EMS_CFG.AUTO_BOX_ON_EDIT && range.getNumRows() <= EMS_CFG.ONEDIT_MAX_ROWS) {
        EMS_updateDatesByRows_(
          sh,
          range.getRow(),
          range.getNumRows(),
          false
        );
      } else if (EMS_CFG.AUTO_BOX_ON_EDIT) {
        SpreadsheetApp.getActive().toast(
          '大量編集のため、ボタンから全体更新を実行してください。'
        );
      }

      // EMSリスト編集後、発注シートの色も更新
      if (isEmsList && typeof colorKeshikomiAllRows_ === 'function') {
        SpreadsheetApp.flush();

        const orderSheet = e.source.getSheetByName(KESHIKOMI_COLOR_CFG.SHEET_NAME);
        if (orderSheet) {
          colorKeshikomiAllRows_(orderSheet);
        }
      }

    } finally {
      lock.releaseLock();
    }

  } catch (err) {
    SpreadsheetApp.getActive().toast(
      'onEditエラー: ' + err.message,
      'エラー',
      8
    );
    throw err;
  }
}


// =====================================================
// メニューから手動実行：チェックボックス同期
// =====================================================

function syncCheckboxes() {
  const sh = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(HATCHU_CFG.SHEET_NAME);

  if (!sh) {
    SpreadsheetApp.getUi().alert('発注シートが見つかりません。');
    return;
  }

  syncCheckboxes_inner_(sh);
  SpreadsheetApp.getActive().toast('チェックボックスを同期しました。', '発注', 2);
}


// =====================================================
// 実処理：B列にNo.がある行だけA列チェックボックスを表示
// =====================================================

function syncCheckboxes_inner_(sh) {
  SpreadsheetApp.flush();

  const cfg = HATCHU_CFG;

  const startRow = cfg.START_ROW;
  const maxRows = sh.getMaxRows();
  const numRows = maxRows - startRow + 1;

  if (numRows <= 0) return;

  const noValues = sh
    .getRange(startRow, cfg.COL_NO, numRows, 1)
    .getDisplayValues();

  const checkRange = sh.getRange(startRow, cfg.COL_CHK, numRows, 1);

  // まずA列のチェックボックス設定だけ全部外す
  // 値は一旦残す
  checkRange.clearDataValidations();

  let segStart = null;

  for (let i = 0; i <= numRows; i++) {
    const hasNo =
      i < numRows &&
      String(noValues[i][0] || '').trim() !== '';

    if (hasNo) {
      if (segStart === null) segStart = i;
    } else {
      // No.がある連続区間にチェックボックスを入れる
      if (segStart !== null) {
        sh.getRange(startRow + segStart, cfg.COL_CHK, i - segStart, 1)
          .insertCheckboxes();
        segStart = null;
      }
    }
  }

  // No.がない行のA列は内容も消す
  let blankStart = null;

  for (let i = 0; i <= numRows; i++) {
    const hasNo =
      i < numRows &&
      String(noValues[i][0] || '').trim() !== '';

    if (!hasNo && i < numRows) {
      if (blankStart === null) blankStart = i;
    } else {
      if (blankStart !== null) {
        sh.getRange(startRow + blankStart, cfg.COL_CHK, i - blankStart, 1)
          .clearContent();
        blankStart = null;
      }
    }
  }
}


// =====================================================
// 手動実行：チェックボックスと罫線を全体更新
// =====================================================

function 発注_チェックボックスと罫線を全体更新() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(HATCHU_CFG.SHEET_NAME);

  if (!sh) {
    SpreadsheetApp.getUi().alert('発注シートが見つかりません。');
    return;
  }

  SpreadsheetApp.flush();

  syncCheckboxes_inner_(sh);
  HATCHU_updateAllBorders_(sh);

  if (typeof colorKeshikomiAllRows_ === 'function') {
    colorKeshikomiAllRows_(sh);
  }

  SpreadsheetApp.getActive().toast('発注シートのチェックボックスと罫線を全体更新しました。');
}


// =====================================================
// 手動実行：B列No.がある行だけ格子罫線
// =====================================================

function 発注_番号あり行に格子罫線() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(HATCHU_CFG.SHEET_NAME);

  if (!sh) {
    SpreadsheetApp.getUi().alert('発注シートが見つかりません。');
    return;
  }

  HATCHU_updateAllBorders_(sh);

  SpreadsheetApp.getActive().toast('B列No.がある行に格子罫線を入れました。');
}


// =====================================================
// 全体罫線更新
// いったんA〜AB列の罫線を全部消して、No.がある行だけ格子を入れる
// =====================================================

function HATCHU_updateAllBorders_(sh) {
  const cfg = HATCHU_CFG;

  const startRow = cfg.START_ROW;
  const maxRows = sh.getMaxRows();
  const numRows = maxRows - startRow + 1;

  if (numRows <= 0) return;

  HATCHU_updateBordersCore_(sh, startRow, numRows);
}


// =====================================================
// 編集行周辺だけ罫線更新
// データ削除時の罫線残りを防ぐため、前後2行も含めて処理
// =====================================================

function HATCHU_updateBordersForRows_(sh, startRow, numRows) {
  const cfg = HATCHU_CFG;

  if (numRows <= 0) return;

  const maxRows = sh.getMaxRows();

  const targetStartRow = Math.max(cfg.START_ROW, startRow - 2);
  const targetEndRow = Math.min(maxRows, startRow + numRows + 2);
  const targetNumRows = targetEndRow - targetStartRow + 1;

  if (targetNumRows <= 0) return;

  HATCHU_updateBordersCore_(sh, targetStartRow, targetNumRows);
}


// =====================================================
// 罫線更新の本体
// =====================================================

function HATCHU_updateBordersCore_(sh, startRow, numRows) {
  const cfg = HATCHU_CFG;

  const startCol = cfg.BORDER_START_COL;
  const colCount = cfg.BORDER_COLS;

  const targetRange = sh.getRange(startRow, startCol, numRows, colCount);

  // まず対象範囲の罫線を全部消す
  targetRange.setBorder(
    false,
    false,
    false,
    false,
    false,
    false
  );

  const noValues = sh
    .getRange(startRow, cfg.COL_NO, numRows, 1)
    .getDisplayValues();

  let segStart = null;

  for (let i = 0; i <= numRows; i++) {
    const hasNo =
      i < numRows &&
      String(noValues[i][0] || '').trim() !== '';

    if (hasNo) {
      if (segStart === null) segStart = i;
    } else {
      if (segStart !== null) {
        sh.getRange(startRow + segStart, startCol, i - segStart, colCount)
          .setBorder(
            true,
            true,
            true,
            true,
            true,
            true,
            '#000000',
            SpreadsheetApp.BorderStyle.SOLID
          );

        segStart = null;
      }
    }
  }
}

// =====================================================
// 消込判定による行色変更
// =====================================================

const KESHIKOMI_COLOR_CFG = {
  SHEET_NAME: '発注',
  EMS_SHEET_NAME: 'EMSリスト',

  START_ROW: 7,

  // A列〜AC列まで色をつける
  // ABまでなら 28、ACまでなら 29
  START_COL: 1,
  END_COL: 29,

  // AA列 = 消込判定
  STATUS_COL: 27,

  COLORS: {
    OK: '#f4b183',           // オレンジ：消込OK
    KOREA_REMAIN: '#ffff00'  // 黄色：韓国残
  }
};


/**
 * 編集された行だけ色更新
 */
function applyKeshikomiColorForEditedRows_(sh, editStartRow, editNumRows) {
  const cfg = KESHIKOMI_COLOR_CFG;

  const startRow = Math.max(editStartRow, cfg.START_ROW);
  const endRow = editStartRow + editNumRows - 1;
  const numRows = endRow - startRow + 1;

  if (numRows <= 0) return;

  colorKeshikomiRows_(sh, startRow, numRows);
}


/**
 * 発注シート全体の色更新
 */
function colorKeshikomiAllRows_(sh) {
  const cfg = KESHIKOMI_COLOR_CFG;

  const lastRow = sh.getLastRow();
  if (lastRow < cfg.START_ROW) return;

  const numRows = lastRow - cfg.START_ROW + 1;
  colorKeshikomiRows_(sh, cfg.START_ROW, numRows);
}


/**
 * AA列の消込判定を見て行色をつける
 */
/**
 * AA列の消込判定を見て行色をつける
 * 消込OK → オレンジ
 * 韓国残 → 黄色
 * 韓国未入荷・未入荷・空欄 → 色なし
 */
function colorKeshikomiRows_(sh, startRow, numRows) {
  const cfg = KESHIKOMI_COLOR_CFG;

  if (!sh) return;
  if (numRows <= 0) return;

  const startCol = cfg.START_COL || 1;
  const endCol = cfg.END_COL || 29;
  const width = endCol - startCol + 1;

  const statusValues = sh
    .getRange(startRow, cfg.STATUS_COL, numRows, 1)
    .getDisplayValues();

  for (let i = 0; i < numRows; i++) {
    const rowNo = startRow + i;
    const status = String(statusValues[i][0] || '').trim();

    let color = null;

    if (status === '消込OK') {
      color = cfg.COLORS.OK; // オレンジ
    } else if (status.indexOf('韓国残') !== -1) {
      color = cfg.COLORS.KOREA_REMAIN; // 黄色
    } else {
      color = null; // 色なし
    }

    const rowRange = sh.getRange(rowNo, startCol, 1, width);

    if (color) {
      rowRange.setBackground(color);
    } else {
      rowRange.setBackground(null);
    }
  }
}


/**
 * 手動更新用
 */
function 消込判定の色を更新する() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(KESHIKOMI_COLOR_CFG.SHEET_NAME);
  if (!sh) return;

  SpreadsheetApp.flush();
  colorKeshikomiAllRows_(sh);

  SpreadsheetApp.getActive().toast('消込判定の色を更新しました。');
}

// =====================================================
// 発注チェック行 → EMSリストへ転送
// =====================================================

const HATCHU_TO_EMS_CFG = {
  SRC_SHEET: '発注',
  DST_SHEET: 'EMSリスト',

  HEADER_ROWS: [5, 6],
  START_ROW: 7,

  SRC_COL_CHECK: 1, // A列チェックボックス

  // EMSリスト側：書き込む列だけ
  DST_COL_NO: 1,           // A列 No.
  DST_COL_ARRIVAL: 2,      // B列 入荷日
  DST_COL_EMS_DATE: 3,     // C列 EMS発送日
  DST_COL_ARROW: 4,        // D列 ⇒
  DST_COL_PURCHASE_NO: 6,  // F列 購入No.
  DST_COL_STATUS: 7,       // G列 ステータス列
  DST_COL_CODE: 9,         // I列 商品コード
  DST_COL_QTY: 10,         // J列 数量

  // K列 品目はEMSリスト側の関数で出すので触らない
  // O列 照合キーもEMSリスト側の関数で出すので触らない

  UNCHECK_AFTER_SEND: true,
  SKIP_DUPLICATES: true
};


function 発注_チェック行をEMSリストへ送る() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(HATCHU_TO_EMS_CFG.SRC_SHEET);
  const dst = ss.getSheetByName(HATCHU_TO_EMS_CFG.DST_SHEET);

  if (!src) {
    SpreadsheetApp.getUi().alert('発注シートが見つかりません。');
    return;
  }

  if (!dst) {
    SpreadsheetApp.getUi().alert('EMSリストシートが見つかりません。');
    return;
  }

  const cfg = HATCHU_TO_EMS_CFG;
  const startRow = cfg.START_ROW;
  const srcLastRow = src.getLastRow();
  const srcLastCol = src.getLastColumn();

  if (srcLastRow < startRow) {
    SpreadsheetApp.getActive().toast('発注シートに転送対象行がありません。');
    return;
  }

  // 発注シートの列を見出しから探す
  const colPurchaseNo = H2E_findHeaderCol_(src, ['OrderNo', '購入No.', '購入No']);
  const colCode       = H2E_findHeaderCol_(src, ['Code', '商品コード']);
  const colQty        = H2E_findHeaderCol_(src, ['Units', '数量']);

  if (!colPurchaseNo || !colCode || !colQty) {
    SpreadsheetApp.getUi().alert(
      '必要な列が見つかりません。\n' +
      '発注シートに「OrderNo / 購入No.」「Code / 商品コード」「Units / 数量」があるか確認してください。'
    );
    return;
  }

  const numRows = srcLastRow - startRow + 1;

  const checkValues = src
    .getRange(startRow, cfg.SRC_COL_CHECK, numRows, 1)
    .getValues();

  const srcValues = src
    .getRange(startRow, 1, numRows, srcLastCol)
    .getDisplayValues();

  const existingKeys = H2E_getExistingEmsKeys_(dst);
  const today = H2E_today_();

  const appendRows = [];
  const copiedSrcRowNumbers = [];

  let skippedDuplicate = 0;
  let skippedNoData = 0;
  let nextNo = H2E_getNextEmsNo_(dst);

  for (let i = 0; i < numRows; i++) {
    const checked = checkValues[i][0] === true;
    if (!checked) continue;

    const row = srcValues[i];

    const purchaseNo = String(row[colPurchaseNo - 1] || '').trim();
    const code = String(row[colCode - 1] || '').trim();
    const qty = String(row[colQty - 1] || '').trim();

    if (!purchaseNo || !code) {
      skippedNoData++;
      continue;
    }

    const key = purchaseNo + '-' + code;

    if (cfg.SKIP_DUPLICATES && existingKeys.has(key)) {
      skippedDuplicate++;
      continue;
    }

    appendRows.push({
      no: nextNo++,
      arrivalDate: today,
      emsDate: today,
      arrow: '⇒',
      purchaseNo: purchaseNo,
      status: '未着',
      code: code,
      qty: qty
    });

    copiedSrcRowNumbers.push(startRow + i);
    existingKeys.add(key);
  }

  if (appendRows.length === 0) {
    SpreadsheetApp.getActive().toast(
      `転送対象なし：重複 ${skippedDuplicate}件 / データ不足 ${skippedNoData}件`
    );
    return;
  }

  const appendStartRow = H2E_findNextAppendRow_(dst);

  // A:D に書く
  const valuesAD = appendRows.map(r => [
    r.no,
    r.arrivalDate,
    r.emsDate,
    r.arrow
  ]);

  dst.getRange(appendStartRow, 1, valuesAD.length, 4)
    .setValues(valuesAD);

  // F:G に書く
  const valuesFG = appendRows.map(r => [
    r.purchaseNo,
    r.status
  ]);

  dst.getRange(appendStartRow, 6, valuesFG.length, 2)
    .setValues(valuesFG);

  // I:J に書く
  const valuesIJ = appendRows.map(r => [
    r.code,
    r.qty
  ]);

  dst.getRange(appendStartRow, 9, valuesIJ.length, 2)
    .setValues(valuesIJ);

  // 日付表示
  dst.getRange(appendStartRow, 2, appendRows.length, 2)
    .setNumberFormat('yyyy/m/d');

  // コピー後にチェックを外す
  if (cfg.UNCHECK_AFTER_SEND) {
    copiedSrcRowNumbers.forEach(rowNo => {
      src.getRange(rowNo, cfg.SRC_COL_CHECK).setValue(false);
    });
  }

  SpreadsheetApp.getActive().toast(
    `EMSリストへ転送しました：${appendRows.length}件 / 重複スキップ ${skippedDuplicate}件 / データ不足 ${skippedNoData}件`
  );
}


// =====================================================
// 補助関数
// =====================================================

function H2E_findHeaderCol_(sh, names) {
  const cfg = HATCHU_TO_EMS_CFG;
  const lastCol = sh.getLastColumn();

  const normalizedNames = names.map(H2E_normHeader_);

  for (const headerRow of cfg.HEADER_ROWS) {
    const values = sh
      .getRange(headerRow, 1, 1, lastCol)
      .getDisplayValues()[0];

    for (let c = 0; c < values.length; c++) {
      const h = H2E_normHeader_(values[c]);
      if (normalizedNames.includes(h)) {
        return c + 1;
      }
    }
  }

  return 0;
}


function H2E_normHeader_(v) {
  return String(v || '')
    .replace(/\s/g, '')
    .replace(/　/g, '')
    .trim()
    .toLowerCase();
}


function H2E_today_() {
  const now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), now.getDate());
}


function H2E_getExistingEmsKeys_(dst) {
  const cfg = HATCHU_TO_EMS_CFG;
  const startRow = cfg.START_ROW;
  const lastRow = dst.getLastRow();
  const set = new Set();

  if (lastRow < startRow) return set;

  const numRows = lastRow - startRow + 1;

  const purchaseValues = dst
    .getRange(startRow, cfg.DST_COL_PURCHASE_NO, numRows, 1)
    .getDisplayValues();

  const codeValues = dst
    .getRange(startRow, cfg.DST_COL_CODE, numRows, 1)
    .getDisplayValues();

  for (let i = 0; i < numRows; i++) {
    const purchaseNo = String(purchaseValues[i][0] || '').trim();
    const code = String(codeValues[i][0] || '').trim();

    if (purchaseNo && code) {
      set.add(purchaseNo + '-' + code);
    }
  }

  return set;
}


function H2E_getNextEmsNo_(dst) {
  const cfg = HATCHU_TO_EMS_CFG;
  const startRow = cfg.START_ROW;
  const lastRow = dst.getLastRow();

  if (lastRow < startRow) return 1;

  const numRows = lastRow - startRow + 1;
  const values = dst
    .getRange(startRow, cfg.DST_COL_NO, numRows, 1)
    .getDisplayValues();

  let maxNo = 0;

  values.forEach(r => {
    const n = Number(String(r[0] || '').trim());
    if (!isNaN(n) && n > maxNo) maxNo = n;
  });

  return maxNo + 1;
}


function H2E_findNextAppendRow_(dst) {
  const cfg = HATCHU_TO_EMS_CFG;
  const startRow = cfg.START_ROW;
  const maxRows = dst.getMaxRows();

  // K列・O列は関数列なので見ない
  // M列EMS番号も未入力のことが多いので見ない
  const checkCols = [
    cfg.DST_COL_PURCHASE_NO, // F列
    cfg.DST_COL_CODE,        // I列
    cfg.DST_COL_QTY          // J列
  ];

  let lastDataRow = startRow - 1;

  checkCols.forEach(col => {
    const values = dst
      .getRange(startRow, col, maxRows - startRow + 1, 1)
      .getDisplayValues();

    for (let i = values.length - 1; i >= 0; i--) {
      if (String(values[i][0] || '').trim() !== '') {
        const rowNo = startRow + i;
        if (rowNo > lastDataRow) lastDataRow = rowNo;
        break;
      }
    }
  });

  return lastDataRow + 1;
}
