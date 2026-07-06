function onEdit(e) {
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  const sheetName = sh.getName();
  const range = e.range;

  try {
    // =====================================================
    // 大邱未作業データ：検索セル（1〜2行目）の編集で自動絞り込み
    // =====================================================
    if (typeof MISAGYO_CFG !== 'undefined' && sheetName === MISAGYO_CFG.DST_SHEET) {
      大邱未作業_onEdit_(e);
      return;
    }

    // =====================================================
    // 発注リスト大邱データ：F列に発注NOが入ったら一意の連番を付ける
    // =====================================================
    if (sheetName === DAEGU_CFG.HACHU_SRC) {
      _autoFillArrivalDateFromQty_(sh, range, 4, 3, 6, 1, 6); // D入荷数 -> C入荷日 + Aチェック
      autofillHachuByCfg_(e, DAEGU_HACHU_MASTER_CFG);
      // H業者/I商品名/K商品コード/N品目/O重さ/P価格の編集 → 編集行の最新値で商品マスタを即更新
      if (_rangeHitsAnyCol_(range.getColumn(), range.getLastColumn(), [8, 9, 11, 14, 15, 16])) {
        大邱_マスタ自動更新_(sh, range);
      }
      大邱発注_onEdit採番_(e);
      大邱発注_WX自動計算_(sh, range);
      // C入荷日/D入荷数の編集 → 送信済みのEMS大邱行(A列空欄)へ入荷日を自動反映
      if (_rangeHitsAnyCol_(range.getColumn(), range.getLastColumn(), [3, 4]) &&
          typeof EMS大邱_入荷日補完_ === 'function') {
        EMS大邱_入荷日補完_();
      }
      if (typeof 大邱未作業_同期予約_ === 'function') 大邱未作業_同期予約_(); // 未作業リストの自動同期を予約
      return;
    }

    // =====================================================
    // EMS大邱側：EMS番号/商品コード/数量/購入Noが変わったら韓国側残り数量を再計算
    // =====================================================
    if (sheetName === DAEGU_CFG.EMS_SRC) {
      const editStartCol = range.getColumn();
      const editEndCol = range.getLastColumn();
      if (_rangeHitsAnyCol_(editStartCol, editEndCol, [9, 12, 13])) {
        EMS大邱_QRS自動計算_(sh, range);
      }
      if (_rangeHitsAnyCol_(editStartCol, editEndCol, [4, 8, 9, 20])) {
        大邱発注_チェックと残り数量を設置();
      }
      if (typeof 大邱未作業_同期予約_ === 'function') 大邱未作業_同期予約_(); // 残り数量の変化は未作業判定に影響
      return;
    }

    // =====================================================
    // 発注シート側の処理
    // =====================================================
    if (sheetName === HATCHU_CFG.SHEET_NAME) {
      const editStartCol = range.getColumn();
      const editEndCol = range.getLastColumn();
      const editStartRow = range.getRow();
      const editNumRows = range.getNumRows();
      const hitPurchaseNo = _rangeHitsCol_(editStartCol, editEndCol, HATCHU_GROUP_COL);

      _autoFillArrivalDateFromQty_(sh, range, 5, 4, 6, HATCHU_CFG.COL_CHK, HATCHU_GROUP_COL); // E入荷数 -> D入荷日 + Aチェック

      if (hitPurchaseNo) {
        発注_onEdit購入No採番_(e);
      }

      if (
        _rangeHitsCol_(editStartCol, editEndCol, CFG.HACHU_CODE) ||
        _rangeHitsCol_(editStartCol, editEndCol, CFG.HACHU_VENDOR)
      ) {
        autofillHachu(e);
      }

      商品コードと業者と商品名と価格がそろったら自動でマスタ登録(e);

      const onlyCheckboxCol =
        editStartCol === HATCHU_CFG.COL_CHK &&
        editEndCol === HATCHU_CFG.COL_CHK;

      let refreshGroupBorders = hitPurchaseNo;
      if (!onlyCheckboxCol) {
        const startRow = Math.max(HATCHU_CFG.START_ROW, editStartRow - 2);
        const endRow = Math.min(
          sh.getMaxRows(),
          editStartRow + editNumRows + 2
        );
        const numRows = endRow - startRow + 1;

        if (numRows > 0) {
          HATCHU_processBordersAndCheckboxes_(sh, startRow, numRows);
          refreshGroupBorders = true;
        }
      }

      if (typeof applyKeshikomiColorForEditedRows_ === 'function') {
        SpreadsheetApp.flush();
        applyKeshikomiColorForEditedRows_(sh, editStartRow, editNumRows);
      }

      if (refreshGroupBorders) {
        発注_drawGroupBorders_(sh);
      }

      return;
    }

    // =====================================================
    // EMSリスト側の処理
    // =====================================================
    if (typeof EMS_isTargetSheet_ !== 'function') return;
    if (typeof EMS_CFG === 'undefined') return;
    if (!EMS_isTargetSheet_(sh)) return;

    const hitDate = EMS_rangeHitsColumn_(range, EMS_CFG.COL_EMS_DATE);
    const hitEms = EMS_rangeHitsColumn_(range, EMS_CFG.COL_EMS_NO);

    const isEmsList =
      typeof KESHIKOMI_COLOR_CFG !== 'undefined' &&
      sheetName === KESHIKOMI_COLOR_CFG.EMS_SHEET_NAME;
    const hitShippingCount =
      isEmsList &&
      _rangeHitsAnyCol_(range.getColumn(), range.getLastColumn(), [6, 9, 10, 13]);

    if (!hitDate && !hitEms) {
      if (hitShippingCount && typeof 発注_EMS発送数数式を一括修正 === 'function') {
        発注_EMS発送数数式を一括修正();
      } else if (isEmsList && typeof colorKeshikomiAllRows_ === 'function') {
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

      if (
        EMS_CFG.AUTO_BOX_ON_EDIT &&
        range.getNumRows() <= EMS_CFG.ONEDIT_MAX_ROWS
      ) {
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

      if (hitShippingCount && typeof 発注_EMS発送数数式を一括修正 === 'function') {
        発注_EMS発送数数式を一括修正();
      } else if (isEmsList && typeof colorKeshikomiAllRows_ === 'function') {
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
function _rangeHitsCol_(startCol, endCol, targetCol) {
  targetCol = Number(targetCol);
  return !!targetCol && startCol <= targetCol && targetCol <= endCol;
}

function _rangeHitsAnyCol_(startCol, endCol, targetCols) {
  return targetCols.some(col => _rangeHitsCol_(startCol, endCol, col));
}

function EMS大邱_QRS自動計算_(sh, range) {
  if (!sh || !range || sh.getName() !== DAEGU_CFG.EMS_SRC) return 0;
  const rawEndRow = range.getLastRow();
  if (rawEndRow < 3) return 0;
  const startRow = Math.max(3, range.getRow());
  return EMS大邱_QRS行範囲を再計算_(sh, startRow, rawEndRow - startRow + 1);
}

function EMS大邱_QRSを全体再計算() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(DAEGU_CFG.EMS_SRC);
  if (!sh) {
    SpreadsheetApp.getUi().alert('「' + DAEGU_CFG.EMS_SRC + '」が見つかりません。');
    return;
  }

  const lastRow = EMS大邱_QRS最終データ行_(sh);
  if (lastRow < 3) {
    ss.toast('EMS大邱作業データ: 再計算対象がありません。');
    return 0;
  }

  const count = EMS大邱_QRS行範囲を再計算_(sh, 3, lastRow - 2);
  ss.toast('EMS大邱 N/Q/R/Sを再計算しました: ' + count + '行');
  return count;
}

function EMS大邱_NQRSを全体再計算() {
  return EMS大邱_QRSを全体再計算();
}

function EMS大邱_QRS最終データ行_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 3) return 2;
  const numRows = lastRow - 2;
  const values = sh.getRange(3, 1, numRows, 20).getDisplayValues();
  const checkIndexes = [0, 3, 7, 8, 11, 12, 19]; // A/D/H/I/L/M/T

  for (let i = values.length - 1; i >= 0; i--) {
    if (checkIndexes.some(idx => String(values[i][idx] || '').trim() !== '')) {
      return 3 + i;
    }
  }
  return 2;
}

function EMS大邱_QRS行範囲を再計算_(sh, startRow, numRows) {
  if (!sh || numRows <= 0) return 0;
  const firstRow = Math.max(3, startRow);
  const count = numRows - (firstRow - startRow);
  if (count <= 0) return 0;

  const rows = sh.getRange(firstRow, 1, count, 13).getDisplayValues();
  const nValues = [];
  const qrsValues = rows.map(row => {
    const qty = EMS大邱_QRS数値_(row[8]);       // I 数量
    const weight = EMS大邱_QRS数値_(row[11]);  // L weight(g)
    const unitPrice = EMS大邱_QRS数値_(row[12]); // M Unit Price

    const amountWon = (qty !== null && unitPrice !== null) ? unitPrice * qty : '';
    const totalWeight = (qty !== null && weight !== null) ? qty * weight : '';
    const billedPrice = (unitPrice !== null) ? unitPrice * 0.1 : '';
    const billedAmount = (qty !== null && billedPrice !== '') ? billedPrice * qty : '';

    nValues.push([amountWon]);
    return [totalWeight, billedPrice, billedAmount];
  });

  sh.getRange(firstRow, 14, nValues.length, 1).setValues(nValues); // N Amount(Won)
  sh.getRange(firstRow, 17, qrsValues.length, 3).setValues(qrsValues); // Q:R:S
  return qrsValues.length;
}

function EMS大邱_QRS数値_(value) {
  if (typeof value === 'number' && isFinite(value)) return value;
  const text = String(value == null ? '' : value)
    .normalize('NFKC')
    .replace(/,/g, '')
    .replace(/[\s\u3000]+/g, '')
    .replace(/[^\d.+-]/g, '');
  if (!text || text === '-' || text === '+' || text === '.' || text === '+.' || text === '-.') return null;
  const num = Number(text);
  return isFinite(num) ? num : null;
}

function 大邱発注_WX自動計算_(sh, range) {
  if (!sh || !range || sh.getName() !== DAEGU_CFG.HACHU_SRC) return 0;

  const sc = range.getColumn();
  const ec = range.getLastColumn();
  const editedW3 = range.getRow() <= 3 && range.getLastRow() >= 3 && sc <= 23 && ec >= 23;
  if (editedW3) return 大邱発注_WXを全体再計算_シート_(sh, false);

  if (_rangeHitsAnyCol_(sc, ec, [
    6,  // F 発注NO
    12, // L 数量
    16, // P 価格
    DAEGU_HACHU_MASTER_CFG.HACHU_CODE,
    DAEGU_HACHU_MASTER_CFG.HACHU_VENDOR
  ])) {
    return 大邱発注_WXを全体再計算_シート_(sh, false);
  }
  return 0;
}

function 大邱発注_WXを全体再計算() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(DAEGU_CFG.HACHU_SRC);
  if (!sh) {
    SpreadsheetApp.getUi().alert('「' + DAEGU_CFG.HACHU_SRC + '」が見つかりません。');
    return;
  }
  return 大邱発注_WXを全体再計算_シート_(sh, true);
}

function 大邱発注_QSWXを全体再計算() {
  return 大邱発注_WXを全体再計算();
}

function 大邱発注_WXを全体再計算_シート_(sh, notify) {
  const lastRow = 大邱発注_WX最終データ行_(sh);
  if (lastRow < 6) {
    if (notify) SpreadsheetApp.getActive().toast('発注リスト大邱データ: 再計算対象がありません。');
    return 0;
  }

  const count = 大邱発注_WX行範囲を再計算_(sh, 6, lastRow - 5);
  if (notify) SpreadsheetApp.getActive().toast('発注リスト大邱データ Q/S/W/Xを再計算しました: ' + count + '行');
  return count;
}

function 大邱発注_WX最終データ行_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 6) return 5;
  const numRows = lastRow - 5;
  const values = sh.getRange(6, 1, numRows, 24).getDisplayValues();
  const checkIndexes = [5, 10, 11, 15, 16, 18, 22, 23]; // F/K/L/P/Q/S/W/X

  for (let i = values.length - 1; i >= 0; i--) {
    if (checkIndexes.some(idx => String(values[i][idx] || '').trim() !== '')) {
      return 6 + i;
    }
  }
  return 5;
}

function 大邱発注_WX行範囲を再計算_(sh, startRow, numRows) {
  if (!sh || numRows <= 0) return 0;
  const firstRow = Math.max(6, startRow);
  const count = numRows - (firstRow - startRow);
  if (count <= 0) return 0;

  const rate = 大邱発注_WX数値_(sh.getRange(3, 23).getValue()); // W3
  const rows = sh.getRange(firstRow, 1, count, 24).getDisplayValues();
  const qValues = [];
  const sValues = [];
  const wxValues = [];
  const calcRows = rows.map(row => {
    const qty = 大邱発注_WX数値_(row[11]);   // L 数量
    const price = 大邱発注_WX数値_(row[15]); // P 価格
    const orderNo = row[5];                  // F 発注NO

    const rowSubtotal = (qty !== null && price !== null) ? qty * price : '';
    const billedPrice = (price !== null && rate !== null) ? price * rate : '';
    const billedAmount = (billedPrice !== '' && qty !== null) ? billedPrice * qty : '';
    const key = (typeof 大邱発注_カートグループキー_ === 'function')
      ? 大邱発注_カートグループキー_(orderNo)
      : '';

    qValues.push([rowSubtotal]);
    sValues.push(['']);
    wxValues.push([billedPrice, billedAmount]);
    return { key: key, subtotal: rowSubtotal };
  });

  let curKey = '';
  let groupLast = -1;
  let groupSum = 0;
  let groupHasAmount = false;
  const flushGroup = () => {
    if (curKey && groupLast >= 0 && groupHasAmount) {
      sValues[groupLast][0] = groupSum;
    }
  };

  for (let i = 0; i < calcRows.length; i++) {
    const row = calcRows[i];
    if (!row.key) {
      flushGroup();
      curKey = '';
      groupLast = -1;
      groupSum = 0;
      groupHasAmount = false;
      continue;
    }

    if (row.key !== curKey) {
      flushGroup();
      curKey = row.key;
      groupSum = 0;
      groupHasAmount = false;
    }

    if (row.subtotal !== '') {
      groupSum += row.subtotal;
      groupHasAmount = true;
    }
    groupLast = i;
  }
  flushGroup();

  sh.getRange(firstRow, 17, qValues.length, 1).setValues(qValues); // Q 小計
  sh.getRange(firstRow, 19, sValues.length, 1).setValues(sValues); // S 支払金額
  sh.getRange(firstRow, 23, wxValues.length, 2).setValues(wxValues); // W:X
  return qValues.length;
}

function 大邱発注_WX数値_(value) {
  if (typeof value === 'number' && isFinite(value)) return value;
  const raw = String(value == null ? '' : value).normalize('NFKC');
  const text = raw
    .replace(/,/g, '')
    .replace(/[\s\u3000]+/g, '')
    .replace(/[^\d.+%-]/g, '');
  if (!text || text === '-' || text === '+' || text === '.' || text === '+.' || text === '-.' || text === '%') return null;

  const hasPercent = text.indexOf('%') >= 0;
  const num = Number(text.replace(/%/g, ''));
  if (!isFinite(num)) return null;
  return hasPercent ? num / 100 : num;
}

function _autoFillArrivalDateFromQty_(sh, range, qtyCol, dateCol, startRow, checkboxCol, requiredValueCol) {
  const hitsQty = _rangeHitsCol_(range.getColumn(), range.getLastColumn(), qtyCol);
  const hitsDate = _rangeHitsCol_(range.getColumn(), range.getLastColumn(), dateCol);
  if (!hitsQty && !hitsDate) return;

  // 入荷数あり & 入荷日なし → 入荷日を今日で補完（入荷数列を編集したときだけ・シート全体をバックフィル）
  // freshRows = 今回入荷日を補完した行（=新しく入荷した行）。自動チェックはこの行だけに付ける。
  // ※以前は「入荷数がある行全部」にチェックを付けていたため、広い範囲を編集すると
  //   昔の入荷済み行まで全部チェックが付いてしまっていた。
  let freshRows = null;
  if (hitsQty) {
    freshRows = new Set();
    const firstRow = startRow;
    const lastRow = Math.max(range.getLastRow(), sh.getLastRow());
    if (lastRow >= firstRow) {
      const numRows = lastRow - firstRow + 1;
      const qtyVals = sh.getRange(firstRow, qtyCol, numRows, 1).getDisplayValues();
      const dateRange = sh.getRange(firstRow, dateCol, numRows, 1);
      const dateVals = dateRange.getValues();
      const dateDisplayVals = dateRange.getDisplayValues();
      const today = _todayOnly_();
      let changed = false;
      for (let i = 0; i < numRows; i++) {
        const qty = String(qtyVals[i][0] || '').replace(/,/g, '').trim();
        const hasDate = String(dateDisplayVals[i][0] || '').trim() !== '';
        if (qty && qty !== '0' && !hasDate) {
          dateVals[i][0] = today;
          changed = true;
          freshRows.add(firstRow + i); // この行は「今回入荷」= 自動チェック対象
        }
      }
      if (changed) dateRange.setValues(dateVals);
    }
  }

  // 入荷日と入荷数の連動クリア（編集した行だけ）：片方を消したらもう片方も消す＋チェックも外す
  _linkClearArrivalDateAndQty_(sh, range, qtyCol, dateCol, startRow, hitsQty, hitsDate, checkboxCol);

  // 今回新しく入荷した行（入荷日を補完した行）だけチェックを付ける
  if (checkboxCol && hitsQty) {
    _checkRowsFromEditedQty_(sh, range, qtyCol, checkboxCol, startRow, requiredValueCol, freshRows);
  }
}

// 入荷日(dateCol)と入荷数(qtyCol)を連動クリアする。編集された行だけを対象にする。
//   ・入荷数を消した → 入荷日も消す（＋チェックも外す）
//   ・入荷日を消した → 入荷数も消す（＋チェックも外す）
// ※片方に値が残る編集（数量変更・日付変更）では何も消さない。
// ※スクリプトによるsetValuesはonEditを再発火しないので無限ループしない。
function _linkClearArrivalDateAndQty_(sh, range, qtyCol, dateCol, startRow, hitsQty, hitsDate, checkboxCol) {
  if (!hitsQty && !hitsDate) return;
  const firstRow = Math.max(startRow, range.getRow());
  const lastRow = Math.min(range.getLastRow(), sh.getLastRow());
  if (lastRow < firstRow) return;

  const numRows = lastRow - firstRow + 1;
  const qtyRange = sh.getRange(firstRow, qtyCol, numRows, 1);
  const dateRange = sh.getRange(firstRow, dateCol, numRows, 1);
  const qtyVals = qtyRange.getValues();
  const qtyDisp = qtyRange.getDisplayValues();
  const dateVals = dateRange.getValues();
  const dateDisp = dateRange.getDisplayValues();
  const checkRange = checkboxCol ? sh.getRange(firstRow, checkboxCol, numRows, 1) : null;
  const checkVals = checkRange ? checkRange.getValues() : null;

  let qtyChanged = false, dateChanged = false, checkChanged = false;
  for (let i = 0; i < numRows; i++) {
    const qtyStr = String(qtyDisp[i][0] || '').replace(/,/g, '').trim();
    const hasQty = qtyStr !== '' && qtyStr !== '0';
    const hasDate = String(dateDisp[i][0] || '').trim() !== '';

    let qtyClearedHere = false;
    if (hitsQty && !hasQty && hasDate) {   // 入荷数を消した → 入荷日も消す
      dateVals[i][0] = '';
      dateChanged = true;
    }
    if (hitsDate && !hasDate && hasQty) {   // 入荷日を消した → 入荷数も消す
      qtyVals[i][0] = '';
      qtyChanged = true;
      qtyClearedHere = true;
    }

    // 入荷数が空になった行はA列チェックも外す（チェック済みのときだけ）
    const qtyNowEmpty = (hitsQty && !hasQty) || qtyClearedHere;
    if (qtyNowEmpty && checkVals && checkVals[i][0] === true) {
      checkVals[i][0] = false;
      checkChanged = true;
    }
  }
  if (dateChanged) dateRange.setValues(dateVals);
  if (qtyChanged) qtyRange.setValues(qtyVals);
  if (checkChanged) checkRange.setValues(checkVals);
}

function _checkRowsFromEditedQty_(sh, range, qtyCol, checkboxCol, startRow, requiredValueCol, freshRows) {
  const firstRow = Math.max(startRow, range.getRow());
  const lastRow = Math.min(range.getLastRow(), sh.getLastRow());
  if (lastRow < firstRow) return 0;

  const numRows = lastRow - firstRow + 1;
  const qtyVals = sh.getRange(firstRow, qtyCol, numRows, 1).getDisplayValues();
  const requiredVals = requiredValueCol
    ? sh.getRange(firstRow, requiredValueCol, numRows, 1).getDisplayValues()
    : null;
  const checkRange = sh.getRange(firstRow, checkboxCol, numRows, 1);
  const checkVals = checkRange.getValues();
  const validations = checkRange.getDataValidations();
  const checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  let validationChanged = false;
  let valueChanged = false;
  let checked = 0;

  for (let i = 0; i < numRows; i++) {
    // 今回入荷（入荷日を補完した行）だけを対象にする。昔の入荷済み行にはチェックを付けない。
    if (freshRows && !freshRows.has(firstRow + i)) continue;

    const qty = String(qtyVals[i][0] || '').replace(/,/g, '').trim();
    if (!qty || qty === '0') continue;

    if (requiredVals && String(requiredVals[i][0] || '').trim() === '') continue;

    const rule = validations[i][0];
    if (!rule || rule.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
      validations[i][0] = checkboxRule;
      validationChanged = true;
    }

    if (checkVals[i][0] !== true) {
      checkVals[i][0] = true;
      valueChanged = true;
      checked++;
    }
  }

  if (validationChanged) checkRange.setDataValidations(validations);
  if (valueChanged) checkRange.setValues(checkVals);
  return checked;
}

function _todayOnly_() {
  const now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), now.getDate());
}

function 入荷数あり行_入荷日を補完() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const sh = ss.getActiveSheet();
  const sheetName = sh.getName();
  let cfg = null;

  if (sheetName === DAEGU_CFG.HACHU_SRC) {
    cfg = { label: DAEGU_CFG.HACHU_SRC, startRow: 6, qtyCol: 4, dateCol: 3 };
  } else if (sheetName === CFG.HACHU_SHEET) {
    cfg = { label: CFG.HACHU_SHEET, startRow: 6, qtyCol: 5, dateCol: 4 };
  }

  if (!cfg) {
    ui.alert('発注リスト大邱データ、または発注シートで実行してください。');
    return;
  }

  const count = _fillArrivalDatesForAllQtyRows_(sh, cfg.qtyCol, cfg.dateCol, cfg.startRow);
  ss.toast(cfg.label + ': 入荷日を補完しました ' + count + '行');
}

function _fillArrivalDatesForAllQtyRows_(sh, qtyCol, dateCol, startRow) {
  const lastRow = sh.getLastRow();
  if (lastRow < startRow) return 0;

  const numRows = lastRow - startRow + 1;
  const qtyVals = sh.getRange(startRow, qtyCol, numRows, 1).getDisplayValues();
  const dateRange = sh.getRange(startRow, dateCol, numRows, 1);
  const dateVals = dateRange.getValues();
  const dateDisplayVals = dateRange.getDisplayValues();
  const today = _todayOnly_();
  let changed = 0;

  for (let i = 0; i < numRows; i++) {
    const qty = String(qtyVals[i][0] || '').replace(/,/g, '').trim();
    const hasDate = String(dateDisplayVals[i][0] || '').trim() !== '';
    if (qty && qty !== '0' && !hasDate) {
      dateVals[i][0] = today;
      changed++;
    }
  }

  if (changed) dateRange.setValues(dateVals);
  return changed;
}

function チェック行_入荷数を発注数量にする() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const sh = ss.getActiveSheet();
  const sheetName = sh.getName();
  let cfg = null;

  if (sheetName === DAEGU_CFG.HACHU_SRC) {
    cfg = {
      label: DAEGU_CFG.HACHU_SRC,
      startRow: 6,
      checkCol: 1,
      dateCol: 3,       // C 入荷日
      arrivalQtyCol: 4, // D 入荷数
      orderQtyCol: 12   // L 発注数量
    };
  } else if (sheetName === CFG.HACHU_SHEET) {
    cfg = {
      label: CFG.HACHU_SHEET,
      startRow: 6,
      checkCol: 1,
      dateCol: 4,       // D 入荷日
      arrivalQtyCol: 5, // E 入荷数
      orderQtyCol: 13   // M 発注数量
    };
  }

  if (!cfg) {
    ui.alert('発注リスト大邱データ、または発注シートで実行してください。');
    return;
  }

  const lastRow = sh.getLastRow();
  if (lastRow < cfg.startRow) {
    ss.toast(cfg.label + ': 対象行がありません。');
    return;
  }

  // 自動再計算(onChange)が動いていても最大30秒待って順番を取る
  //（従来は2秒で諦めて「他の処理が実行中です」で失敗することがあった）
  const lock = (typeof 大邱_ロック取得_ === 'function')
    ? 大邱_ロック取得_(30000, cfg.label)
    : LockService.getDocumentLock();
  if (!lock) return;

  try {
    const n = lastRow - cfg.startRow + 1;
    const checkRange = sh.getRange(cfg.startRow, cfg.checkCol, n, 1);
    const checks = checkRange.getValues();
    const datesRange = sh.getRange(cfg.startRow, cfg.dateCol, n, 1);
    const dates = datesRange.getValues();
    const arrivalRange = sh.getRange(cfg.startRow, cfg.arrivalQtyCol, n, 1);
    const arrivals = arrivalRange.getValues();
    const orders = sh.getRange(cfg.startRow, cfg.orderQtyCol, n, 1).getValues();
    const today = _todayOnly_();

    let checked = 0;
    let updated = 0;
    let skippedNoQty = 0;

    for (let i = 0; i < n; i++) {
      if (checks[i][0] !== true) continue;
      checked++;

      const orderQty = orders[i][0];
      const hasOrderQty = orderQty !== '' && orderQty !== null && typeof orderQty !== 'undefined';
      if (!hasOrderQty) {
        skippedNoQty++;
        continue;
      }

      arrivals[i][0] = orderQty;
      if (dates[i][0] === '' || dates[i][0] === null) {
        dates[i][0] = today;
      }
      // 発注リスト大邱データではチェックを残す(そのまま「チェック行をEMS大邱へ送る」で送れるように)。
      // 発注シートは従来どおり外す。
      if (sheetName !== DAEGU_CFG.HACHU_SRC) checks[i][0] = false;
      updated++;
    }

    if (!checked) {
      ss.toast(cfg.label + ': チェックされた行がありません。');
      return;
    }

    arrivalRange.setValues(arrivals);
    datesRange.setValues(dates);
    checkRange.setValues(checks);
    SpreadsheetApp.flush();
    const filledDates = _fillArrivalDatesForAllQtyRows_(sh, cfg.arrivalQtyCol, cfg.dateCol, cfg.startRow);

    if (sheetName === DAEGU_CFG.HACHU_SRC && typeof 大邱発注_チェックと残り数量を設置 === 'function') {
      大邱発注_チェックと残り数量を設置();
      if (typeof 大邱未作業_同期予約_ === 'function') 大邱未作業_同期予約_(); // 未作業リストへ反映予約
    } else if (sheetName === CFG.HACHU_SHEET && typeof colorKeshikomiAllRows_ === 'function') {
      colorKeshikomiAllRows_(sh);
    }

    const checkNote = (sheetName === DAEGU_CFG.HACHU_SRC) ? '（チェックは残っています。続けて送信できます）' : '';
    ss.toast(cfg.label + ': 入荷数を発注数量にしました ' + updated + '行 / 入荷日補完 ' + filledDates + '行 / 数量なしスキップ ' + skippedNoQty + '行' + checkNote);
  } finally {
    lock.releaseLock();
  }
}

// 選択変更（シート切り替え含む）: 大邱未作業データを開いたら必要なときだけ自動同期
function onSelectionChange(e) {
  try {
    if (!e || !e.range) return;
    const sheetName = e.range.getSheet().getName();
    if (typeof MISAGYO_CFG !== 'undefined' &&
        sheetName === MISAGYO_CFG.DST_SHEET &&
        typeof 大邱未作業_onSelectionChange_ === 'function') {
      大邱未作業_onSelectionChange_(e);
    }
  } catch (err) {
    // onSelectionChangeは頻繁に発火するためエラーはログのみ（UIは出さない）
    console.error('onSelectionChange: ' + err);
  }
}

function autoRefreshShippingCountsOnChange(e) {
  const changeType = String((e && e.changeType) || '');
  if (changeType === 'FORMAT') {
    // 色塗り（背景色の変更）は未作業リストの判定に影響するので同期だけ予約して終了
    // ※ 大邱未作業データ自身の書式変更（＝リスト再構築の書き込み）では予約しない（自己ループ防止）
    const fmtSheet = SpreadsheetApp.getActive().getActiveSheet();
    const fmtName = fmtSheet ? fmtSheet.getName() : '';
    if (typeof MISAGYO_CFG !== 'undefined' && fmtName === MISAGYO_CFG.DST_SHEET) return;
    if (typeof 大邱未作業_同期予約_ === 'function') 大邱未作業_同期予約_();
    return;
  }
  if (changeType === 'EDIT') return;

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  const sheetName = sh ? sh.getName() : '';
  const daeguSheets = [DAEGU_CFG.HACHU_SRC, DAEGU_CFG.EMS_SRC];   // 大邱側
  const hatchuSheets = [DAEGU_CFG.HACHU_DST, DAEGU_CFG.EMS_DST];  // 発注/EMSリスト側
  const isDaegu = daeguSheets.indexOf(sheetName) >= 0;
  const isHatchu = hatchuSheets.indexOf(sheetName) >= 0;
  if (sheetName && !isDaegu && !isHatchu) return;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(1000)) return; // 自動再計算は混んでいたらスキップ（次の変更でまた走る）
  try {
    // 変更があった側だけ再計算する（従来は毎回両方やって10秒超かかっていた）
    if ((isDaegu || !sheetName) && typeof 大邱発注_チェックと残り数量を設置 === 'function') {
      大邱発注_チェックと残り数量を設置();
      if (typeof 大邱未作業_同期予約_ === 'function') 大邱未作業_同期予約_(); // 行追加・転送後も未作業リストへ反映
    }
    if ((isHatchu || !sheetName) && typeof 発注_EMS発送数数式を一括修正 === 'function') {
      発注_EMS発送数数式を一括修正();
    }
  } finally {
    lock.releaseLock();
  }
}

function 発送数自動再計算トリガーを設置() {
  const ss = SpreadsheetApp.getActive();
  const handler = 'autoRefreshShippingCountsOnChange';
  let deleted = 0;
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === handler) {
      ScriptApp.deleteTrigger(trigger);
      deleted++;
    }
  });
  ScriptApp.newTrigger(handler).forSpreadsheet(ss).onChange().create();
  ss.toast('発送数の自動再計算トリガーを設置しました。旧トリガー削除: ' + deleted + '件');
}

function 発注_購入No情報_(value) {
  if (typeof 大邱発注_発注No情報_ === 'function') {
    return 大邱発注_発注No情報_(value);
  }
  const raw = String(value || '')
    .trim()
    .replace(/[＿]/g, '_')
    .replace(/[\s\u3000]+/g, '')
    .replace(/_+/g, '_');
  if (!raw || /発注NO|OrderNo|DorderDate/i.test(raw)) return null;

  const m = raw.match(/^(.+\d{6,8})_([^_]+)(?:_(\d+)(?:_\d+)*)?$/);
  if (m) {
    const seq = Number(m[3]) || 0;
    const base = m[1] + '_' + m[2];
    return { base: base, seq: seq, numbered: seq > 0, normalized: seq > 0 ? base + '_' + seq : base };
  }
  return { base: raw.replace(/_+$/, ''), seq: 0, numbered: false, normalized: raw.replace(/_+$/, '') };
}

function 発注_連番管理_() {
  if (typeof 大邱発注_連番管理_ === 'function') {
    return 大邱発注_連番管理_();
  }
  const used = {};
  const next = {};
  const has = (base, seq) => !!(used[base] && used[base][seq]);
  const reserve = (base, seq) => {
    if (!base || !seq) return;
    if (!used[base]) used[base] = {};
    used[base][seq] = true;
    if (!next[base] || next[base] <= seq) next[base] = seq + 1;
  };
  const takeNext = base => {
    if (!next[base]) next[base] = 1;
    while (has(base, next[base])) next[base]++;
    const seq = next[base];
    reserve(base, seq);
    return seq;
  };
  return {
    reserve: reserve,
    assign: info => {
      if (info.numbered && !has(info.base, info.seq)) {
        reserve(info.base, info.seq);
        return info.seq;
      }
      return takeNext(info.base);
    }
  };
}

function 発注_onEdit購入No採番_(e) {
  if (!e || !e.range) return;

  const col = HATCHU_GROUP_COL; // G列 購入No
  const sc = e.range.getColumn();
  const ec = e.range.getLastColumn();
  if (sc > col || ec < col) return;

  const sh = e.range.getSheet();
  const editStart = Math.max(HATCHU_CFG.START_ROW, e.range.getRow());
  const editEnd = Math.max(editStart - 1, e.range.getLastRow());
  if (editEnd < HATCHU_CFG.START_ROW) return;

  const numRows = editEnd - editStart + 1;
  if (numRows <= 0) return;

  const cells = sh.getRange(editStart, col, numRows, 1);
  const vals = cells.getValues();
  const tracker = 発注_連番管理_();
  const lastRow = sh.getLastRow();
  if (lastRow < HATCHU_CFG.START_ROW) return;

  const allVals = sh.getRange(HATCHU_CFG.START_ROW, col, lastRow - HATCHU_CFG.START_ROW + 1, 1).getValues();
  for (let i = 0; i < allVals.length; i++) {
    const rowNo = HATCHU_CFG.START_ROW + i;
    if (editStart <= rowNo && rowNo <= editEnd) continue;

    const info = 発注_購入No情報_(allVals[i][0]);
    if (!info) continue;
    if (info.numbered) {
      tracker.reserve(info.base, info.seq);
    } else {
      tracker.reserve(info.base, 1);
    }
  }

  let changed = false;
  for (let i = 0; i < vals.length; i++) {
    const info = 発注_購入No情報_(vals[i][0]);
    if (!info) continue;

    const seq = tracker.assign(info);
    const fixed = info.base + '_' + seq;
    if (String(vals[i][0] || '').trim() !== fixed) {
      vals[i][0] = fixed;
      changed = true;
    }
  }

  if (changed) cells.setValues(vals);
}
// =====================================================
// 発注シート：チェックボックスと罫線の一括高速処理ロジック
// =====================================================
function HATCHU_processBordersAndCheckboxes_(sh, startRow, numRows) {
  const cfg = HATCHU_CFG;

  const maxCol = Number(cfg.MAX_COL || cfg.BORDER_COLS || 28);
  const colNo = Number(cfg.COL_NO || 2);
  const colChk = Number(cfg.COL_CHK || 1);

  startRow = Number(startRow);
  numRows = Number(numRows);

  if (!sh) return;
  if (!startRow || !numRows || numRows <= 0) return;

  const maxRows = sh.getMaxRows();

  // 編集行の前後も広めに見る
  startRow = Math.max(cfg.START_ROW, startRow - 2);
  const endRow = Math.min(maxRows, startRow + numRows + 4);
  numRows = endRow - startRow + 1;

  if (numRows <= 0) return;

  const noVals = sh
    .getRange(startRow, colNo, numRows, 1)
    .getDisplayValues();

  const checkboxRule = SpreadsheetApp
    .newDataValidation()
    .requireCheckbox()
    .build();

  const fullRange = sh.getRange(startRow, 1, numRows, maxCol);

  // まず対象範囲の罫線を全部消す
  fullRange.setBorder(
    false,
    false,
    false,
    false,
    false,
    false
  );

  // A列チェックボックスも対象範囲で一度整理
  const chkRangeAll = sh.getRange(startRow, colChk, numRows, 1);
  chkRangeAll.clearDataValidations();

  // No.がない行のチェックボックス内容だけ消す
  for (let i = 0; i < numRows; i++) {
    const rowNo = startRow + i;
    const hasNo = String(noVals[i][0] || '').trim() !== '';

    const chkCell = sh.getRange(rowNo, colChk, 1, 1);

    if (hasNo) {
      chkCell.setDataValidation(checkboxRule);
    } else {
      chkCell.clearContent();
    }
  }

  // ここが重要：
  // No.がある行を「1行ずつ」格子罫線にする
  // これで最終データ行の下罫線も確実に入る
  for (let i = 0; i < numRows; i++) {
    const rowNo = startRow + i;
    const hasNo = String(noVals[i][0] || '').trim() !== '';

    if (!hasNo) continue;

    sh.getRange(rowNo, 1, 1, maxCol)
      .setBorder(
        true,   // 上
        true,   // 左
        true,   // 下
        true,   // 右
        true,   // 縦線
        true,   // 横線
        '#000000',
        SpreadsheetApp.BorderStyle.SOLID
      );
  }
}

/*** 設定:自分の列に合わせる(A=1,B=2,…)。商品マスタの列はうろ覚えやから要確認 ***/
/*** 設定:自分の列に合わせる(A=1,B=2,…) ***/
const CFG = {
  HACHU_SHEET: '発注',
  HACHU_HEADER_ROW: 6,

  // 発注シート側
  HACHU_CODE: 12,    // L列 商品コード
  HACHU_VENDOR: 9,   // I列 業者
  HACHU_NAME: 10,    // J列 商品名
  HACHU_ITEM: 15,    // O列 品目
  HACHU_WEIGHT: 16,  // P列 重さ
  HACHU_PRICE: 17,   // Q列 価格

  // 商品マスタ側
  MASTER_SHEET: '商品マスタ',
  MASTER_HEADER_ROW: 6,

  M_NO: 0,
  M_CODE: 2,     // B列 商品コード
  M_VENDOR: 3,   // C列 業者
  M_NAME: 4,     // D列 商品名
  M_ITEM: 5,     // E列 品目
  M_PRICE: 6,    // F列 価格
  M_WEIGHT: 7,   // G列 重さ

  UPDATE_PRICE_IF_DIFF: true,
  SEP: '│'
};

const DAEGU_HACHU_MASTER_CFG = {
  HACHU_SHEET: '発注リスト大邱データ',
  HACHU_HEADER_ROW: 5,
  HACHU_CODE: 11,    // K列 商品コード
  HACHU_VENDOR: 8,   // H列 業者
  HACHU_NAME: 9,     // I列 商品名
  HACHU_ITEM: 14,    // N列 品目
  HACHU_WEIGHT: 15,  // O列 重さ
  HACHU_PRICE: 16,   // P列 価格
  // コード入力時にマスタから自動補完するのは重さだけ。
  // 商品名/品目/価格/業者は毎回発注リスト側に手入力し、その最新値がマスタへ流れる。
  AUTOFILL_WEIGHT_ONLY: true
};

/************************************************************
 * 商品マスタ 既存データ補完設定
 * 今は「重さ」だけ。
 * 今後列が増えたら MASTER_SYNC_FIELD_RULES に1つ追加するだけでOK。
 ************************************************************/
const MASTER_SYNC_FIELD_RULES = [
  {
    label: '重さ',
    sourceCfgKey: 'HACHU_WEIGHT',
    sourceHeaders: ['重さ', 'weight(g)', 'weight', '重量'],
    masterCfgKey: 'M_WEIGHT',
    masterHeaders: ['重さ', 'weight(g)', 'weight', '重量'],
    overwrite: true // true=既存値も違えば更新 / false=空欄だけ補完
  }

  // 例：今後「サイズ」列も同期したくなったらこう追加
  // ,
  // {
  //   label: 'サイズ',
  //   sourceCfgKey: 'HACHU_SIZE',
  //   sourceHeaders: ['サイズ', 'Size'],
  //   masterCfgKey: 'M_SIZE',
  //   masterHeaders: ['サイズ', 'Size'],
  //   overwrite: true
  // }
];

/************************************************************
 * 商品マスタで管理しない共通コード
 * ふろく等、1つのコードを複数の別商品で使い回すもの。
 * ・マスタ更新（一括／自動登録）の対象外＝マスタに行を作らない・更新しない
 * ・コード入力時の自動補完もしない＝業者/商品名/品目/価格/重さは行ごとに手入力
 * リストは「マスタ除外コード」シートのA列(2行目以降)で管理する。
 * ここの配列は初期値（シートが無いときの保険＋シート作成時の種）。
 * どちらに書いても効く（両方を合わせて判定・表記ゆれは正規化で吸収）。
 ************************************************************/
const MASTER_EXCLUDE_CODES = ['Promotional Item'];
const MASTER_EXCLUDE_SHEET = 'マスタ除外コード';

let _masterExcludeKeysCache_ = null; // 1回の実行内でシートを読むのは1度だけ

function _masterExcludeKeySet_() {
  if (_masterExcludeKeysCache_) return _masterExcludeKeysCache_;
  const keys = new Set();
  MASTER_EXCLUDE_CODES.forEach(c => { const k = _masterPrimaryCodeKey_(c); if (k) keys.add(k); });
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(MASTER_EXCLUDE_SHEET);
    if (sh) {
      const last = sh.getLastRow();
      if (last >= 2) {
        sh.getRange(2, 1, last - 1, 1).getValues().forEach(r => {
          const k = _masterPrimaryCodeKey_(r[0]);
          if (k) keys.add(k);
        });
      }
    }
  } catch (e) {} // シートが読めなくても配列側の既定で動く
  _masterExcludeKeysCache_ = keys;
  return keys;
}

// リクエスト注文などの「ラベル」判定：商品コードに半角英数字が1文字も無いもの。
// 例) 森 富士子 / 小杉山 由乃 / 自家用 / ★コピペ → true（マスタ対象外）
// 本物の商品コードは必ず英数字を含む（KRSJCM03、TUMBL48-5-페네로페 等）ので誤判定しない。
function _masterIsLabelCode_(code) {
  const raw = String(code || '').normalize('NFKC').trim();
  if (!raw) return false;
  return !/[0-9A-Za-z]/.test(raw);
}

function _masterIsExcludedCode_(code) {
  if (_masterIsLabelCode_(code)) return true; // 人名・自家用・★コピペ等のラベルは商品マスタで管理しない
  const key = _masterPrimaryCodeKey_(code);
  if (!key) return false;
  return _masterExcludeKeySet_().has(key);
}

// メニュー用: 「マスタ除外コード」シートを作成し、配列側の既定コードを追記する
function 商品マスタ_除外コードシートを作成() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(MASTER_EXCLUDE_SHEET);
  if (!sh) {
    sh = ss.insertSheet(MASTER_EXCLUDE_SHEET);
    sh.getRange(1, 1).setValue('商品マスタ除外コード').setFontWeight('bold').setFontSize(12);
    sh.getRange(1, 2).setValue('← ふろく等の共通コード。A2以降に1行1コードで追加すると、マスタ登録・自動補完の対象外になる')
      .setFontColor('#888888');
    sh.setColumnWidth(1, 220);
    sh.setColumnWidth(2, 600);
    sh.setFrozenRows(1);
  }
  const last = sh.getLastRow();
  const existing = new Set();
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 1).getValues().forEach(r => {
      const k = _masterPrimaryCodeKey_(r[0]);
      if (k) existing.add(k);
    });
  }
  const toAdd = MASTER_EXCLUDE_CODES.filter(c => {
    const k = _masterPrimaryCodeKey_(c);
    return k && !existing.has(k);
  });
  if (toAdd.length) sh.getRange(sh.getLastRow() + 1, 1, toAdd.length, 1).setValues(toAdd.map(c => [c]));
  _masterExcludeKeysCache_ = null; // 次の判定でシートを読み直す
  ss.toast('マスタ除外コードシートを更新しました（追加 ' + toAdd.length + '件）', '🚫 マスタ除外コード', 5);
}


/**
 * 発注シートの既存データから商品マスタを一括補完・更新する
 * 現在は発注リスト大邱データを正として、商品コードだけで商品マスタを更新する
 */
function 商品マスタ_既存データを発注から補完更新(silent) {
  return 商品マスタ_発注大邱から更新(silent);

  const ss = SpreadsheetApp.getActive();
  const cfg = (typeof CFG !== 'undefined') ? CFG : {};

  const hachu = ss.getSheetByName(cfg.HACHU_SHEET || '発注');
  const master = ss.getSheetByName(cfg.MASTER_SHEET || '商品マスタ');
  // トリガー実行時はUIが使えないのでログに切り替える
  const say = msg => { if (silent) Logger.log(msg); else SpreadsheetApp.getUi().alert(msg); };

  if (!hachu) {
    say('発注シートが見つからんで');
    return;
  }
  if (!master) {
    say('商品マスタシートが見つからんで');
    return;
  }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) {
    say('今ほかの処理が動いてるみたい。少し待ってもう一回。');
    return;
  }

  try {
    const hHeaderRow = cfg.HACHU_HEADER_ROW || 6;
    const mHeaderRow = cfg.MASTER_HEADER_ROW || 1;

    const hBase = hHeaderRow + 1;
    const mBase = mHeaderRow + 1;

    const hLastRow = hachu.getLastRow();
    const mLastRow = master.getLastRow();

    if (hLastRow < hBase) {
      say('発注データが無いで');
      return;
    }
    if (mLastRow < mBase) {
      say('商品マスタにデータが無いで');
      return;
    }

    const hHeaderRows = _補完_headerRows_(hHeaderRow);
    const mHeaderRows = _補完_headerRows_(mHeaderRow);

    const hCodeCol = _補完_resolveCol_(hachu, 'HACHU_CODE', ['商品コード', 'Code', 'code'], hHeaderRows);
    const hVendorCol = _補完_resolveCol_(hachu, 'HACHU_VENDOR', ['業者', 'Vendor', 'vendor'], hHeaderRows);

    const mCodeCol = _補完_resolveCol_(master, 'M_CODE', ['商品コード', 'Code', 'code'], mHeaderRows);
    const mVendorCol = _補完_resolveCol_(master, 'M_VENDOR', ['業者', 'Vendor', 'vendor'], mHeaderRows);

    if (!hCodeCol || !hVendorCol || !mCodeCol || !mVendorCol) {
      say(
        '商品コード列または業者列が見つからんで。\n\n' +
        'CFGの HACHU_CODE / HACHU_VENDOR / M_CODE / M_VENDOR を確認してな。'
      );
      return;
    }

    const rules = MASTER_SYNC_FIELD_RULES.map(rule => {
      return {
        label: rule.label,
        sourceCol: _補完_resolveCol_(hachu, rule.sourceCfgKey, rule.sourceHeaders, hHeaderRows),
        masterCol: _補完_resolveCol_(master, rule.masterCfgKey, rule.masterHeaders, mHeaderRows),
        overwrite: rule.overwrite !== false
      };
    });

    const usableRules = rules.filter(r => r.sourceCol && r.masterCol);
    const missingRules = rules.filter(r => !r.sourceCol || !r.masterCol);

    if (usableRules.length === 0) {
      say('補完対象の列が見つからんで。重さ列の見出し、またはCFGの M_WEIGHT / HACHU_WEIGHT を確認してな。');
      return;
    }

    const sep = cfg.SEP || '___SEP___';

    const hLastCol = hachu.getLastColumn();
    const hData = hachu.getRange(hBase, 1, hLastRow - hHeaderRow, hLastCol).getValues();

    // 発注データ側：商品コード＋業者ごとに、補完したい値を集める
    // 同じ商品コード＋業者が複数ある場合は、下の行の非空値を優先
    const sourceMap = new Map();

    for (let i = 0; i < hData.length; i++) {
      const row = hData[i];

      const code = String(row[hCodeCol - 1] || '').trim();
      const vendor = String(row[hVendorCol - 1] || '').trim();

      if (!code || !vendor) continue;

      const key = code + sep + vendor;
      const rec = sourceMap.get(key) || {};

      usableRules.forEach(rule => {
        const value = row[rule.sourceCol - 1];
        if (_補完_notBlank_(value)) {
          rec[rule.label] = value;
        }
      });

      sourceMap.set(key, rec);
    }

    const mRows = mLastRow - mHeaderRow;
    const maxMasterCol = Math.max(
      master.getLastColumn(),
      mCodeCol,
      mVendorCol,
      ...usableRules.map(r => r.masterCol)
    );

    const mData = master.getRange(mBase, 1, mRows, maxMasterCol).getValues();

    // 更新列ごとに配列を用意して、最後にまとめて書き戻す
    const colValuesMap = {};
    usableRules.forEach(rule => {
      colValuesMap[rule.masterCol] = master.getRange(mBase, rule.masterCol, mRows, 1).getValues();
    });

    let matched = 0;
    let updated = 0;
    let skipped = 0;
    let noSource = 0;

    const updateCountByLabel = {};

    for (let i = 0; i < mData.length; i++) {
      const mRow = mData[i];

      const code = String(mRow[mCodeCol - 1] || '').trim();
      const vendor = String(mRow[mVendorCol - 1] || '').trim();

      if (!code || !vendor) {
        skipped++;
        continue;
      }

      const key = code + sep + vendor;
      const src = sourceMap.get(key);

      if (!src) {
        noSource++;
        continue;
      }

      matched++;

      usableRules.forEach(rule => {
        const newValue = src[rule.label];

        if (!_補完_notBlank_(newValue)) return;

        const oldValue = mRow[rule.masterCol - 1];

        // overwrite=false の場合、マスタに既に値があれば触らない
        if (!rule.overwrite && _補完_notBlank_(oldValue)) {
          skipped++;
          return;
        }

        if (_補完_normCell_(oldValue) !== _補完_normCell_(newValue)) {
          colValuesMap[rule.masterCol][i][0] = newValue;
          mData[i][rule.masterCol - 1] = newValue;

          updated++;
          updateCountByLabel[rule.label] = (updateCountByLabel[rule.label] || 0) + 1;
        } else {
          skipped++;
        }
      });
    }

    // 一括書き戻し
    Object.keys(colValuesMap).forEach(col => {
      master.getRange(mBase, Number(col), mRows, 1).setValues(colValuesMap[col]);
    });

    let msg =
      '商品マスタの既存データ補完が完了したで\n\n' +
      `発注側キー数：${sourceMap.size}\n` +
      `マスタ一致行：${matched}\n` +
      `更新：${updated}\n` +
      `変更なし・スキップ：${skipped}\n` +
      `発注側に元データなし：${noSource}`;

    const labels = Object.keys(updateCountByLabel);
    if (labels.length) {
      msg += '\n\n【更新内訳】';
      labels.forEach(label => {
        msg += `\n${label}：${updateCountByLabel[label]}`;
      });
    }

    if (missingRules.length) {
      msg += '\n\n⚠ 見つからなかった補完ルール';
      missingRules.forEach(r => {
        msg += `\n${r.label}：発注列 or マスタ列が見つからんかった`;
      });
    }

    say(msg);

  } finally {
    lock.releaseLock();
  }
}


/************************************************************
 * 補助関数
 ************************************************************/

function _補完_headerRows_(mainHeaderRow) {
  const rows = [];
  if (mainHeaderRow - 1 >= 1) rows.push(mainHeaderRow - 1);
  rows.push(mainHeaderRow);
  return rows;
}

function _補完_resolveCol_(sheet, cfgKey, headerNames, headerRows) {
  const cfg = (typeof CFG !== 'undefined') ? CFG : {};

  if (cfgKey && cfg[cfgKey]) {
    return cfg[cfgKey];
  }

  return _補完_findColByHeaders_(sheet, headerNames, headerRows);
}

function _補完_findColByHeaders_(sheet, headerNames, headerRows) {
  const lastCol = sheet.getLastColumn();
  const targets = headerNames.map(_補完_normHeader_);

  for (let r = 0; r < headerRows.length; r++) {
    const rowNo = headerRows[r];
    if (rowNo < 1) continue;

    const headers = sheet.getRange(rowNo, 1, 1, lastCol).getDisplayValues()[0];

    for (let c = 0; c < headers.length; c++) {
      const h = _補完_normHeader_(headers[c]);
      if (targets.includes(h)) {
        return c + 1;
      }
    }
  }

  return 0;
}

function _補完_normHeader_(v) {
  return String(v || '')
    .trim()
    .toLowerCase()
    .replace(/\s/g, '')
    .replace(/[（）]/g, '')
    .replace(/[()]/g, '');
}

function _補完_notBlank_(v) {
  return v !== '' && v !== null && typeof v !== 'undefined';
}

function _補完_normCell_(v) {
  if (v instanceof Date) {
    return v.getTime();
  }
  return String(v ?? '').trim();
}


function _masterMap(){
  const sh=SpreadsheetApp.getActive().getSheetByName(CFG.MASTER_SHEET), last=sh.getLastRow();
  const info={}, byKey={}, byKeyW={}, byCode={}, bestByCode={};
  if(last<=CFG.MASTER_HEADER_ROW) return {info,byKey,byKeyW,byCode,bestByCode};
  const n=last-CFG.MASTER_HEADER_ROW, base=CFG.MASTER_HEADER_ROW+1;
  const v=sh.getRange(base,1,n,sh.getLastColumn()).getValues();
  for(let i=0;i<v.length;i++){
    const r=v[i];
    const rawCode=String(r[CFG.M_CODE-1]||'').trim(); if(!rawCode) continue;
    const rec={
      code: rawCode,
      vendor: String(r[CFG.M_VENDOR-1]||'').trim(),
      name: r[CFG.M_NAME-1],
      item: r[CFG.M_ITEM-1],
      price: r[CFG.M_PRICE-1],
      weight: r[CFG.M_WEIGHT-1],
      row: base+i
    };
    const keys=_masterCodeKeys_(rawCode);
    keys.forEach(c=>{
      bestByCode[c]=_masterMergeLatestRecord_(bestByCode[c], rec);
      if(_masterHasValue_(rec.price) || !(c in byKey)) byKey[c]=rec.price;
      if(_masterHasValue_(rec.weight) || !(c in byKeyW)) byKeyW[c]=rec.weight;
    });
  }

  Object.keys(bestByCode).forEach(c=>{
    const rec=bestByCode[c];
    info[c]={name: rec.name || '', item: rec.item || ''};
    byCode[c]=[rec];
  });
  return {info,byKey,byKeyW,byCode,bestByCode};
}

function _masterCodeKeys_(code) {
  const keys = [];
  const raw = String(code || '').trim();
  if (raw) keys.push(raw);
  if (typeof normCode_ === 'function') {
    const norm = normCode_(raw);
    if (norm && keys.indexOf(norm) < 0) keys.push(norm);
  }
  return keys;
}

function _masterPrimaryCodeKey_(code) {
  const keys = _masterCodeKeys_(code);
  return keys.length ? keys[keys.length - 1] : '';
}

function _masterHasValue_(value) {
  return value !== '' && value !== null && typeof value !== 'undefined' && String(value).trim() !== '';
}

function _masterMergeLatestRecord_(oldRec, newRec) {
  const rec = oldRec ? Object.assign({}, oldRec) : { code: newRec.code };
  ['vendor', 'name', 'item', 'price', 'weight'].forEach(field => {
    if (_masterHasValue_(newRec[field])) rec[field] = newRec[field];
  });
  rec.row = newRec.row;
  return rec;
}

function _masterGetByCode_(map, code) {
  const primary = _masterPrimaryCodeKey_(code);
  if (primary && map[primary] !== undefined) return map[primary];
  const keys = _masterCodeKeys_(code);
  for (let i=0;i<keys.length;i++) {
    if (map[keys[i]] !== undefined) return map[keys[i]];
  }
  return undefined;
}

function _masterGetVendorValue_(map, code, vendor) {
  const primary = _masterPrimaryCodeKey_(code);
  if (primary && map[primary] !== undefined) return map[primary];
  const keys = _masterCodeKeys_(code);
  for (let i=0;i<keys.length;i++) {
    if (map[keys[i]] !== undefined) return map[keys[i]];
    const key = keys[i] + CFG.SEP + String(vendor || '').trim();
    if (map[key] !== undefined) return map[key];
  }
  return undefined;
}

function 商品マスタ_発注大邱から更新(silent) {
  return 商品マスタ_発注ソースから更新_(DAEGU_HACHU_MASTER_CFG, '発注リスト大邱データ', silent);
}

function 商品マスタ_発注ソースから更新_(sourceCfg, sourceLabel, silent) {
  const ss = SpreadsheetApp.getActive();
  const source = ss.getSheetByName(sourceCfg.HACHU_SHEET);
  const master = ss.getSheetByName(CFG.MASTER_SHEET);
  const say = msg => { if (silent) Logger.log(msg); else SpreadsheetApp.getUi().alert(msg); };

  if (!source) {
    say(sourceLabel + 'シートが見つからんで');
    return;
  }
  if (!master) {
    say('商品マスタシートが見つからんで');
    return;
  }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) {
    say('今ほかの処理が動いてるみたい。少し待ってもう一回。');
    return;
  }

  try {
    SpreadsheetApp.flush();

    const sourceStart = sourceCfg.HACHU_HEADER_ROW + 1;
    const sourceLast = source.getLastRow();
    if (sourceLast < sourceStart) {
      say(sourceLabel + 'にデータが無いで');
      return;
    }

    const sourceWidth = Math.max(
      sourceCfg.HACHU_CODE,
      sourceCfg.HACHU_VENDOR,
      sourceCfg.HACHU_NAME,
      sourceCfg.HACHU_ITEM,
      sourceCfg.HACHU_PRICE,
      sourceCfg.HACHU_WEIGHT || 0
    );
    const sourceValues = source
      .getRange(sourceStart, 1, sourceLast - sourceStart + 1, sourceWidth)
      .getValues();

    const sourceByCode = {};
    for (let i = 0; i < sourceValues.length; i++) {
      const r = sourceValues[i];
      const code = String(r[sourceCfg.HACHU_CODE - 1] || '').trim();
      if (!code || /^(商品コード|code)$/i.test(code)) continue;
      if (_masterIsExcludedCode_(code)) continue; // 共通コード(Promotional Item等)はマスタで管理しない

      const rec = {
        code,
        vendor: String(r[sourceCfg.HACHU_VENDOR - 1] || '').trim(),
        name: r[sourceCfg.HACHU_NAME - 1],
        item: r[sourceCfg.HACHU_ITEM - 1],
        price: r[sourceCfg.HACHU_PRICE - 1],
        weight: sourceCfg.HACHU_WEIGHT ? r[sourceCfg.HACHU_WEIGHT - 1] : '',
        row: sourceStart + i
      };

      const hasData = ['vendor', 'name', 'item', 'price', 'weight']
        .some(field => _masterHasValue_(rec[field]));
      if (!hasData) continue;

      const key = _masterPrimaryCodeKey_(code);
      if (!key) continue;
      sourceByCode[key] = _masterMergeLatestRecord_(sourceByCode[key], rec);
      sourceByCode[key].code = code;
    }

    const sourceRecords = Object.keys(sourceByCode)
      .map(key => sourceByCode[key])
      .sort((a, b) => (a.row || 0) - (b.row || 0));

    if (sourceRecords.length === 0) {
      say(sourceLabel + 'に商品マスタへ反映できる商品コードが無いで');
      return;
    }

    const mStart = CFG.MASTER_HEADER_ROW + 1;
    const mLast = master.getLastRow();
    const mRows = Math.max(0, mLast - CFG.MASTER_HEADER_ROW);
    const masterByCode = {};
    let masterValues = [];

    if (mRows > 0) {
      masterValues = master.getRange(mStart, CFG.M_CODE, mRows, 6).getValues();
      for (let i = 0; i < masterValues.length; i++) {
        const code = String(masterValues[i][0] || '').trim();
        const key = _masterPrimaryCodeKey_(code);
        if (key) masterByCode[key] = { offset: i, row: mStart + i };
      }
    }

    const fields = [
      { name: 'code', index: 0 },
      { name: 'vendor', index: 1 },
      { name: 'name', index: 2 },
      { name: 'item', index: 3 },
      { name: 'price', index: 4 },
      { name: 'weight', index: 5 }
    ];

    const toAdd = [];
    let matched = 0;
    let updatedRows = 0;
    let updatedCells = 0;

    sourceRecords.forEach(rec => {
      const key = _masterPrimaryCodeKey_(rec.code);
      const existing = masterByCode[key];

      if (!existing) {
        toAdd.push([
          rec.code,
          rec.vendor || '',
          _masterHasValue_(rec.name) ? rec.name : '',
          _masterHasValue_(rec.item) ? rec.item : '',
          _masterHasValue_(rec.price) ? rec.price : '',
          _masterHasValue_(rec.weight) ? rec.weight : ''
        ]);
        masterByCode[key] = { offset: -1, row: -1 };
        return;
      }

      matched++;
      let rowChanged = false;
      const row = masterValues[existing.offset];

      fields.forEach(field => {
        const nextValue = rec[field.name];
        if (!_masterHasValue_(nextValue)) return;

        const currentText = String(row[field.index] ?? '').trim();
        const nextText = String(nextValue ?? '').trim();
        if (currentText === nextText) return;

        // 発注リスト大邱の最新値でマスタをどんどん更新する（発注リストが正）。
        // 共通コード(Promotional Item等)は sourceByCode 構築時点で除外済み。
        row[field.index] = nextValue;
        updatedCells++;
        rowChanged = true;
      });

      if (rowChanged) updatedRows++;
    });

    if (mRows > 0 && updatedCells > 0) {
      master.getRange(mStart, CFG.M_CODE, mRows, 6).setValues(masterValues);
    }

    if (toAdd.length > 0) {
      const writeRow = Math.max(master.getLastRow() + 1, mStart);
      master.getRange(writeRow, CFG.M_CODE, toAdd.length, 6).setValues(toAdd);
    }

    SpreadsheetApp.flush();

    const uniqueResult = (typeof 商品マスタ_商品コード一意化_実行_ === 'function')
      ? 商品マスタ_商品コード一意化_実行_(master)
      : { groups: 0, deleted: 0 };

    say(
      '商品マスタを' + sourceLabel + 'から更新したで\n\n' +
      `大邱側の商品コード：${sourceRecords.length}件\n` +
      `商品マスタ既存一致：${matched}件\n` +
      `追加：${toAdd.length}件\n` +
      `更新行：${updatedRows}件\n` +
      `更新セル：${updatedCells}件\n` +
      `商品コード統合：${uniqueResult.groups}件\n` +
      `重複削除：${uniqueResult.deleted}行\n\n` +
      '発注リスト大邱の最新値でマスタを更新しています（除外コードは対象外・空欄では消しません）。'
    );

    return {
      source: sourceRecords.length,
      matched,
      added: toAdd.length,
      updatedRows,
      updatedCells,
      uniqueGroups: uniqueResult.groups,
      deleted: uniqueResult.deleted
    };
  } finally {
    lock.releaseLock();
  }
}

function autofillHachu(e){
  autofillHachuByCfg_(e, CFG);
}

// ============================================================
// 発注リスト大邱データのonEdit用: 編集された行の最新値で商品マスタを即更新する
//   「発注リストの最新情報でどんどんマスタを新しくする」運用の自動化。
//   反映は 商品マスタ_候補行を商品コードで反映_（除外コード対応・最新値で上書き）。
//   混雑時はスキップ（次の編集かメニューの一括更新で追いつく）。
// ============================================================
function 大邱_マスタ自動更新_(sh, range) {
  const cfg = DAEGU_HACHU_MASTER_CFG;
  if (!sh || !range || sh.getName() !== cfg.HACHU_SHEET) return 0;
  const master = SpreadsheetApp.getActive().getSheetByName(CFG.MASTER_SHEET);
  if (!master) return 0;

  const firstRow = Math.max(cfg.HACHU_HEADER_ROW + 1, range.getRow());
  const lastRow = Math.min(sh.getLastRow(), range.getLastRow());
  if (lastRow < firstRow) return 0;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(1000)) return 0;
  try {
    const width = Math.max(cfg.HACHU_CODE, cfg.HACHU_VENDOR, cfg.HACHU_NAME,
      cfg.HACHU_ITEM, cfg.HACHU_PRICE, cfg.HACHU_WEIGHT || 0);
    const rows = sh.getRange(firstRow, 1, lastRow - firstRow + 1, width).getValues();
    const candidates = [];
    for (const r of rows) {
      const code = String(r[cfg.HACHU_CODE - 1] || '').trim();
      const name = String(r[cfg.HACHU_NAME - 1] || '').trim();
      if (!code || !name) continue; // コードと商品名が入ってから登録（価格は必須にしない=定期購読も対象）
      candidates.push([
        code,
        r[cfg.HACHU_VENDOR - 1],
        r[cfg.HACHU_NAME - 1],
        r[cfg.HACHU_ITEM - 1],
        r[cfg.HACHU_PRICE - 1],
        cfg.HACHU_WEIGHT ? r[cfg.HACHU_WEIGHT - 1] : ''
      ]);
    }
    if (!candidates.length) return 0;
    const result = 商品マスタ_候補行を商品コードで反映_(master, candidates);
    return result.added + result.updatedRows;
  } finally {
    lock.releaseLock();
  }
}

function autofillHachuByCfg_(e, sheetCfg){
  if(!e || !e.range || !sheetCfg) return;
  const sh=e.range.getSheet(); if(sh.getName()!==sheetCfg.HACHU_SHEET) return;
  const startCol=e.range.getColumn(), endCol=e.range.getLastColumn();
  if(!_rangeHitsAnyCol_(startCol,endCol,[sheetCfg.HACHU_CODE,sheetCfg.HACHU_VENDOR])) return;
  const firstRow=Math.max(sheetCfg.HACHU_HEADER_ROW+1,e.range.getRow());
  const lastRow=e.range.getLastRow();
  if(lastRow<firstRow) return;
  const master=_masterMap();
  for(let row=firstRow;row<=lastRow;row++){
    const code=String(sh.getRange(row,sheetCfg.HACHU_CODE).getValue()).trim(); if(!code) continue;
    if(_masterIsExcludedCode_(code)) continue; // 共通コードは自動補完しない(全項目を行ごとに手入力)
    const rec=_masterGetByCode_(master.bestByCode, code);
    if(!rec) continue;

    // AUTOFILL_WEIGHT_ONLY: マスタから返すのは重さだけ（発注リスト大邱用）。
    // 商品名/品目/価格/業者は発注リスト側の入力が正で、マスタ更新で吸い上げる。
    if(!sheetCfg.AUTOFILL_WEIGHT_ONLY){
      let vendor=String(sh.getRange(row,sheetCfg.HACHU_VENDOR).getValue()).trim();
      if(_masterHasValue_(rec.name)) sh.getRange(row,sheetCfg.HACHU_NAME).setValue(rec.name);
      if(_masterHasValue_(rec.item)) sh.getRange(row,sheetCfg.HACHU_ITEM).setValue(rec.item);
      if(!vendor && _masterHasValue_(rec.vendor)){
        vendor=String(rec.vendor || '').trim();
        sh.getRange(row,sheetCfg.HACHU_VENDOR).setValue(vendor);
      }
      if(_masterHasValue_(rec.price)) sh.getRange(row,sheetCfg.HACHU_PRICE).setValue(rec.price);
    }
    if(_masterHasValue_(rec.weight)) sh.getRange(row,sheetCfg.HACHU_WEIGHT).setValue(rec.weight);
  }
}

function registerToMaster(){
  return 商品マスタ_発注大邱から更新(false);
}
function _updateMasterPrice(master,code,vendor,price){
  const last=master.getLastRow(), n=last-CFG.MASTER_HEADER_ROW; if(n<=0) return;
  const base=CFG.MASTER_HEADER_ROW+1;
  const cs=master.getRange(base,CFG.M_CODE,n,1).getValues();
  const targetKey = _masterPrimaryCodeKey_(code);
  for(let i=0;i<n;i++) if(_masterPrimaryCodeKey_(cs[i][0])===targetKey){
    master.getRange(base+i,CFG.M_PRICE).setValue(price); return;
  }
}
function _nextHachuNo(sh, colNo){
  const start=CFG.HACHU_HEADER_ROW+1, last=sh.getLastRow();
  if(last<start) return 1;
  const vals=sh.getRange(start, colNo, last-start+1, 1).getValues();
  let max=0;
  vals.forEach(r=>{ const v=Number(r[0]); if(!isNaN(v)&&v>max) max=v; });
  return max+1;
}

function _masterNextRow_(master) {
  const start = CFG.MASTER_HEADER_ROW + 1;
  const last = Math.max(master.getLastRow(), start);

  const numRows = last - start + 1;
  const codes = master.getRange(start, CFG.M_CODE, numRows, 1).getValues();

  for (let i = 0; i < codes.length; i++) {
    if (String(codes[i][0] || '').trim() === '') {
      return start + i;
    }
  }

  return last + 1;
}

// グループの基準列を選ぶ:3=発注日(C) / 7=購入No(G)
const HATCHU_GROUP_COL = 7;   // ← 購入Noで囲みたいなら 7 に変える

function 発注_カートグループキー_(value) {
  const info = 発注_購入No情報_(value);
  if (info) return info.base;
  return String(value || '')
    .trim()
    .replace(/[＿]/g, '_')
    .replace(/[\s\u3000]+/g, '')
    .replace(/_+/g, '_');
}

// 発注シート:基準列が同じ連続行を太枠で囲む(中の格子は残す)
function 発注_drawGroupBorders_(sh) {
  const startRow = HATCHU_CFG.START_ROW;          // 7
  const last = sh.getLastRow();
  if (last < startRow) return;
  const numRows = last - startRow + 1;
  const startCol = 1;
  const maxCol = Number(HATCHU_CFG.MAX_COL || HATCHU_CFG.BORDER_COLS || 28);

  const keys = sh.getRange(startRow, HATCHU_GROUP_COL, numRows, 1)
                 .getDisplayValues().map(r => 発注_カートグループキー_(r[0]));

  let gStart = null, cur = '';
  for (let i = 0; i <= numRows; i++) {
    const isEnd = i === numRows;
    const key = isEnd ? '' : keys[i];

    if (gStart === null) {
      if (!isEnd && key !== '') { gStart = i; cur = key; }
      continue;
    }
    if (isEnd || key !== cur) {
      // グループの外枠だけ太線(縦線・横線=null で中の格子は触らへん)
      sh.getRange(startRow + gStart, startCol, i - gStart, maxCol)
        .setBorder(true, true, true, true, null, null,
                   '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
      if (!isEnd && key !== '') { gStart = i; cur = key; }
      else { gStart = null; cur = ''; }
    }
  }
}

// ボタン用:格子を引き直してから太枠を付ける(まるごとリフレッシュ)
function 発注_グループ罫線() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== HATCHU_CFG.SHEET_NAME) {
    SpreadsheetApp.getUi().alert('発注シートで実行してな');
    return;
  }
  const numRows = sh.getLastRow() - HATCHU_CFG.START_ROW + 1;
  if (numRows > 0) HATCHU_processBordersAndCheckboxes_(sh, HATCHU_CFG.START_ROW, numRows);
  発注_drawGroupBorders_(sh);
  SpreadsheetApp.getActive().toast('グループの太枠を更新したで');
}

function 商品マスタ_候補行を商品コードで反映_(master, candidates) {
  const sourceByCode = {};

  candidates.forEach((row, idx) => {
    const code = String(row[0] || '').trim();
    const key = _masterPrimaryCodeKey_(code);
    if (!key) return;
    if (_masterIsExcludedCode_(code)) return; // 共通コード(Promotional Item等)はマスタで管理しない

    const rec = {
      code,
      vendor: row[1],
      name: row[2],
      item: row[3],
      price: row[4],
      weight: row[5],
      row: idx + 1
    };

    const hasData = ['vendor', 'name', 'item', 'price', 'weight']
      .some(field => _masterHasValue_(rec[field]));
    if (!hasData) return;

    sourceByCode[key] = _masterMergeLatestRecord_(sourceByCode[key], rec);
    sourceByCode[key].code = code;
  });

  const sourceRecords = Object.keys(sourceByCode)
    .map(key => sourceByCode[key])
    .sort((a, b) => (a.row || 0) - (b.row || 0));

  if (sourceRecords.length === 0) {
    return { source: 0, matched: 0, added: 0, updatedRows: 0, updatedCells: 0, uniqueGroups: 0, deleted: 0 };
  }

  const mStart = CFG.MASTER_HEADER_ROW + 1;
  const mRows = Math.max(0, master.getLastRow() - CFG.MASTER_HEADER_ROW);
  const masterByCode = {};
  let masterValues = [];

  if (mRows > 0) {
    masterValues = master.getRange(mStart, CFG.M_CODE, mRows, 6).getValues();
    for (let i = 0; i < masterValues.length; i++) {
      const key = _masterPrimaryCodeKey_(masterValues[i][0]);
      if (key) masterByCode[key] = { offset: i, row: mStart + i };
    }
  }

  const fields = [
    { name: 'code', index: 0 },
    { name: 'vendor', index: 1 },
    { name: 'name', index: 2 },
    { name: 'item', index: 3 },
    { name: 'price', index: 4 },
    { name: 'weight', index: 5 }
  ];

  const toAdd = [];
  let matched = 0;
  let updatedRows = 0;
  let updatedCells = 0;

  sourceRecords.forEach(rec => {
    const key = _masterPrimaryCodeKey_(rec.code);
    const existing = masterByCode[key];

    if (!existing) {
      toAdd.push([
        rec.code,
        _masterHasValue_(rec.vendor) ? rec.vendor : '',
        _masterHasValue_(rec.name) ? rec.name : '',
        _masterHasValue_(rec.item) ? rec.item : '',
        _masterHasValue_(rec.price) ? rec.price : '',
        _masterHasValue_(rec.weight) ? rec.weight : ''
      ]);
      masterByCode[key] = { offset: -1, row: -1 };
      return;
    }

    matched++;
    let rowChanged = false;
    const row = masterValues[existing.offset];

    fields.forEach(field => {
      const nextValue = rec[field.name];
      if (!_masterHasValue_(nextValue)) return;

      const currentText = String(row[field.index] ?? '').trim();
      const nextText = String(nextValue ?? '').trim();
      if (currentText === nextText) return;

      // 発注側の最新値でマスタをどんどん更新する（発注リストが正）
      row[field.index] = nextValue;
      updatedCells++;
      rowChanged = true;
    });

    if (rowChanged) updatedRows++;
  });

  if (mRows > 0 && updatedCells > 0) {
    master.getRange(mStart, CFG.M_CODE, mRows, 6).setValues(masterValues);
  }

  if (toAdd.length > 0) {
    const writeRow = _masterNextRow_(master);
    master.getRange(writeRow, CFG.M_CODE, toAdd.length, 6).setValues(toAdd);
  }

  SpreadsheetApp.flush();

  const uniqueResult = (typeof 商品マスタ_商品コード一意化_実行_ === 'function')
    ? 商品マスタ_商品コード一意化_実行_(master)
    : { groups: 0, deleted: 0 };

  return {
    source: sourceRecords.length,
    matched,
    added: toAdd.length,
    updatedRows,
    updatedCells,
    uniqueGroups: uniqueResult.groups,
    deleted: uniqueResult.deleted
  };
}

// onEditから呼ぶ用:監視列を触ったときだけ全体補完を起動
function 商品コードと業者と商品名と価格がそろったら自動でマスタ登録(e){
  const sh = e.range.getSheet();
  if (sh.getName() !== CFG.HACHU_SHEET) return;

  const sCol = e.range.getColumn(), eCol = e.range.getLastColumn();
  const watch = [CFG.HACHU_VENDOR, CFG.HACHU_NAME, CFG.HACHU_ITEM, CFG.HACHU_PRICE, CFG.HACHU_WEIGHT];
  const pastedFullRow = e.range.getNumColumns() > 1 &&
    [CFG.HACHU_CODE, CFG.HACHU_VENDOR, CFG.HACHU_NAME, CFG.HACHU_PRICE].some(w => sCol <= w && w <= eCol);

  if (!watch.some(w => sCol <= w && w <= eCol) && !pastedFullRow) return;

  マスタ自動補完_編集範囲_(e.range);
}

function マスタ自動補完_編集範囲_(range, silent) {
  if (!range) return 0;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(1000)) return 0;

  try {
    const ss = SpreadsheetApp.getActive();
    const hachu = range.getSheet();
    const master = ss.getSheetByName(CFG.MASTER_SHEET);

    if (!hachu || !master || hachu.getName() !== CFG.HACHU_SHEET) return 0;

    const hStart = CFG.HACHU_HEADER_ROW + 1;
    const startRow = Math.max(hStart, range.getRow());
    const endRow = Math.min(hachu.getLastRow(), range.getLastRow());
    if (endRow < startRow) return 0;

    const width = Math.max(
      CFG.HACHU_CODE,
      CFG.HACHU_VENDOR,
      CFG.HACHU_NAME,
      CFG.HACHU_ITEM,
      CFG.HACHU_PRICE,
      CFG.HACHU_WEIGHT || 0
    );

    const rows = hachu.getRange(startRow, 1, endRow - startRow + 1, width).getValues();
    const candidates = [];

    for (const r of rows) {
      const code = String(r[CFG.HACHU_CODE - 1] || '').trim();
      const vendor = String(r[CFG.HACHU_VENDOR - 1] || '').trim();
      const name = String(r[CFG.HACHU_NAME - 1] || '').trim();
      const item = String(r[CFG.HACHU_ITEM - 1] || '').trim();
      const price = r[CFG.HACHU_PRICE - 1];
      const weight = CFG.HACHU_WEIGHT ? r[CFG.HACHU_WEIGHT - 1] : '';

      if (!code || !name) continue;
      if (price === '' || price === null || typeof price === 'undefined') continue;

      candidates.push([code, vendor, name, item, price, weight]);
    }

    if (candidates.length === 0) return 0;

    const result = 商品マスタ_候補行を商品コードで反映_(master, candidates);
    const changed = result.added + result.updatedRows + result.uniqueGroups;

    if (changed === 0) return 0;

    if (!silent) {
      SpreadsheetApp.getActive().toast(
        `商品マスタ更新：追加 ${result.added} / 更新 ${result.updatedRows} / 統合 ${result.uniqueGroups}`
      );
    }

    return changed;
  } finally {
    lock.releaseLock();
  }
}

function マスタ自動補完_(silent) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) return 0;

  try {
    const ss = SpreadsheetApp.getActive();
    const hachu  = ss.getSheetByName(CFG.HACHU_SHEET);
    const master = ss.getSheetByName(CFG.MASTER_SHEET);

    if (!hachu || !master) return 0;

    const hStart = CFG.HACHU_HEADER_ROW + 1;
    const hLast  = hachu.getLastRow();
    if (hLast < hStart) return 0;

    SpreadsheetApp.flush();

    const hv = hachu
      .getRange(hStart, 1, hLast - hStart + 1, hachu.getLastColumn())
      .getValues();

    const candidates = []; // [code, vendor, name, item, price, weight]

    for (const r of hv) {
      const code   = String(r[CFG.HACHU_CODE   - 1] || '').trim();
      const vendor = String(r[CFG.HACHU_VENDOR - 1] || '').trim();
      const name   = String(r[CFG.HACHU_NAME   - 1] || '').trim();
      const item   = String(r[CFG.HACHU_ITEM   - 1] || '').trim();
      const price  = r[CFG.HACHU_PRICE  - 1];
      const weight = CFG.HACHU_WEIGHT ? r[CFG.HACHU_WEIGHT - 1] : '';

      if (!code || !name) continue;
      if (price === '' || price === null) continue;

      candidates.push([code, vendor, name, item, price, weight]);
    }

    if (candidates.length === 0) return 0;

    const result = 商品マスタ_候補行を商品コードで反映_(master, candidates);
    const changed = result.added + result.updatedRows + result.uniqueGroups;
    if (changed === 0) return 0;

    if (!silent) {
      SpreadsheetApp.getActive().toast(
        `商品マスタ更新：追加 ${result.added} / 更新 ${result.updatedRows} / 統合 ${result.uniqueGroups}`
      );
    }

    return changed;

  } finally {
    lock.releaseLock();
  }
}

function 商品マスタ_発注シートから今すぐ補完登録() {
  return 商品マスタ_発注大邱から更新(false);
}

function _masterSeen_(master) {
  const seen = new Set();

  const base = CFG.MASTER_HEADER_ROW + 1;
  const last = master.getLastRow();

  if (last < base) {
    return { seen };
  }

  const n = last - CFG.MASTER_HEADER_ROW;
  const vals = master.getRange(base, 1, n, master.getLastColumn()).getValues();

  for (let i = 0; i < vals.length; i++) {
    const code = String(vals[i][CFG.M_CODE - 1] || '').trim();
    const key = _masterPrimaryCodeKey_(code);

    if (!key) continue;
    seen.add(key);
  }

  return { seen };
}

function 商品マスタ_重さだけ発注から補完更新() {
  return 商品マスタ_発注大邱から更新(false);

  const ss = SpreadsheetApp.getActive();
  const hachu = ss.getSheetByName(CFG.HACHU_SHEET);
  const master = ss.getSheetByName(CFG.MASTER_SHEET);
  const ui = SpreadsheetApp.getUi();

  if (!hachu || !master) {
    ui.alert('発注シートか商品マスタが見つからんで');
    return;
  }

  const hStart = CFG.HACHU_HEADER_ROW + 1;
  const hLast = hachu.getLastRow();
  if (hLast < hStart) {
    ui.alert('発注データが無いで');
    return;
  }

  const hVals = hachu
    .getRange(hStart, 1, hLast - hStart + 1, hachu.getLastColumn())
    .getValues();

  // 発注側から「商品コード＋業者 → 重さ」を作る
  const weightMap = new Map();

  for (const r of hVals) {
    const code = String(r[CFG.HACHU_CODE - 1] || '').trim();
    const vendor = String(r[CFG.HACHU_VENDOR - 1] || '').trim();
    const weight = r[CFG.HACHU_WEIGHT - 1];

    if (!code || !vendor) continue;
    if (weight === '' || weight === null || typeof weight === 'undefined') continue;

    weightMap.set(code + CFG.SEP + vendor, weight);
  }

  const mStart = CFG.MASTER_HEADER_ROW + 1;
  const mLast = master.getLastRow();
  if (mLast < mStart) {
    ui.alert('商品マスタにデータが無いで');
    return;
  }

  const mRows = mLast - CFG.MASTER_HEADER_ROW;
  const mVals = master
    .getRange(mStart, 1, mRows, master.getLastColumn())
    .getValues();

  const weightValues = master
    .getRange(mStart, CFG.M_WEIGHT, mRows, 1)
    .getValues();

  let updated = 0;
  let matched = 0;

  for (let i = 0; i < mVals.length; i++) {
    const code = String(mVals[i][CFG.M_CODE - 1] || '').trim();
    const vendor = String(mVals[i][CFG.M_VENDOR - 1] || '').trim();

    if (!code || !vendor) continue;

    const key = code + CFG.SEP + vendor;
    if (!weightMap.has(key)) continue;

    matched++;

    const newWeight = weightMap.get(key);
    const oldWeight = weightValues[i][0];

    if (String(oldWeight || '').trim() !== String(newWeight || '').trim()) {
      weightValues[i][0] = newWeight;
      updated++;
    }
  }

  master.getRange(mStart, CFG.M_WEIGHT, mRows, 1).setValues(weightValues);

  ui.alert(
    `商品マスタの重さ補完完了\n` +
    `発注側 重さデータ：${weightMap.size}件\n` +
    `商品マスタ一致：${matched}件\n` +
    `重さ更新：${updated}件`
  );
}

function 商品マスタ_足りないデータだけ発注から補完(silent) {
  return 商品マスタ_発注大邱から更新(silent);

  const ss = SpreadsheetApp.getActive();
  const hachu = ss.getSheetByName(CFG.HACHU_SHEET);
  const master = ss.getSheetByName(CFG.MASTER_SHEET);
  // トリガー実行時はUIが使えないのでログに切り替える
  const say = msg => { if (silent) Logger.log(msg); else SpreadsheetApp.getUi().alert(msg); };

  if (!hachu || !master) {
    say('発注シートか商品マスタが見つからんで');
    return;
  }

  const hStart = CFG.HACHU_HEADER_ROW + 1;
  const hLast = hachu.getLastRow();
  if (hLast < hStart) {
    say('発注データが無いで');
    return;
  }

  const hVals = hachu
    .getRange(hStart, 1, hLast - hStart + 1, hachu.getLastColumn())
    .getValues();

  // 発注シート側の 商品コード＋業者 → 補完元データ
  const srcMap = new Map();

  for (const r of hVals) {
    const code = String(r[CFG.HACHU_CODE - 1] || '').trim();
    const vendor = String(r[CFG.HACHU_VENDOR - 1] || '').trim();

    if (!code || !vendor) continue;

    const key = code + CFG.SEP + vendor;

    const old = srcMap.get(key) || {};

    srcMap.set(key, {
      name: old.name || r[CFG.HACHU_NAME - 1],
      item: old.item || r[CFG.HACHU_ITEM - 1],
      price: old.price || r[CFG.HACHU_PRICE - 1],
      weight: old.weight || r[CFG.HACHU_WEIGHT - 1]
    });
  }

  const mStart = CFG.MASTER_HEADER_ROW + 1;
  const mLast = master.getLastRow();
  if (mLast < mStart) {
    say('商品マスタにデータが無いで');
    return;
  }

  const mRows = mLast - CFG.MASTER_HEADER_ROW;
  const mLastCol = master.getLastColumn();

  const mVals = master
    .getRange(mStart, 1, mRows, mLastCol)
    .getValues();

  let updated = 0;
  let matched = 0;

  for (let i = 0; i < mVals.length; i++) {
    const code = String(mVals[i][CFG.M_CODE - 1] || '').trim();
    const vendor = String(mVals[i][CFG.M_VENDOR - 1] || '').trim();

    if (!code || !vendor) continue;

    const key = code + CFG.SEP + vendor;
    const src = srcMap.get(key);

    if (!src) continue;

    matched++;

    // 商品名：空欄だけ補完
    if (isBlankForMasterFill_(mVals[i][CFG.M_NAME - 1]) && !isBlankForMasterFill_(src.name)) {
      master.getRange(mStart + i, CFG.M_NAME).setValue(src.name);
      updated++;
    }

    // 品目：空欄だけ補完
    if (isBlankForMasterFill_(mVals[i][CFG.M_ITEM - 1]) && !isBlankForMasterFill_(src.item)) {
      master.getRange(mStart + i, CFG.M_ITEM).setValue(src.item);
      updated++;
    }

    // 価格：空欄だけ補完
    if (isBlankForMasterFill_(mVals[i][CFG.M_PRICE - 1]) && !isBlankForMasterFill_(src.price)) {
      master.getRange(mStart + i, CFG.M_PRICE).setValue(src.price);
      updated++;
    }

    // 重さ：空欄だけ補完
    if (isBlankForMasterFill_(mVals[i][CFG.M_WEIGHT - 1]) && !isBlankForMasterFill_(src.weight)) {
      master.getRange(mStart + i, CFG.M_WEIGHT).setValue(src.weight);
      updated++;
    }
  }

  say(
    '商品マスタの足りないデータ補完が完了したで\n\n' +
    `発注側データ：${srcMap.size}件\n` +
    `商品マスタ一致：${matched}件\n` +
    `補完したセル：${updated}件`
  );
}

function isBlankForMasterFill_(v) {
  return v === null ||
         typeof v === 'undefined' ||
         String(v).trim() === '' ||
         String(v).trim() === '　';
}

function テスト_STD093をマスタで探す() {
  const m = _masterMap();
  const code = 'STD093';
  const msg =
    'STD093 マスタ照合結果\n\n' +
    'info(名前/品目)：' + JSON.stringify(m.info[code]) + '\n' +
    'byCode(業者/価格/重さ)：' + JSON.stringify(m.byCode[code]);
  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}
