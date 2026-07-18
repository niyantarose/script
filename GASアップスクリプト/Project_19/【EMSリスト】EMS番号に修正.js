/************************************************************
 * EMS自社用ファイル 自動処理 【列番号確定版】
 *
 * 実際のシートのアルファベットに完全に合わせました：
 * A列：No.（商品点数カウンタ）
 * C列：EMS発送日
 * D列：ステータス列（⇒）
 * E列：EMS到着日（発送日+3日で予定日を自動補完）
 * L列：EMS番号（12番目の列）
 * M列：Box No.（13番目の列）
 ************************************************************/

/************************************************************
 * 設定
 ************************************************************/
const EMS_CFG = {
  HEADER_ROW: 6,
  START_ROW: 7,

  COL_NO: 1,         // A列：No.
  COL_EMS_DATE: 3,   // C列：EMS発送日
  COL_STATUS: 4,     // D列：⇒
  COL_ARRIVAL_DATE: 5, // E列：EMS到着日

  COL_EMS_NO: 13,    // M列：EMS番号
  COL_BOX_NO: 14,    // N列：Box No.

  // A〜N列まで読む
  READ_START_COL: 1,
  READ_COLS: 14,

  // 読み込んだ配列内の位置（0始まり）
  IDX_NO: 0,         // A列
  IDX_DATE: 2,       // C列
  IDX_STATUS: 3,     // D列
  IDX_EMS_NO: 12,    // M列：EMS番号

  // 色を付ける範囲：M:N列
  COLOR_START_COL: 13,
  COLOR_COLS: 2,

  CHUNK_SIZE: 3000,

  CLEAR_STATUS_WHEN_NO_DATE: true,
  ARRIVAL_LEAD_DAYS: 3,
  AUTO_BOX_ON_EDIT: false,
  ONEDIT_MAX_ROWS: 100
};
/************************************************************
 * 50色パレット
 ************************************************************/
const EMS_COLORS = [
  '#FFF2CC', '#D9EAD3', '#CFE2F3', '#F4CCCC', '#EADCF8', '#D0E0E3', '#FCE5CD', '#EEEEEE',
  '#FFE599', '#B6D7A8', '#9FC5E8', '#EA9999', '#B4A7D6', '#A2C4C9', '#F9CB9C', '#D9D2E9',
  '#C9DAF8', '#D9E2F3', '#E2F0D9', '#FCE4D6', '#F8CBAD', '#DDEBF7', '#E4DFEC', '#DAEEF3',
  '#EBF1DE', '#F2DCDB', '#DCE6F1', '#E6E0EC', '#FDE9D9', '#DBEEF3', '#EAF2F8', '#E2EFDA',
  '#FFFACD', '#E0FFFF', '#F0FFF0', '#FFF0F5', '#F5F5DC', '#F0F8FF', '#FAEBD7', '#E6E6FA',
  '#FFE4E1', '#F5DEB3', '#D8BFD8', '#AFEEEE', '#98FB98', '#FFDAB9', '#B0E0E6', '#FFC0CB',
  '#D3D3D3', '#EEE8AA'
];


/************************************************************
 * onEdit（C列編集時にD列の⇒のみを最速反映）
 ************************************************************/
function EMS_onEdit_legacy_(e) {
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  if (!EMS_isTargetSheet_(sh)) return;

  const hitDate = EMS_rangeHitsColumn_(e.range, EMS_CFG.COL_EMS_DATE);
  const hitEms = EMS_rangeHitsColumn_(e.range, EMS_CFG.COL_EMS_NO);

  if (!hitDate && !hitEms) return;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(1000)) return;

  try {
    // C列：EMS発送日を編集したら、D列に ⇒ を入れる
    if (hitDate) {
      EMS_updateStatusOnlyForEditedRows_(sh, e.range);
    }

    if (
      EMS_CFG.AUTO_BOX_ON_EDIT &&
      e.range.getNumRows() <= EMS_CFG.ONEDIT_MAX_ROWS
    ) {
      EMS_updateDatesByRows_(sh, e.range.getRow(), e.range.getNumRows(), false);
    } else if (EMS_CFG.AUTO_BOX_ON_EDIT) {
      SpreadsheetApp.getActive().toast('大量編集のため、ボタンから全体更新を実行してください。');
    }

  } finally {
    lock.releaseLock();
  }
}


function EMS_全体更新() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (!EMS_isTargetSheet_(sh)) {
    SpreadsheetApp.getUi().alert('このシートはEMS自社用の見出しではないようです。\n6行目の構成を確認してください。');
    return;
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    EMS_updateAllSheet_(sh, true);
    EMS_updateBordersByEmsNo_(sh);
  } finally {
    lock.releaseLock();
  }
}

function EMSリスト_発送日基準に並べ替え() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (!EMS_isTargetSheet_(sh)) {
    SpreadsheetApp.getUi().alert('EMSリストで実行してください。');
    return;
  }

  const cfg = EMS_CFG;
  const lastRow = EMS_findLastDataRow_(sh);
  if (lastRow < cfg.START_ROW) {
    SpreadsheetApp.getActive().toast('並べ替え対象のデータ行がありません。');
    return;
  }

  const numRows = lastRow - cfg.START_ROW + 1;
  const lastCol = sh.getLastColumn();

  sh.getRange(cfg.START_ROW, 1, numRows, lastCol).sort([
    { column: cfg.COL_EMS_DATE, ascending: true },
    { column: cfg.COL_EMS_NO, ascending: true },
    { column: cfg.COL_NO, ascending: true },
  ]);

  SpreadsheetApp.flush();
  EMS_updateAllSheet_(sh, false);
  EMS_updateBordersByEmsNo_(sh);

  SpreadsheetApp.getActive().toast('EMSリストをEMS発送日基準で並べ替えました。');
}

// エディタ（アクティブシートに依存できない文脈）から実行する並べ替え
function EMSリスト_発送日基準に並べ替え_エディタ実行() {
  const sh = SpreadsheetApp.getActive().getSheetByName('EMSリスト');
  if (!sh || !EMS_isTargetSheet_(sh)) {
    Logger.log('EMSリストが見つからないか、見出しが想定と異なります。');
    return;
  }

  const cfg = EMS_CFG;
  const lastRow = EMS_findLastDataRow_(sh);
  if (lastRow < cfg.START_ROW) {
    Logger.log('並べ替え対象のデータ行がありません。');
    return;
  }

  const numRows = lastRow - cfg.START_ROW + 1;
  const lastCol = sh.getLastColumn();

  sh.getRange(cfg.START_ROW, 1, numRows, lastCol).sort([
    { column: cfg.COL_EMS_DATE, ascending: true },
    { column: cfg.COL_EMS_NO, ascending: true },
    { column: cfg.COL_NO, ascending: true },
  ]);

  SpreadsheetApp.flush();
  EMS_updateAllSheet_(sh, false);
  EMS_updateBordersByEmsNo_(sh);
  Logger.log('EMSリストをEMS発送日基準で並べ替えました（%s行）。', numRows);
}

function EMS_findLastDataRow_(sh) {
  const cfg = EMS_CFG;
  const lastRow = sh.getLastRow();
  if (lastRow < cfg.START_ROW) return cfg.START_ROW - 1;

  const numRows = lastRow - cfg.START_ROW + 1;
  const maxCol = Math.max(cfg.COL_EMS_NO, 13);
  const values = sh.getRange(cfg.START_ROW, 1, numRows, maxCol).getDisplayValues();
  const checkIdx = [0, 1, 2, 5, 8, 9, 12]; // A/B/C/F/I/J/M

  for (let i = values.length - 1; i >= 0; i--) {
    if (checkIdx.some(idx => String(values[i][idx] || '').trim() !== '')) {
      return cfg.START_ROW + i;
    }
  }
  return cfg.START_ROW - 1;
}


function EMS_選択行の発送日だけ更新() {
  const sh = SpreadsheetApp.getActiveSheet();

  if (!EMS_isTargetSheet_(sh)) {
    SpreadsheetApp.getUi().alert('このシートはEMS自社用の見出しではないようです。');
    return;
  }

  const range = sh.getActiveRange();
  if (!range) return;

  const cfg = EMS_CFG;

  const startRow = Math.max(range.getRow(), cfg.START_ROW);
  const endRow = range.getRow() + range.getNumRows() - 1;
  const numRows = endRow - startRow + 1;

  if (numRows <= 0) return;

  // 選択範囲のC列 EMS発送日を取得
  const selectedDates = sh
    .getRange(startRow, cfg.COL_EMS_DATE, numRows, 1)
    .getDisplayValues()
    .map(r => EMS_norm_(r[0]))
    .filter(v => v !== '');

  if (selectedDates.length === 0) {
    SpreadsheetApp.getUi().alert('選択行にEMS発送日がありません。');
    return;
  }

  const targetDateSet = new Set(selectedDates);

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const result = EMS_updateOnlySelectedDates_(sh, targetDateSet);
    EMS_updateBordersBySelectedDatesAndEmsNo_(sh, targetDateSet);

    SpreadsheetApp.getActive().toast(
      `選択発送日を更新：${result.targetRows}行 / ${result.totalBoxes}箱`
    );
  } finally {
    lock.releaseLock();
  }
}

function EMS_updateOnlySelectedDates_(sh, targetDateSet) {
  const cfg = EMS_CFG;

  const lastRow = sh.getLastRow();
  if (lastRow < cfg.START_ROW) {
    return { targetRows: 0, totalBoxes: 0 };
  }

  const numRows = lastRow - cfg.START_ROW + 1;

  const values = sh
    .getRange(cfg.START_ROW, cfg.READ_START_COL, numRows, cfg.READ_COLS)
    .getDisplayValues();

  const outputs = new Array(values.length);
  const boxMapByDate = new Map();

  let totalItemCount = 0;
  let targetRows = 0;
  const countedBoxKeySet = new Set();

  // 箱グループ（発送日＋EMS番号）の登場順。全体更新と色が揃うよう全行で数える
  const groupColorIndex = new Map();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];

    const dateKey = EMS_norm_(row[cfg.IDX_DATE]);      // C列
    const emsNo = EMS_normTrack_(row[cfg.IDX_EMS_NO]); // M列
    const currentStatus = row[cfg.IDX_STATUS];         // D列

    // A列No.は全体通し番号にするため、全行のEMS番号を数える
    if (emsNo) {
      totalItemCount++;
    }

    // 箱番号と色の通し順は全行ぶん先に確定させる（選択外でも数える）
    let boxNo = '';
    let colorIdx = -1;
    if (emsNo) {
      if (!boxMapByDate.has(dateKey)) {
        boxMapByDate.set(dateKey, new Map());
      }
      const dateMap = boxMapByDate.get(dateKey);
      if (!dateMap.has(emsNo)) {
        dateMap.set(emsNo, dateMap.size + 1);
      }
      boxNo = dateMap.get(emsNo);

      const groupKey = dateKey + '||' + emsNo;
      if (!groupColorIndex.has(groupKey)) {
        groupColorIndex.set(groupKey, groupColorIndex.size);
      }
      colorIdx = groupColorIndex.get(groupKey);
    }

    // 選択したEMS発送日以外は書き換えない
    if (!targetDateSet.has(dateKey)) {
      continue;
    }

    // EMS番号がない行は対象外にしたい場合はここでcontinue
    // 今回は「EMS番号があるところが対象」なので、空欄は触らない
    if (!emsNo) {
      continue;
    }

    targetRows++;

    const statusValue = dateKey ? '⇒' : (cfg.CLEAR_STATUS_WHEN_NO_DATE ? '' : currentStatus);

    countedBoxKeySet.add(dateKey + '||' + emsNo);

    // 50色を箱グループの登場順に順繰りで使う
    const color = EMS_COLORS[colorIdx % EMS_COLORS.length];

    outputs[i] = {
      no: totalItemCount,
      status: statusValue,
      boxNo: boxNo,
      bg: Array(cfg.COLOR_COLS).fill(color)
    };
  }

  EMS_writeOutputsInChunks_(sh, outputs);

  return {
    targetRows: targetRows,
    totalBoxes: countedBoxKeySet.size
  };
}

function EMS_updateBordersBySelectedDatesAndEmsNo_(sh, targetDateSet) {
  const cfg = EMS_CFG;

  const startRow = cfg.START_ROW;
  const lastRow = sh.getLastRow();
  if (lastRow < startRow) return;

  const numRows = lastRow - startRow + 1;

  // A列〜O列まで罫線対象
  const startCol = 1;
  const colCount = 15;

  const dateValues = sh
    .getRange(startRow, cfg.COL_EMS_DATE, numRows, 1)
    .getDisplayValues()
    .map(r => EMS_norm_(r[0]));

  const emsValues = sh
    .getRange(startRow, cfg.COL_EMS_NO, numRows, 1)
    .getDisplayValues()
    .map(r => EMS_normTrack_(r[0]));

  // 対象発送日の行だけ、まず罫線を消す
  for (let i = 0; i < numRows; i++) {
    const rowNo = startRow + i;
    const dateKey = dateValues[i];

    if (!targetDateSet.has(dateKey)) continue;

    sh.getRange(rowNo, startCol, 1, colCount)
      .setBorder(false, false, false, false, false, false);
  }

  let groupStartIndex = null;
  let currentDate = '';
  let currentEmsNo = '';

  for (let i = 0; i <= numRows; i++) {
    const isEnd = i === numRows;

    const dateKey = isEnd ? '' : dateValues[i];
    const emsNo = isEnd ? '' : emsValues[i];

    const isTarget =
      !isEnd &&
      targetDateSet.has(dateKey) &&
      emsNo !== '';

    if (groupStartIndex === null) {
      if (isTarget) {
        groupStartIndex = i;
        currentDate = dateKey;
        currentEmsNo = emsNo;
      }
      continue;
    }

    const sameGroup =
      isTarget &&
      dateKey === currentDate &&
      emsNo === currentEmsNo;

    if (isEnd || !sameGroup) {
      const groupStartRow = startRow + groupStartIndex;
      const groupRowCount = i - groupStartIndex;

      const groupRange = sh.getRange(
        groupStartRow,
        startCol,
        groupRowCount,
        colCount
      );

      // グループ内は細い格子
      groupRange.setBorder(
        true,
        true,
        true,
        true,
        true,
        true,
        '#999999',
        SpreadsheetApp.BorderStyle.SOLID
      );

      // 外枠を太線
      groupRange.setBorder(
        true,
        true,
        true,
        true,
        null,
        null,
        '#000000',
        SpreadsheetApp.BorderStyle.SOLID_THICK
      );

      if (isTarget) {
        groupStartIndex = i;
        currentDate = dateKey;
        currentEmsNo = emsNo;
      } else {
        groupStartIndex = null;
        currentDate = '';
        currentEmsNo = '';
      }
    }
  }
}


function EMS_色だけクリア() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (!EMS_isTargetSheet_(sh)) {
    SpreadsheetApp.getUi().alert('このシートはEMS自社用の見出しではないようです。');
    return;
  }

  const cfg = EMS_CFG;
  const lastRow = sh.getLastRow();
  if (lastRow < cfg.START_ROW) return;

  const numRows = lastRow - cfg.START_ROW + 1;
  sh.getRange(cfg.START_ROW, cfg.COLOR_START_COL, numRows, cfg.COLOR_COLS).setBackground(null);
  SpreadsheetApp.getActive().toast('色をクリアしました。');
}


function EMS_updateAllSheet_(sh, showToast) {
  const data = EMS_loadData_(sh);
  if (!data) {
    if (showToast) SpreadsheetApp.getActive().toast('更新対象の行がありません。');
    return;
  }

  const result = EMS_writeForTargetKeys_(sh, data.values, data.keys, null);

  if (showToast) {
    SpreadsheetApp.getActive().toast(`EMS更新完了：${result.targetRows}行 / 総商品点数: ${result.totalItems}点 / ${result.totalBoxes}箱`);
  }
}


function EMS_updateDatesByRows_(sh, row, numRows, showToast) {
  const data = EMS_loadData_(sh);
  if (!data) return;

  const cfg = EMS_CFG;
  const startIdx = Math.max(row, cfg.START_ROW) - cfg.START_ROW;
  const endIdx = Math.min(row + numRows - 1, cfg.START_ROW + data.values.length - 1) - cfg.START_ROW;

  if (startIdx < 0 || endIdx < 0 || startIdx > endIdx) return;

  const targetKeys = new Set();
  for (let i = startIdx; i <= endIdx; i++) {
    targetKeys.add(data.keys[i]);
  }

  const result = EMS_writeForTargetKeys_(sh, data.values, data.keys, targetKeys);

  if (showToast) {
    SpreadsheetApp.getActive().toast(`選択部分を更新：${result.targetRows}行 / ${result.totalBoxes}箱`);
  }
}


function EMS_loadData_(sh) {
  const cfg = EMS_CFG;
  const lastRow = sh.getLastRow();
  if (lastRow < cfg.START_ROW) return null;

  const numRows = lastRow - cfg.START_ROW + 1;
  const values = sh.getRange(cfg.START_ROW, cfg.READ_START_COL, numRows, cfg.READ_COLS).getDisplayValues();

  const keys = [];

  for (let i = 0; i < values.length; i++) {
    keys.push(EMS_boxDateKeyFromRow_(values[i]));
  }

  return { values, keys };
}

function EMS_boxDateKeyFromRow_(row) {
  const cfg = EMS_CFG;
  const shipDate = EMS_norm_(row[cfg.IDX_DATE]); // C列 EMS発送日
  if (shipDate) return shipDate;

  // EMS発送日が空欄なら、まだ箱に入っていない扱い。
  // B列の入荷日は箱番号・色・罫線の判定には使わない。
  return '';
}


function EMS_writeForTargetKeys_(sh, values, keys, targetKeys) {
  const cfg = EMS_CFG;
  const boxMapByDate = new Map();
  const outputs = new Array(values.length);

  let targetRows = 0;
  let totalItemCount = 0; 
  const countedBoxKeySet = new Set();
  // 箱グループ（発送日＋EMS番号）の登場順。部分更新でも色が変わらないよう全行で数える
  const groupColorIndex = new Map();

  for (let i = 0; i < values.length; i++) {
    const dateKey = keys[i];
    const row = values[i];
    const emsNo = EMS_normTrack_(row[cfg.IDX_EMS_NO]);

    let boxNo = '';
    let colorIdx = -1;

    if (emsNo) {
      totalItemCount++;
    }

    if (emsNo && dateKey) {
      if (!boxMapByDate.has(dateKey)) {
        boxMapByDate.set(dateKey, new Map());
      }
      const dateMap = boxMapByDate.get(dateKey);
      if (!dateMap.has(emsNo)) {
        dateMap.set(emsNo, dateMap.size + 1);
      }
      boxNo = dateMap.get(emsNo);

      const groupKey = dateKey + '||' + emsNo;
      if (!groupColorIndex.has(groupKey)) {
        groupColorIndex.set(groupKey, groupColorIndex.size);
      }
      colorIdx = groupColorIndex.get(groupKey);
    }

    if (targetKeys && !targetKeys.has(dateKey)) {
      continue;
    }

    targetRows++;

    const dateCell = EMS_norm_(row[cfg.IDX_DATE]);
    const currentStatus = row[cfg.IDX_STATUS];

    const statusValue = dateCell ? '⇒' : (cfg.CLEAR_STATUS_WHEN_NO_DATE ? '' : currentStatus);

    let noValue = ''; 
    let bg = Array(cfg.COLOR_COLS).fill('#ffffff');

    if (emsNo) {
      noValue = totalItemCount; 
    }

    if (emsNo && dateKey) {
      countedBoxKeySet.add(dateKey + '||' + emsNo);

      // 50色を箱グループの登場順に順繰りで使う
      const color = EMS_COLORS[colorIdx % EMS_COLORS.length];
      bg = Array(cfg.COLOR_COLS).fill(color);
    }

    outputs[i] = {
      no: noValue, 
      status: statusValue,
      boxNo: boxNo,
      bg: bg
    };
  }

  EMS_writeOutputsInChunks_(sh, outputs);

  return {
    targetRows: targetRows,
    totalBoxes: countedBoxKeySet.size,
    totalItems: totalItemCount
  };
}


function EMS_writeOutputsInChunks_(sh, outputs) {
  const cfg = EMS_CFG;

  let startIdx = null;
  let nos = [];         
  let statuses = [];
  let boxNos = [];
  let backgrounds = [];

  for (let i = 0; i <= outputs.length; i++) {
    const out = outputs[i];

    if (out) {
      if (startIdx === null) startIdx = i;

      nos.push([out.no]); 
      statuses.push([out.status]);
      boxNos.push([out.boxNo]);
      backgrounds.push(out.bg);

      if (statuses.length >= cfg.CHUNK_SIZE) {
        EMS_flushSegment_(sh, startIdx, nos, statuses, boxNos, backgrounds);
        startIdx = null;
        nos = [];
        statuses = [];
        boxNos = [];
        backgrounds = [];
      }
    } else {
      if (startIdx !== null) {
        EMS_flushSegment_(sh, startIdx, nos, statuses, boxNos, backgrounds);
        startIdx = null;
        nos = [];
        statuses = [];
        boxNos = [];
        backgrounds = [];
      }
    }
  }
}


function EMS_flushSegment_(sh, startIdx, nos, statuses, boxNos, backgrounds) {
  if (!statuses.length) return;

  const cfg = EMS_CFG;
  const row = cfg.START_ROW + startIdx;
  const numRows = statuses.length;

  sh.getRange(row, cfg.COL_NO, numRows, 1).setValues(nos); 
  sh.getRange(row, cfg.COL_STATUS, numRows, 1).setValues(statuses);
  sh.getRange(row, cfg.COL_BOX_NO, numRows, 1).setValues(boxNos);
  sh.getRange(row, cfg.COLOR_START_COL, numRows, cfg.COLOR_COLS).setBackgrounds(backgrounds);
}


function EMS_updateStatusOnlyForEditedRows_(sh, editedRange) {
  const cfg = EMS_CFG;
  const startRow = Math.max(editedRange.getRow(), cfg.START_ROW);
  const endRow = editedRange.getRow() + editedRange.getNumRows() - 1;
  const numRows = endRow - startRow + 1;

  if (numRows <= 0) return;

  const dateValues = sh.getRange(startRow, cfg.COL_EMS_DATE, numRows, 1).getDisplayValues();
  const statusRange = sh.getRange(startRow, cfg.COL_STATUS, numRows, 1);
  const currentStatuses = statusRange.getValues();

  const newStatuses = dateValues.map((r, i) => {
    const hasDate = EMS_norm_(r[0]) !== '';
    if (hasDate) return ['⇒'];
    return [cfg.CLEAR_STATUS_WHEN_NO_DATE ? '' : currentStatuses[i][0]];
  });

  statusRange.setValues(newStatuses);
  EMS_fillArrivalEstimateRows_(sh, startRow, numRows, false);
  EMS_updateAllSheet_(sh, false);
  EMS_updateBordersByEmsNo_(sh);
}

function EMS_到着予定日を補完() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (!EMS_isTargetSheet_(sh)) {
    SpreadsheetApp.getUi().alert('EMSリストで実行してください。');
    return;
  }

  const cfg = EMS_CFG;
  const lastRow = EMS_findLastDataRow_(sh);
  if (lastRow < cfg.START_ROW) {
    SpreadsheetApp.getActive().toast('補完対象の行がありません。');
    return;
  }

  const count = EMS_fillArrivalEstimateRows_(sh, cfg.START_ROW, lastRow - cfg.START_ROW + 1, false);
  SpreadsheetApp.getActive().toast(`EMS到着予定日を補完しました：${count}件`);
}

function EMS_fillArrivalEstimateRows_(sh, startRow, numRows, overwrite) {
  const cfg = EMS_CFG;
  if (numRows <= 0) return 0;

  const shipRange = sh.getRange(startRow, cfg.COL_EMS_DATE, numRows, 1);
  const shipValues = shipRange.getValues();
  const shipDisplays = shipRange.getDisplayValues();
  const arrivalRange = sh.getRange(startRow, cfg.COL_ARRIVAL_DATE, numRows, 1);
  const arrivalValues = arrivalRange.getValues();
  const arrivalDisplays = arrivalRange.getDisplayValues();

  let changed = false;
  let count = 0;
  for (let i = 0; i < numRows; i++) {
    const shipDate = EMS_toDateOnly_(shipValues[i][0]) || EMS_toDateOnly_(shipDisplays[i][0]);
    if (!shipDate) continue;
    const hasArrival = EMS_norm_(arrivalDisplays[i][0]) !== '';
    if (hasArrival && !overwrite) continue;

    arrivalValues[i][0] = EMS_addDays_(shipDate, cfg.ARRIVAL_LEAD_DAYS);
    changed = true;
    count++;
  }

  if (changed) arrivalRange.setValues(arrivalValues);
  return count;
}

function EMS_addDays_(date, days) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate() + days);
}

function EMS_toDateOnly_(value) {
  if (typeof _emsDateOnly_ === 'function') return _emsDateOnly_(value);

  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  if (typeof value === 'number' && isFinite(value)) {
    const d = new Date(Date.UTC(1899, 11, 30) + Math.round(value) * 24 * 60 * 60 * 1000);
    return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  }

  const text = String(value || '')
    .normalize('NFKC')
    .replace(/\(.+?\)/g, '')
    .replace(/（.+?）/g, '')
    .replace(/[年月]/g, '/')
    .replace(/日/g, '')
    .replace(/\s+/g, '')
    .trim();
  if (!text) return null;

  let m = text.match(/^(\d{2,4})[\/.-](\d{1,2})[\/.-](\d{1,2})$/);
  if (m) {
    let y = Number(m[1]);
    if (y < 100) y += 2000;
    return new Date(y, Number(m[2]) - 1, Number(m[3]));
  }

  m = text.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));

  return null;
}


function EMS_isTargetSheet_(sh) {
  const cfg = EMS_CFG;

  const hDate = EMS_norm_(sh.getRange(cfg.HEADER_ROW, cfg.COL_EMS_DATE).getDisplayValue());
  const hEms = EMS_norm_(sh.getRange(cfg.HEADER_ROW, cfg.COL_EMS_NO).getDisplayValue());
  const hBox = EMS_norm_(sh.getRange(cfg.HEADER_ROW, cfg.COL_BOX_NO).getDisplayValue())
    .replace(/\s/g, '').replace(/　/g, '').replace('．', '.').replace('。', '.').toLowerCase();

  if (hDate.includes('EMS発送日') && hEms.includes('EMS番号') && hBox.includes('boxno')) {
    return true;
  }

  const headerValues = sh.getRange(cfg.HEADER_ROW, 1, 1, sh.getLastColumn()).getDisplayValues()[0]
    .map(v => EMS_norm_(v).replace(/\s/g, '').replace(/　/g, '').replace('．', '.').replace('。', '.').toLowerCase());

  const hasDate = headerValues.some(v => v.includes('ems発送日'));
  const hasEms = headerValues.some(v => v.includes('ems番号'));
  const hasBox = headerValues.some(v => v.includes('boxno'));

  return hasDate && hasEms && hasBox;
}


function EMS_rangeHitsColumn_(range, col) {
  const startCol = range.getColumn();
  const endCol = startCol + range.getNumColumns() - 1;
  return startCol <= col && col <= endCol;
}


function EMS_norm_(value) {
  return String(value == null ? '' : value).replace(/\u00A0/g, ' ').trim();
}

function EMS_normTrack_(value) {
  if (typeof normTrack_ === 'function') return normTrack_(value);
  return String(value == null ? '' : value)
    .normalize('NFKC')
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, '');
}
// =====================================================
// EMS番号ごとに見やすい罫線を入れる
// 同じEMS番号の連続行を1グループとして外枠を太くする
// =====================================================
function EMS_EMS番号ごとに罫線を更新() {
  const sh = SpreadsheetApp.getActiveSheet();

  if (!EMS_isTargetSheet_(sh)) {
    SpreadsheetApp.getUi().alert('このシートはEMSリストではないようです。');
    return;
  }

  EMS_updateBordersByEmsNo_(sh);
  SpreadsheetApp.getActive().toast('EMS番号ごとの罫線を更新しました。');
}

function EMS_updateBordersByEmsNo_(sh) {
  const cfg = EMS_CFG;
  const startRow = cfg.START_ROW;        // 7行目
  const lastRow = sh.getLastRow();
  if (lastRow < startRow) return;

  const numRows = lastRow - startRow + 1;
  const startCol = 1;
  const colCount = 15; // A列〜O列
  const emsNoCol = cfg.COL_EMS_NO;       // M列 = 13

  // 対象範囲の罫線を一度全部消す
  const allRange = sh.getRange(startRow, startCol, numRows, colCount);
  allRange.setBorder(false, false, false, false, false, false);

  const values = sh
    .getRange(startRow, 1, numRows, Math.max(emsNoCol, cfg.COL_ARRIVAL_DATE))
    .getDisplayValues();
  const groupKeys = values.map(row => {
    const emsNo = EMS_normTrack_(row[emsNoCol - 1]);
    if (!emsNo) return '';
    const dateKey = EMS_boxDateKeyFromRow_(row);
    if (!dateKey) return '';
    return dateKey + '||' + emsNo;
  });

  let groupStartIndex = null;
  let currentGroupKey = '';

  for (let i = 0; i <= numRows; i++) {
    const isEnd = i === numRows;
    const groupKey = isEnd ? '' : groupKeys[i];

    if (groupStartIndex === null) {
      if (!isEnd && groupKey !== '') {
        groupStartIndex = i;
        currentGroupKey = groupKey;
      }
      continue;
    }

    // EMS発送日（なければB列日付）＋EMS番号が変わったら1グループ確定
    if (isEnd || groupKey !== currentGroupKey) {
      const groupStartRow = startRow + groupStartIndex;
      const groupRowCount = i - groupStartIndex;

      const groupRange = sh.getRange(groupStartRow, startCol, groupRowCount, colCount);

      // 1. まずグループ全体に細い格子（グレー）を引く
      groupRange.setBorder(
        true, true, true, true, true, true,
        '#999999',
        SpreadsheetApp.BorderStyle.SOLID
      );

      // 2. 【修正】グループの外枠（上下左右）を一発で黒の太線にする
      // これにより1行だけのデータでもバグらず、綺麗に四方を閉じることができます
      groupRange.setBorder(
        true, true, true, true, null, null,
        '#000000',
        SpreadsheetApp.BorderStyle.SOLID_THICK
      );

      // 次のグループ開始
      if (!isEnd && groupKey !== '') {
        groupStartIndex = i;
        currentGroupKey = groupKey;
      } else {
        groupStartIndex = null;
        currentGroupKey = '';
      }
    }
  }
}

// =====================================================
// EMS番号の空白削除
// M列：EMS番号 / 7行目以下
// 半角スペース・全角スペース・ノーブレークスペースを削除
// =====================================================
function EMS_EMS番号の空白を削除() {
  const sh = SpreadsheetApp.getActiveSheet();

  if (!EMS_isTargetSheet_(sh)) {
    SpreadsheetApp.getUi().alert('このシートはEMSリストではないようです。');
    return;
  }

  const startRow = 7;
  const emsCol = 13; // M列
  const lastRow = sh.getLastRow();

  if (lastRow < startRow) {
    SpreadsheetApp.getActive().toast('対象行がありません。');
    return;
  }

  const numRows = lastRow - startRow + 1;
  const range = sh.getRange(startRow, emsCol, numRows, 1);
  const values = range.getDisplayValues();

  let changedCount = 0;

  const newValues = values.map(row => {
    const original = String(row[0] || '');

    const cleaned = EMS_normTrack_(original);

    if (original !== cleaned) changedCount++;

    return [cleaned];
  });

  range.setValues(newValues);

  SpreadsheetApp.getActive().toast(
    `EMS番号の空白を削除しました：${changedCount}件`
  );
}
