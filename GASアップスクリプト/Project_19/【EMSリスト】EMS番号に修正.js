/************************************************************
 * EMS自社用ファイル 自動処理 【列番号確定版】
 *
 * 実際のシートのアルファベットに完全に合わせました：
 * A列：No.（商品点数カウンタ）
 * C列：EMS発送日
 * D列：ステータス列（⇒）
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

  CLEAR_STATUS_WHEN_NO_DATE: false,
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
    const emsNo = EMS_norm_(row[cfg.IDX_EMS_NO]);      // M列
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
    .map(r => EMS_norm_(r[0]));

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
  let currentDateKey = '';

  for (let i = 0; i < values.length; i++) {
    const dateCell = EMS_norm_(values[i][cfg.IDX_DATE]);
    if (dateCell) {
      currentDateKey = dateCell;
    }
    keys.push(currentDateKey || '日付未入力');
  }

  return { values, keys };
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
    const emsNo = EMS_norm_(row[cfg.IDX_EMS_NO]);

    let boxNo = '';
    let colorIdx = -1;

    if (emsNo) {
      totalItemCount++;

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
    let bg = Array(cfg.COLOR_COLS).fill(null);

    if (emsNo) {
      noValue = totalItemCount; 
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

  const emsValues = sh
    .getRange(startRow, emsNoCol, numRows, 1)
    .getDisplayValues()
    .map(r => String(r[0] || '').trim());

  let groupStartIndex = null;
  let currentEmsNo = '';

  for (let i = 0; i <= numRows; i++) {
    const isEnd = i === numRows;
    const emsNo = isEnd ? '' : emsValues[i];

    if (groupStartIndex === null) {
      if (!isEnd && emsNo !== '') {
        groupStartIndex = i;
        currentEmsNo = emsNo;
      }
      continue;
    }

    // EMS番号が変わった、または最後まで来たら1グループ確定
    if (isEnd || emsNo !== currentEmsNo) {
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
      if (!isEnd && emsNo !== '') {
        groupStartIndex = i;
        currentEmsNo = emsNo;
      } else {
        groupStartIndex = null;
        currentEmsNo = '';
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

    const cleaned = original
      .replace(/ /g, '')        // 半角スペース
      .replace(/　/g, '')       // 全角スペース
      .replace(/\u00A0/g, '');  // ノーブレークスペース

    if (original !== cleaned) changedCount++;

    return [cleaned];
  });

  range.setValues(newValues);

  SpreadsheetApp.getActive().toast(
    `EMS番号の空白を削除しました：${changedCount}件`
  );
}
