/******************************************************
 * ★ Qoo10在庫照合 専用ファイル（超高速版）
 *    Sheets API batchUpdate + スパース書き込み
 ******************************************************/

/******************************************************
 * 色定数
 ******************************************************/
var QOO10_COLOR_LESS = '#F4C7C3';   // 🔴 Qoo10が少ない
var QOO10_COLOR_MORE = '#C9DAF8';   // 🔵 Qoo10が多い
var QOO10_COLOR_WARN = '#FFE599';   // 🟡 Qoo10にだけ在庫あり
var QOO10_COLOR_DUP  = '#E1BEE7';   // 🟣 重複

var COLOR_SOKUNO = '#D9EAD3';
var COLOR_OTORI  = '#FFF2CC';

if (typeof AMZ_COLOR_NONE === 'undefined') var AMZ_COLOR_NONE = null;
if (typeof AMZ_COLOR_LESS === 'undefined') var AMZ_COLOR_LESS = QOO10_COLOR_LESS;
if (typeof AMZ_COLOR_MORE === 'undefined') var AMZ_COLOR_MORE = QOO10_COLOR_MORE;
if (typeof AMZ_COLOR_WARN === 'undefined') var AMZ_COLOR_WARN = QOO10_COLOR_WARN;
if (typeof AMZ_COLOR_SOKUNO === 'undefined') var AMZ_COLOR_SOKUNO = COLOR_SOKUNO;
if (typeof AMZ_COLOR_OTORI === 'undefined') var AMZ_COLOR_OTORI = COLOR_OTORI;

var COLOR_Q_LESS = QOO10_COLOR_LESS;
var COLOR_Q_MORE = QOO10_COLOR_MORE;
var COLOR_Q_WARN = QOO10_COLOR_WARN;

/******************************************************
 * 列定義（Yahoo在庫ビューのQoo10エリア）
 ******************************************************/
var QOO10_COL = {
  CHECK: 14, // N: 反映Qoo10
  SKU:   15, // O: Qoo10SKU
  QTY:   16, // P: Qoo10在庫
  DIFF:  17, // Q: 差分(Q-Y)
  JUDGE: 18  // R: 判定(Qoo10)
};


/******************************************************
 * メイン：Qoo10在庫照合（超高速版）
 ******************************************************/
function Qoo10在庫照合() {
  var t0 = new Date().getTime();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssId = ss.getId();

  console.log('[Qoo10在庫照合] START');

  var Y = ss.getSheetByName('Yahoo在庫ビュー');
  var Q = ss.getSheetByName('Qoo10在庫取込');
  if (!Y) throw new Error('Yahoo在庫ビューが見つかりません');
  if (!Q) throw new Error('Qoo10在庫取込が見つかりません');

  var HEADER_ROW = 2;
  var DATA_START_ROW = 3;
  var shYId = Y.getSheetId();

  // 1) Yahoo 行数
  var yLast = getLastDataRowByCol_API_(ssId, 'Yahoo在庫ビュー', 1, DATA_START_ROW);
  var dataRows = yLast - HEADER_ROW;
  console.log('[Qoo10在庫照合] Yahoo行数: ' + dataRows);

  if (yLast < DATA_START_ROW) {
    ss.toast('Yahoo在庫ビューにデータがありません', 'Qoo10在庫照合', 5);
    return;
  }

  ss.toast('Qoo10在庫照合を開始します…', 'Qoo10在庫照合', 10);

  // 2) エリアリセット（Sheets API）
  var tReset = new Date().getTime();
  var clearRange = "'Yahoo在庫ビュー'!N" + DATA_START_ROW + ":R" + yLast;
  Sheets.Spreadsheets.Values.clear({}, ssId, clearRange);
  
  // 背景色リセット
  Sheets.Spreadsheets.batchUpdate({
    requests: [{
      repeatCell: {
        range: {
          sheetId: shYId,
          startRowIndex: DATA_START_ROW - 1,
          endRowIndex: yLast,
          startColumnIndex: QOO10_COL.CHECK - 1,
          endColumnIndex: QOO10_COL.JUDGE
        },
        cell: { userEnteredFormat: { backgroundColor: { red: 1, green: 1, blue: 1 } } },
        fields: 'userEnteredFormat.backgroundColor'
      }
    }]
  }, ssId);

  // ヘッダー & チェックボックス
  Y.getRange(HEADER_ROW, QOO10_COL.CHECK, 1, 5).setValues([[
    '反映Qoo10', 'Qoo10SKU', 'Qoo10在庫', '差分(Q-Y)', '判定(Qoo10)'
  ]]).setFontWeight('bold');

  var cbRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  Y.getRange(DATA_START_ROW, QOO10_COL.CHECK, dataRows, 1).setDataValidation(cbRule);
  console.log('[Qoo10在庫照合] Reset done: ' + (new Date().getTime() - tReset) + 'ms');

  // 3) Qoo10在庫取込 読み込み
  var tQ = new Date().getTime();
  var qData = buildQoo10Map_(ssId, Q);
  var qMap = qData.map;
  var iMap = qData.itemNoMap;
  var duplicates = qData.duplicates;
  console.log('[Qoo10在庫照合] Qoo10読込: ' + qMap.size + '件 (' + (new Date().getTime() - tQ) + 'ms)');

  // 4) Yahoo在庫ビュー A〜F を読み込み
  var tY = new Date().getTime();
  var yRange = "'Yahoo在庫ビュー'!A" + DATA_START_ROW + ":F" + yLast;
  var yResult = Sheets.Spreadsheets.Values.get(ssId, yRange, { valueRenderOption: 'UNFORMATTED_VALUE' });
  var yVals = yResult.values || [];
  console.log('[Qoo10在庫照合] Yahoo読込: ' + yVals.length + '行 (' + (new Date().getTime() - tY) + 'ms)');

  // 5) スパース出力データ作成
  var sparseData = [];  // { row, values: [4列], color }
  var diffRows = [];
  var unregRows = [];

  var dupSet = new Set();
  if (duplicates) {
    for (var i = 0; i < duplicates.length; i++) {
      dupSet.add(duplicates[i].key);
    }
  }

  for (var idx = 0; idx < yVals.length; idx++) {
    var row = yVals[idx] || [];
    var parent = String(row[0] || '').trim();
    var varCode = String(row[2] || '').trim();
    var unionKey = String(row[5] || '').trim();

    if (!parent && !unionKey && !varCode) continue;

    var sokuno = Number(row[3] || 0) || 0;
    var otori = Number(row[4] || 0) || 0;

    var fullKey = unionKey || varCode || parent;
    if (!fullKey) continue;

    var totalY = sokuno + otori;

    var yahooQty = 0;
    var yahooMode = '在庫なし';

    if (looksLikeB_(parent, varCode)) {
      yahooQty = otori;
      yahooMode = (otori > 0) ? 'お取り寄せ' : '在庫なし';
    } else {
      yahooQty = sokuno;
      yahooMode = (sokuno > 0) ? '即納' : '在庫なし';
    }

    // Qoo10 側の在庫取得
    var qQty = 0;
    var itemNoStr = '';

    if (qMap.has(fullKey)) {
      qQty = qMap.get(fullKey);
      if (iMap.has(fullKey)) itemNoStr = iMap.get(fullKey);
    } else {
      var parentKey = stripOptionSuffix_(fullKey);
      if (qMap.has(parentKey)) {
        qQty = qMap.get(parentKey);
        if (iMap.has(parentKey)) itemNoStr = iMap.get(parentKey);
      }
    }

    // 双方 0 は無視
    if (totalY === 0 && qQty === 0) continue;

    var rowData = [fullKey, '', '', ''];  // SKU, QTY, DIFF, JUDGE
    var rowColor = null;

    if (totalY > 0 && qQty === 0) {
      rowData[3] = 'Qoo10未登録/在庫0';
      rowColor = QOO10_COLOR_WARN;
      unregRows.push([fullKey, totalY, yahooMode, 'Qoo10未登録/在庫0']);
    } else if (totalY === 0 && qQty === 0) {
      // 両方0の場合は色なし（スキップされるはずだが念のため）
      rowData[3] = '';
      rowColor = null;
    } else {
      rowData[1] = qQty;
      var diff = qQty - yahooQty;
      rowData[2] = diff;

      var memo = '';

      if (totalY === 0 && qQty > 0) {
        rowData[3] = 'Qoo10にだけ在庫あり';
        // Yahoo側に在庫がないので色なし
        rowColor = null;
        memo = 'Qoo10にだけ在庫あり';
      } else if (diff === 0) {
        rowData[3] = 'OK（' + yahooMode + '：' + yahooQty + '）';
      } else if (diff > 0) {
        rowData[3] = 'Qoo10が多い';
        rowColor = QOO10_COLOR_MORE;
        memo = 'Qoo10の方が在庫多い';
      } else {
        rowData[3] = 'Qoo10が少ない';
        rowColor = QOO10_COLOR_LESS;
        memo = 'Qoo10の方が在庫少ない';
      }

      if (memo) {
        if (dupSet.has(fullKey)) memo += ' ※重複あり';
        diffRows.push([fullKey, qQty, totalY, yahooMode, memo, itemNoStr]);
      }
    }

    sparseData.push({
      row: DATA_START_ROW + idx,
      values: rowData,
      color: rowColor
    });
  }

  console.log('[Qoo10在庫照合] Sparse data: ' + sparseData.length + '件');

  // 6) スパース書き込み
  var tW = new Date().getTime();

  if (sparseData.length > 0) {
    sparseData.sort(function(a, b) { return a.row - b.row; });

    var valueRequests = [];
    var colorRequests = [];

    var i = 0;
    while (i < sparseData.length) {
      var startIdx = i;
      var startRow = sparseData[i].row;

      while (i < sparseData.length - 1 && sparseData[i + 1].row === sparseData[i].row + 1) {
        i++;
      }

      var endIdx = i;
      var numRows = endIdx - startIdx + 1;

      var values = [];
      for (var j = startIdx; j <= endIdx; j++) {
        values.push(sparseData[j].values);
      }

      var rangeStr = "'Yahoo在庫ビュー'!O" + startRow + ":R" + (startRow + numRows - 1);
      valueRequests.push({
        range: rangeStr,
        values: values
      });

      i++;
    }

    // バッチで値書き込み
    if (valueRequests.length > 0) {
      Sheets.Spreadsheets.Values.batchUpdate({
        valueInputOption: 'RAW',
        data: valueRequests
      }, ssId);
    }

    // 色付け
    var coloredRows = sparseData.filter(function(d) { return d.color !== null; });

    if (coloredRows.length > 0) {
      var colorBatchRequests = [];
      var colorGroups = {};

      for (var k = 0; k < coloredRows.length; k++) {
        var c = coloredRows[k].color;
        if (!colorGroups[c]) colorGroups[c] = [];
        colorGroups[c].push(coloredRows[k].row);
      }

      for (var color in colorGroups) {
        var rows = colorGroups[color].sort(function(a, b) { return a - b; });
        var ranges = mergeConsecutiveRows_Q_(rows);

        var rgb = hexToRgb_Q_(color);

        for (var m = 0; m < ranges.length; m++) {
          colorBatchRequests.push({
            repeatCell: {
              range: {
                sheetId: shYId,
                startRowIndex: ranges[m].start - 1,
                endRowIndex: ranges[m].end - 1,
                startColumnIndex: QOO10_COL.CHECK - 1,
                endColumnIndex: QOO10_COL.JUDGE
              },
              cell: {
                userEnteredFormat: { backgroundColor: rgb }
              },
              fields: 'userEnteredFormat.backgroundColor'
            }
          });
        }
      }

      if (colorBatchRequests.length > 0) {
        Sheets.Spreadsheets.batchUpdate({ requests: colorBatchRequests }, ssId);
      }
    }
  }

  console.log('[Qoo10在庫照合] Write done: ' + (new Date().getTime() - tW) + 'ms');

  // 7) レポート更新
  Qoo10レポートシート更新_(ssId, diffRows, unregRows, duplicates);

  var totalTime = (new Date().getTime() - t0) / 1000;
  ss.toast('Qoo10在庫照合 完了 (' + totalTime.toFixed(1) + '秒)', 'Qoo10在庫照合', 5);
  console.log('[Qoo10在庫照合] END: ' + totalTime + '秒');
}


/**
 * 連続する行番号をマージ
 */
function mergeConsecutiveRows_Q_(rows) {
  if (rows.length === 0) return [];

  var ranges = [];
  var start = rows[0];
  var end = rows[0];

  for (var i = 1; i < rows.length; i++) {
    if (rows[i] === end + 1) {
      end = rows[i];
    } else {
      ranges.push({ start: start, end: end + 1 });
      start = rows[i];
      end = rows[i];
    }
  }
  ranges.push({ start: start, end: end + 1 });

  return ranges;
}

/**
 * HEX色をRGBオブジェクトに変換
 */
function hexToRgb_Q_(hex) {
  var result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  return result ? {
    red: parseInt(result[1], 16) / 255,
    green: parseInt(result[2], 16) / 255,
    blue: parseInt(result[3], 16) / 255
  } : { red: 1, green: 1, blue: 1 };
}


/******************************************************
 * Qoo10 レポートシート更新（高速版）
 ******************************************************/
function Qoo10レポートシート更新_(ssId, diffRows, unregRows, duplicates) {
  var ss = SpreadsheetApp.openById(ssId);

  // ===============================
  // 1) 差異一覧_Qoo10のみ
  // ===============================
  var shDiff = ss.getSheetByName('差異一覧_Qoo10のみ');
  if (!shDiff) shDiff = ss.insertSheet('差異一覧_Qoo10のみ');

  var diffHeader = ['Qoo10SKU', 'Qoo10在庫', 'Yahoo在庫', 'Yahoo区分', 'メモ', '商品番号'];

  var maxRows = shDiff.getMaxRows();
  var maxCols = shDiff.getMaxColumns();
  if (maxRows > 1 && maxCols > 0) {
    shDiff.getRange(2, 1, maxRows - 1, maxCols).clearContent().setBackground(null);
  }

  shDiff.getRange(1, 1, 1, diffHeader.length).setValues([diffHeader]).setFontWeight('bold');

  var diffLen = diffRows ? diffRows.length : 0;
  if (diffLen > 0) {
    var dupSet = new Set();
    if (duplicates && duplicates.length) {
      for (var i = 0; i < duplicates.length; i++) {
        var key = String(duplicates[i].key || '').trim();
        if (key) dupSet.add(key);
      }
    }

    shDiff.getRange(2, 1, diffLen, diffHeader.length).setValues(diffRows);

    var rowColors = [];
    var shipColors = [];

    for (var i = 0; i < diffLen; i++) {
      var sku = String(diffRows[i][0] || '').trim();
      var qQty = Number(diffRows[i][1] || 0);
      var yQty = Number(diffRows[i][2] || 0);
      var yKubun = String(diffRows[i][3] || '');

      var color = null;

      if (dupSet.has(sku)) {
        color = QOO10_COLOR_DUP;
      } else if (yQty === 0 && qQty > 0) {
        color = QOO10_COLOR_WARN;
      } else if (qQty > yQty) {
        color = QOO10_COLOR_MORE;
      } else if (qQty < yQty) {
        color = QOO10_COLOR_LESS;
      }

      var cRow = [];
      for (var j = 0; j < diffHeader.length; j++) {
        cRow.push(color);
      }
      rowColors.push(cRow);

      var cShip = null;
      if (yKubun === '即納') cShip = COLOR_SOKUNO;
      else if (yKubun === 'お取り寄せ') cShip = COLOR_OTORI;
      shipColors.push([cShip]);
    }

    shDiff.getRange(2, 1, diffLen, diffHeader.length).setBackgrounds(rowColors);
    shDiff.getRange(2, 4, diffLen, 1).setBackgrounds(shipColors);
  }

  // ===============================
  // 2) 未登録一覧_Qoo10（要登録）
  // ===============================
  var shUnreg = ss.getSheetByName('未登録一覧_Qoo10（要登録）');
  if (!shUnreg) shUnreg = ss.insertSheet('未登録一覧_Qoo10（要登録）');

  var unregHeader = ['Qoo10SKU', 'Yahoo現在庫', 'Yahoo状態', 'メモ'];

  var maxRows2 = shUnreg.getMaxRows();
  var maxCols2 = shUnreg.getMaxColumns();
  if (maxRows2 > 1 && maxCols2 > 0) {
    shUnreg.getRange(2, 1, maxRows2 - 1, maxCols2).clearContent().setBackground(null);
  }

  shUnreg.getRange(1, 1, 1, unregHeader.length).setValues([unregHeader]).setFontWeight('bold');

  var unregLen = unregRows ? unregRows.length : 0;
  if (unregLen > 0) {
    shUnreg.getRange(2, 1, unregLen, unregHeader.length).setValues(unregRows);

    var rowColors2 = [];
    var stateColors = [];

    for (var i = 0; i < unregLen; i++) {
      var cRow2 = [];
      for (var j = 0; j < unregHeader.length; j++) {
        cRow2.push(QOO10_COLOR_WARN);
      }
      rowColors2.push(cRow2);

      var state = String(unregRows[i][2] || '');
      var cState = null;
      if (state === '即納') cState = COLOR_SOKUNO;
      else if (state === 'お取り寄せ') cState = COLOR_OTORI;
      stateColors.push([cState]);
    }

    shUnreg.getRange(2, 1, unregLen, unregHeader.length).setBackgrounds(rowColors2);
    shUnreg.getRange(2, 3, unregLen, 1).setBackgrounds(stateColors);
  }

  // ===============================
  // 3) 重複一覧_Qoo10
  // ===============================
  if (duplicates && duplicates.length > 0) {
    var shDup = ss.getSheetByName('重複一覧_Qoo10');
    if (!shDup) shDup = ss.insertSheet('重複一覧_Qoo10');

    var dupHeader = ['Qoo10SKU', '重複回数', 'メモ'];

    var maxRowsD = shDup.getMaxRows();
    var maxColsD = shDup.getMaxColumns();
    if (maxRowsD > 1 && maxColsD > 0) {
      shDup.getRange(2, 1, maxRowsD - 1, maxColsD).clearContent().setBackground(null);
    }

    shDup.getRange(1, 1, 1, dupHeader.length).setValues([dupHeader]).setFontWeight('bold');

    var dupData = [];
    for (var i = 0; i < duplicates.length; i++) {
      dupData.push([duplicates[i].key, duplicates[i].count, '重複登録されています']);
    }

    if (dupData.length > 0) {
      shDup.getRange(2, 1, dupData.length, dupHeader.length).setValues(dupData);

      var dupColors = [];
      for (var i = 0; i < dupData.length; i++) {
        dupColors.push([QOO10_COLOR_DUP, QOO10_COLOR_DUP, QOO10_COLOR_DUP]);
      }
      shDup.getRange(2, 1, dupData.length, dupHeader.length).setBackgrounds(dupColors);
    }
  }
}


/******************************************************
 * Qoo10 チェックボックス操作
 ******************************************************/
function Qoo10_チェック全選択() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var Y = ss.getSheetByName('Yahoo在庫ビュー');
  if (!Y) return;

  var rows = getLastDataRowByCol_(Y, 1, 3) - 2;
  if (rows <= 0) return;

  var skuVals = Y.getRange(3, QOO10_COL.SKU, rows, 1).getValues();
  var checks = [];
  for (var i = 0; i < skuVals.length; i++) {
    checks.push([String(skuVals[i][0] || '').trim() !== '']);
  }
  Y.getRange(3, QOO10_COL.CHECK, rows, 1).setValues(checks);
}

function Qoo10_チェック全解除() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var Y = ss.getSheetByName('Yahoo在庫ビュー');
  if (!Y) return;

  var rows = getLastDataRowByCol_(Y, 1, 3) - 2;
  if (rows <= 0) return;

  var falseArray = [];
  for (var i = 0; i < rows; i++) {
    falseArray.push([false]);
  }
  Y.getRange(3, QOO10_COL.CHECK, rows, 1).setValues(falseArray);
}

function Qoo10差異_全選択() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var Y = ss.getSheetByName('Yahoo在庫ビュー');
  if (!Y) return;

  var lastRow = getLastDataRowByCol_(Y, 1, 3);
  if (lastRow < 3) return;

  var rows = lastRow - 2;
  var judgeVals = Y.getRange(3, QOO10_COL.JUDGE, rows, 1).getValues();
  var checks = [];

  for (var i = 0; i < rows; i++) {
    var judge = String(judgeVals[i][0] || '').trim();
    var isDiff = judge !== '' && judge.indexOf('OK') === -1 && judge.indexOf('未登録') === -1;
    checks.push([isDiff]);
  }

  Y.getRange(3, QOO10_COL.CHECK, rows, 1).setValues(checks);
  ss.toast('Qoo10差異の行だけチェックを入れました', 'Qoo10差異だけチェックON', 5);
}

function Qoo10差異_逆選択() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var Y = ss.getSheetByName('Yahoo在庫ビュー');
  if (!Y) return;

  var lastRow = getLastDataRowByCol_(Y, 1, 3);
  if (lastRow < 3) return;

  var rows = lastRow - 2;
  var range = Y.getRange(3, QOO10_COL.CHECK, rows, 1);
  var vals = range.getValues();

  var flipped = [];
  for (var i = 0; i < vals.length; i++) {
    flipped.push([!(vals[i][0] === true)]);
  }

  range.setValues(flipped);
  ss.toast('Qoo10チェック状態を反転しました', 'Qoo10差異チェック反転', 5);
}


/******************************************************
 * 共通ヘルパー
 ******************************************************/
function getSheetLastRow_API_(ssId, sheetName) {
  try {
    var range = "'" + sheetName + "'!A:A";
    var result = Sheets.Spreadsheets.Values.get(ssId, range, { valueRenderOption: 'UNFORMATTED_VALUE' });
    return (result.values || []).length;
  } catch (e) {
    return 0;
  }
}

function getLastDataRowByCol_API_(ssId, sheetName, col, minRow) {
  try {
    var colLetter = colLetter_Q_(col);
    var range = "'" + sheetName + "'!" + colLetter + ":" + colLetter;
    var result = Sheets.Spreadsheets.Values.get(ssId, range, { valueRenderOption: 'UNFORMATTED_VALUE' });
    var values = result.values || [];
    for (var i = values.length - 1; i >= minRow - 1; i--) {
      if (values[i] && String(values[i][0] || '').trim() !== '') return i + 1;
    }
    return minRow - 1;
  } catch (e) {
    return minRow - 1;
  }
}

function getLastDataRowByCol_(sheet, col, minRow) {
  var ssId = sheet.getParent().getId();
  return getLastDataRowByCol_API_(ssId, sheet.getName(), col, minRow);
}

function colLetter_Q_(colNum) {
  var letter = '';
  while (colNum > 0) {
    var mod = (colNum - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    colNum = Math.floor((colNum - 1) / 26);
  }
  return letter;
}

function looksLikeB_(code, sub) {
  var s = String(sub || '').trim().toLowerCase();
  if (s === 'b' || s.endsWith('b')) return true;
  var c = String(code || '').trim().toLowerCase();
  if (c.endsWith('b') && !s) return true;
  return false;
}

function stripOptionSuffix_(code) {
  code = String(code || '');
  if (!code) return '';
  var last = code.slice(-1);
  if (last === 'a' || last === 'b') return code.slice(0, -1);
  return code;
}

function makeParentCode_(sku) {
  if (!sku) return '';
  var m = String(sku).match(/^(.+?)([ab])$/);
  return m ? m[1] : String(sku);
}

function buildQoo10Map_(ssId, sheet) {
  var lastRow = getSheetLastRow_API_(ssId, 'Qoo10在庫取込');
  var result = { map: new Map(), itemNoMap: new Map(), duplicates: [] };
  if (lastRow < 2) return result;

  var lastCol = sheet.getLastColumn();

  var headerRange = "'Qoo10在庫取込'!A1:" + colLetter_Q_(lastCol) + "1";
  var headers = [];
  try {
    var hResult = Sheets.Spreadsheets.Values.get(ssId, headerRange, { valueRenderOption: 'UNFORMATTED_VALUE' });
    headers = (hResult.values && hResult.values[0])
      ? hResult.values[0].map(function (v) { return String(v || '').trim(); })
      : [];
  } catch (e) {
    return result;
  }

  var colItemNo = (headers.indexOf('item_number') + 1) || 1;
  var colItemId = (headers.indexOf('seller_unique_item_id') + 1) || 2;
  var colEditType = (headers.indexOf('edit_type') + 1) || 3;
  var colOptionId = (headers.indexOf('seller_unique_option_id') + 1) || 4;
  var colQty = (headers.indexOf('quantity') + 1) || 6;

  var dataRange = "'Qoo10在庫取込'!A2:" + colLetter_Q_(lastCol) + lastRow;
  var data = [];
  try {
    var dResult = Sheets.Spreadsheets.Values.get(ssId, dataRange, { valueRenderOption: 'UNFORMATTED_VALUE' });
    data = dResult.values || [];
  } catch (e) {
    return result;
  }

  var groups = new Map();
  var seenKeys = new Map();

  for (var i = 0; i < data.length; i++) {
    var row = data[i] || [];
    var itemNo = String(row[colItemNo - 1] || '').trim();
    var itemId = String(row[colItemId - 1] || '').trim();
    var editType = String(row[colEditType - 1] || '').trim().toLowerCase();
    var optionId = String(row[colOptionId - 1] || '').trim();
    var qty = Number(row[colQty - 1]) || 0;

    if (!itemNo && !itemId) continue;

    var key = optionId || itemId;
    if (key) seenKeys.set(key, (seenKeys.get(key) || 0) + 1);

    if (!groups.has(itemNo)) {
      groups.set(itemNo, { itemId: itemId, baseQty: 0, variants: new Map(), hasVariant: false });
    }
    var g = groups.get(itemNo);
    if (!g.itemId && itemId) g.itemId = itemId;

    if (editType === 'i') {
      g.hasVariant = true;
      if (optionId) g.variants.set(optionId, qty);
    } else if (editType === 'g') {
      g.baseQty = qty;
    }
  }

  var duplicates = [];
  seenKeys.forEach(function (count, key) {
    if (count > 1) duplicates.push({ key: key, count: count });
  });

  var map = new Map();
  var itemNoMap = new Map();

  function addItemNo(sku, iNo) {
    var current = itemNoMap.get(sku);
    if (current) {
      var arr = current.split(', ');
      if (arr.indexOf(iNo) === -1) itemNoMap.set(sku, current + ', ' + iNo);
    } else {
      itemNoMap.set(sku, iNo);
    }
  }

  groups.forEach(function (g, itemNoKey) {
    if (g.hasVariant) {
      g.variants.forEach(function (qty, optCode) {
        map.set(optCode, (map.get(optCode) || 0) + qty);
        addItemNo(optCode, itemNoKey);
      });
    } else {
      var key = g.itemId || itemNoKey;
      if (key) {
        map.set(key, (map.get(key) || 0) + g.baseQty);
        addItemNo(key, itemNoKey);
      }
    }
  });

  result.map = map;
  result.itemNoMap = itemNoMap;
  result.duplicates = duplicates;
  return result;
}


/******************************************************
 * Yahoo在庫ビュー → Qoo10価格在庫変更 へ送信
 ******************************************************/
function チェック済みをQoo10価格在庫変更へ送る() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var Y = ss.getSheetByName('Yahoo在庫ビュー');
  var T = ss.getSheetByName('Qoo10価格在庫変更');
  var Q = ss.getSheetByName('Qoo10在庫取込');

  if (!Y) throw new Error('Yahoo在庫ビューが見つかりません');
  if (!T) throw new Error('Qoo10価格在庫変更が見つかりません');
  if (!Q) throw new Error('Qoo10在庫取込が見つかりません');

  var rows = getLastDataRowByCol_(Y, 1, 3) - 2;
  if (rows <= 0) {
    ss.toast('データがありません', '停止', 4);
    return;
  }

  var checkVals = Y.getRange(3, QOO10_COL.CHECK, rows, 1).getValues();
  var skuVals = Y.getRange(3, QOO10_COL.SKU, rows, 1).getValues();

  var targetSkus = [];
  for (var i = 0; i < rows; i++) {
    if (checkVals[i][0] === true) {
      var sku = String(skuVals[i][0] || '').trim();
      if (sku) targetSkus.push(sku);
    }
  }

  if (targetSkus.length === 0) {
    ss.toast('チェック済みSKUがありません', '結果なし', 4);
    return;
  }

  var maps = buildYahooQtyMapsForQoo10_(Y);
  var yahooByKey = maps.byKey;
  var yahooParent = maps.byParent;
  var detailMap = getQoo10DetailsMap_(Q);

  var outValues = [];

  for (var i = 0; i < targetSkus.length; i++) {
    var sku = targetSkus[i];
    var detail = detailMap.get(sku);
    var parent = makeParentCode_(sku);

    var qtyByKey = yahooByKey.get(sku) || 0;
    var qtyByParent = getYahooParentQtyForQoo10_(parent, yahooParent);
    var yahooQty = qtyByKey || qtyByParent || 0;

    if (detail) {
      outValues.push([
        detail.item_number,
        parent,
        detail.edit_type,
        detail.seller_unique_option_id || sku,
        '',
        yahooQty
      ]);
    } else {
      outValues.push(['', parent, '', sku, '', yahooQty]);
    }
  }

  var DST_START_ROW = 5;
  var tLast = T.getLastRow();
  if (tLast >= DST_START_ROW) {
    T.getRange(DST_START_ROW, 1, tLast - DST_START_ROW + 1, 6).clearContent();
  }

  if (outValues.length > 0) {
    T.getRange(DST_START_ROW, 1, outValues.length, 6).setValues(outValues);

    var cdr = T.getRange(DST_START_ROW, 3, outValues.length, 2);
    var cd = cdr.getValues();
    for (var i = 0; i < cd.length; i++) {
      if (String(cd[i][0] || '').trim().toLowerCase() === 'g') {
        cd[i][1] = '';
      }
    }
    cdr.setValues(cd);
  }

  ss.toast('Qoo10送信完了：' + outValues.length + '件', 'Qoo10送信', 5);
}


function buildYahooQtyMapsForQoo10_(sheet) {
  var last = getLastDataRowByCol_(sheet, 1, 3);
  var result = { byKey: new Map(), byParent: new Map() };
  if (last < 3) return result;

  var rows = last - 2;
  var vals = sheet.getRange(3, 1, rows, 6).getValues();

  for (var i = 0; i < vals.length; i++) {
    var code = String(vals[i][0] || '').trim();
    var sub = String(vals[i][2] || '').trim();
    var sokuno = Number(vals[i][3] || 0) || 0;
    var otori = Number(vals[i][4] || 0) || 0;
    var keyF = String(vals[i][5] || '').trim();
    if (!code) continue;

    var rowKey = keyF || (code + sub);
    var parent = makeParentCode_(code);

    var rowQty = looksLikeB_(code, sub) ? otori : sokuno;
    result.byKey.set(rowKey, rowQty);

    if (!result.byParent.has(parent)) {
      result.byParent.set(parent, { sokuno: 0, otori: 0 });
    }
    var agg = result.byParent.get(parent);
    agg.sokuno += sokuno;
    agg.otori += otori;
  }
  return result;
}

function getYahooParentQtyForQoo10_(parent, parentAgg) {
  var rec = parentAgg.get(parent);
  if (!rec) return 0;
  if (rec.otori > 0) return rec.otori;
  if (rec.sokuno > 0) return rec.sokuno;
  return 0;
}

function getQoo10DetailsMap_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return new Map();

  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var headersArr = headers.map(function(h) { return String(h || '').trim(); });

  var colItemNum = (headersArr.indexOf('item_number') + 1) || 1;
  var colItemId = (headersArr.indexOf('seller_unique_item_id') + 1) || 2;
  var colEditType = (headersArr.indexOf('edit_type') + 1) || 3;
  var colOptionId = (headersArr.indexOf('seller_unique_option_id') + 1) || 4;
  var colQty = (headersArr.indexOf('quantity') + 1) || 6;

  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  var tmp = new Map();

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var itemNum = String(row[colItemNum - 1] || '');
    var itemId = String(row[colItemId - 1] || '').trim();
    var editType = String(row[colEditType - 1] || '').trim();
    var optionId = String(row[colOptionId - 1] || '').trim();
    var qty = Number(row[colQty - 1]) || 0;

    var key = optionId || itemId;
    if (!key) continue;

    if (!tmp.has(key)) {
      tmp.set(key, { iDetail: null, qtyI: 0, gDetail: null, qtyG: 0 });
    }
    var rec = tmp.get(key);

    var detail = {
      item_number: itemNum,
      seller_unique_item_id: itemId,
      edit_type: editType,
      seller_unique_option_id: optionId,
      quantity: 0
    };

    if (editType.toLowerCase() === 'i') {
      if (!rec.iDetail) rec.iDetail = detail;
      rec.qtyI += qty;
    } else {
      if (!rec.gDetail) rec.gDetail = detail;
      rec.qtyG += qty;
    }
  }

  var map = new Map();
  tmp.forEach(function(rec, k) {
    if (rec.iDetail) {
      var d = Object.assign({}, rec.iDetail);
      d.quantity = rec.qtyI;
      map.set(k, d);
    } else if (rec.gDetail) {
      var d2 = Object.assign({}, rec.gDetail);
      d2.quantity = rec.qtyG;
      map.set(k, d2);
    }
  });

  return map;
}


/******************************************************
 * その他の関数（既存互換）
 ******************************************************/
function Qoo10未登録_親コード付与() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('未登録一覧_Qoo10（要登録）');
  if (!sh) {
    ss.toast('未登録一覧_Qoo10シートが見つかりません', '親コード付与', 5);
    return;
  }

  var lastRow = getLastDataRowByCol_(sh, 1, 2);
  if (lastRow < 2) {
    ss.toast('データがありません', '親コード付与', 5);
    return;
  }

  var numRows = lastRow - 1;
  var skuVals = sh.getRange(2, 1, numRows, 1).getValues();

  var out = [];
  for (var i = 0; i < skuVals.length; i++) {
    out.push([makeParentCode_(skuVals[i][0])]);
  }

  sh.getRange(2, 5, numRows, 1).setValues(out);
  ss.toast('E列に親コードを付与しました', '親コード付与', 5);
}

function D列から値を検索_Qoo10価格在庫変更() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shQ = ss.getSheetByName('Qoo10価格在庫変更');
  var shSrc = ss.getSheetByName('Qoo10在庫取込');
  if (!shQ || !shSrc) throw new Error('シートが見つかりません');

  var START_ROW = 5;
  var lastRow = shQ.getLastRow();
  if (lastRow < START_ROW) return;

  var numRows = lastRow - START_ROW + 1;
  var range = shQ.getRange(START_ROW, 1, numRows, 6);
  var values = range.getValues();

  var detailMap = getQoo10DetailsMap_(shSrc);
  var stockMap = buildYahooStockMap_();

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var optionCode = String(row[3] || '').trim();
    if (!optionCode) continue;

    var parentCode = String(row[1] || '').trim();
    if (!parentCode) {
      var m = optionCode.match(/^(.+?)([ab])$/);
      parentCode = m ? m[1] : optionCode;
      row[1] = parentCode;
    }

    var d = detailMap.get(optionCode) || detailMap.get(parentCode);
    if (d) {
      if (!row[0]) row[0] = d.item_number || '';
      if (!row[1]) row[1] = d.seller_unique_item_id || parentCode;
      if (!row[2]) row[2] = d.edit_type || 'i';
    } else {
      if (!row[2]) row[2] = 'i';
    }

    if (String(row[2] || '').trim().toLowerCase() === 'g') {
      row[3] = '';
    }

    var optKey = String(row[3] || '').trim();
    var parentKey = String(row[1] || '').trim();
    row[5] = stockMap[optKey] || stockMap[parentKey] || 0;
  }

  range.setValues(values);
  ss.toast('D列から値を検索完了', 'Qoo10価格在庫変更', 5);
}

function buildYahooStockMap_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shY = ss.getSheetByName('Yahoo在庫ビュー');
  if (!shY) return {};

  var lastRow = shY.getLastRow();
  if (lastRow < 3) return {};

  var values = shY.getRange(3, 1, lastRow - 2, 6).getValues();
  var map = {};

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var parent = String(row[0] || '').trim();
    var subFull = String(row[2] || '').trim();
    var unionKey = String(row[5] || '').trim();

    if (!parent && !subFull && !unionKey) continue;

    var stock = (Number(row[3] || 0)) + (Number(row[4] || 0));
    var fullCode = unionKey || subFull || parent;
    map[fullCode] = stock;
  }

  return map;
}

function Qoo10価格在庫変更_クリア_AF() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var T = ss.getSheetByName('Qoo10価格在庫変更');
  if (!T) return;

  var lastRow = T.getLastRow();
  if (lastRow < 5) return;

  T.getRange(5, 1, lastRow - 4, 6).clearContent();
  ss.toast('クリアしました', 'クリア', 4);
}
