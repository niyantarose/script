// ===== デバッグ設定 =====
var AMZ_DEBUG = true;
var AMZ_DEBUG_SAMPLE = 5;

function amzLog_(label, obj) {
  if (!AMZ_DEBUG) return;
  var msg = "[Amazon在庫照合][" + label + "]";
  if (typeof obj !== "undefined") msg += " " + safeStringify_(obj, 2000);
  console.log(msg);
}

function safeStringify_(obj, maxLen) {
  try {
    var s = JSON.stringify(obj);
    if (s.length > maxLen) s = s.slice(0, maxLen) + "...(" + s.length + " chars)";
    return s;
  } catch (e) {
    return String(obj);
  }
}

function amzMs_(t0) {
  return (new Date().getTime() - t0) + "ms";
}

/******************************************************
 * Amazon在庫照合：共通定数
 ******************************************************/
var AMZ_COL = globalThis.AMZ_COL || (globalThis.AMZ_COL = {
  CHECK: 8,  // H列
  SKU:   9,  // I列
  STOCK: 10, // J列
  MODE:  11, // K列
  DIFF:  12, // L列
  JUDGE: 13  // M列
});

var AMZ_COLOR_NONE   = (globalThis.AMZ_COLOR_NONE   ?? (globalThis.AMZ_COLOR_NONE   = null));
var AMZ_COLOR_LESS   = (globalThis.AMZ_COLOR_LESS   ?? (globalThis.AMZ_COLOR_LESS   = "#F4C7C3"));
var AMZ_COLOR_MORE   = (globalThis.AMZ_COLOR_MORE   ?? (globalThis.AMZ_COLOR_MORE   = "#C9DAF8"));
var AMZ_COLOR_WARN   = (globalThis.AMZ_COLOR_WARN   ?? (globalThis.AMZ_COLOR_WARN   = "#FFE699"));
var AMZ_COLOR_SOKUNO = (globalThis.AMZ_COLOR_SOKUNO ?? (globalThis.AMZ_COLOR_SOKUNO = "#D9EAD3"));
var AMZ_COLOR_OTORI  = (globalThis.AMZ_COLOR_OTORI  ?? (globalThis.AMZ_COLOR_OTORI  = "#FFF2CC"));


function Amazon在庫照合() {
  var t0 = new Date().getTime();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssId = ss.getId();

  amzLog_("START", { time: new Date().toISOString() });

  var shY = ss.getSheetByName("Yahoo在庫ビュー");
  var shA = ss.getSheetByName("Amazon在庫取込");
  if (!shY) throw new Error("Yahoo在庫ビュー シートが見つかりません");
  if (!shA) throw new Error("Amazon在庫取込 シートが見つかりません");

  var HEADER_ROW = 2;
  var DATA_START_ROW = 3;
  var shYId = shY.getSheetId();

  // --- 1) Yahoo側行数 ---
  var yLast = getLastDataRowByCol_AMZ_(shY, 1, DATA_START_ROW);
  amzLog_("Yahoo lastRow", { yLast: yLast, DATA_START_ROW: DATA_START_ROW });

  if (yLast < DATA_START_ROW) {
    ss.toast("Yahoo在庫ビューにデータがありません", "Amazon在庫照合", 5);
    return;
  }
  var dataRows = yLast - HEADER_ROW;

  ss.toast("Amazon在庫照合を開始します…", "Amazon在庫照合", 10);

  // --- 2) Amazonエリアリセット（Sheets API）---
  var tReset = new Date().getTime();
  var clearRange = "'Yahoo在庫ビュー'!H" + DATA_START_ROW + ":M" + yLast;
  Sheets.Spreadsheets.Values.clear({}, ssId, clearRange);
  
  // 背景色リセット（batchUpdate）- ヘッダー行(2行目)は除外、データ行(3行目〜)のみ
  Sheets.Spreadsheets.batchUpdate({
    requests: [{
      repeatCell: {
        range: {
          sheetId: shYId,
          startRowIndex: DATA_START_ROW - 1,  // 3行目 = index 2
          endRowIndex: yLast,
          startColumnIndex: AMZ_COL.CHECK - 1,
          endColumnIndex: AMZ_COL.JUDGE + 1  // exclusive なので +1 必要
        },
        cell: { userEnteredFormat: { backgroundColor: { red: 1, green: 1, blue: 1 } } },
        fields: 'userEnteredFormat.backgroundColor'
      }
    }]
  }, ssId);

  // ヘッダー
  shY.getRange(HEADER_ROW, AMZ_COL.CHECK, 1, 6)
    .setValues([["反映Amazon", "AmazonSKU", "Amazon在庫", "発送区分", "差分(A-Y)", "判定(Amazon)"]])
    .setFontWeight("bold");

  // チェックボックス
  var cbRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  shY.getRange(DATA_START_ROW, AMZ_COL.CHECK, dataRows, 1).setDataValidation(cbRule);
  amzLog_("Reset done", { took: amzMs_(tReset) });

  // --- 3) Amazon在庫取込 読み込み ---
  var aMap = new Map();
  var aLast = getLastDataRowByCol_AMZ_(shA, 1, 2);

  if (aLast >= 2) {
    var tA = new Date().getTime();
    var aRange = "'Amazon在庫取込'!A2:D" + aLast;
    var aResult = Sheets.Spreadsheets.Values.get(ssId, aRange, { valueRenderOption: "UNFORMATTED_VALUE" });
    var aVals = aResult.values || [];
    amzLog_("Amazon read", { rows: aVals.length, took: amzMs_(tA) });

    for (var i = 0; i < aVals.length; i++) {
      var rawSku = String(aVals[i][0] || "").trim();
      var qty = Number(aVals[i][3]) || 0;
      if (rawSku) {
        var sku = normalizeAmazonSku_AMZ_(rawSku);
        aMap.set(sku, (aMap.get(sku) || 0) + qty);
      }
    }
    amzLog_("Amazon map built", { size: aMap.size });
  }

  // --- 4) Yahoo在庫ビュー A〜G を読み込み ---
  var tY = new Date().getTime();
  var yRange = "'Yahoo在庫ビュー'!A" + DATA_START_ROW + ":G" + yLast;
  var yResult = Sheets.Spreadsheets.Values.get(ssId, yRange, { valueRenderOption: "UNFORMATTED_VALUE" });
  var yVals = yResult.values || [];
  amzLog_("Yahoo read", { rows: yVals.length, took: amzMs_(tY) });

  // --- 5) 統合キー単位でグループ化 ---
  var groups = new Map();

  for (var r = 0; r < yVals.length; r++) {
    var row = yVals[r] || [];
    var parentCode = String(row[0] || "").trim();
    var subCode = String(row[2] || "").trim();
    var soku = Number(row[3] || 0) || 0;
    var otori = Number(row[4] || 0) || 0;
    var uni = String(row[5] || "").trim();

    var baseKey = "";
    if (uni) baseKey = stripLowerAB_(normalizeAmazonSku_AMZ_(uni));
    else if (parentCode) baseKey = normalizeAmazonSku_AMZ_(parentCode);
    if (!baseKey) continue;

    var isARow = hasSmallA_(subCode);
    var isBRow = hasSmallB_(subCode);

    if (!groups.has(baseKey)) {
      groups.set(baseKey, { rows: [], totalSoku: 0, totalOtori: 0 });
    }
    var g = groups.get(baseKey);
    g.rows.push({ index: r, subCode: subCode, soku: soku, otori: otori, isARow: isARow, isBRow: isBRow });
    g.totalSoku += soku;
    g.totalOtori += otori;
  }

  amzLog_("Groups", { count: groups.size });

  // --- 6) スパース出力データ作成（データがある行だけ）---
  var sparseData = [];  // { row: 実際の行番号, values: [5列], color: 色 }
  var diffRows = [];
  var unregRows = [];

  var stat = { hasAmz: 0, unreg: 0, ok: 0, more: 0, less: 0, warnAmazonOnly: 0 };

  groups.forEach(function (group, baseKey) {
    var totalSoku = group.totalSoku;
    var totalOtori = group.totalOtori;
    var amazonSku = baseKey;
    if (!amazonSku) return;

    var amzQty = aMap.has(amazonSku) ? aMap.get(amazonSku) : 0;
    var hasAmz = aMap.has(amazonSku);

    var displayRowIndex = -1;
    var yahooMode = "";
    var yahooQty = 0;

    if (totalSoku > 0) {
      yahooMode = "即納";
      yahooQty = totalSoku;
      for (var k = 0; k < group.rows.length; k++) if (group.rows[k].soku > 0) { displayRowIndex = group.rows[k].index; break; }
      if (displayRowIndex === -1) for (var k = 0; k < group.rows.length; k++) if (group.rows[k].isARow) { displayRowIndex = group.rows[k].index; break; }
      if (displayRowIndex === -1 && group.rows.length > 0) displayRowIndex = group.rows[0].index;
    } else if (totalOtori > 0) {
      yahooMode = "お取り寄せ";
      yahooQty = totalOtori;
      for (var k = 0; k < group.rows.length; k++) if (group.rows[k].otori > 0) { displayRowIndex = group.rows[k].index; break; }
      if (displayRowIndex === -1) for (var k = 0; k < group.rows.length; k++) if (group.rows[k].isBRow) { displayRowIndex = group.rows[k].index; break; }
      if (displayRowIndex === -1 && group.rows.length > 0) displayRowIndex = group.rows[0].index;
    } else if (amzQty > 0) {
      for (var k = 0; k < group.rows.length; k++) if (group.rows[k].isARow) { displayRowIndex = group.rows[k].index; break; }
      if (displayRowIndex === -1 && group.rows.length > 0) displayRowIndex = group.rows[0].index;
    }

    if (totalSoku === 0 && totalOtori === 0 && amzQty === 0) return;
    if (displayRowIndex === -1) return;

    // displayRowIndex に対応する行の実際の在庫を取得
    var displayRow = null;
    for (var k = 0; k < group.rows.length; k++) {
      if (group.rows[k].index === displayRowIndex) {
        displayRow = group.rows[k];
        break;
      }
    }
    var rowSoku = displayRow ? displayRow.soku : 0;
    var rowOtori = displayRow ? displayRow.otori : 0;
    
    // ★この行自体の在庫フラグ
    var thisRowHasStock = (rowSoku > 0 || rowOtori > 0);

    var rowData = [amazonSku, "", yahooMode, "", ""];
    var rowColor = null;

    if (!hasAmz) {
      rowData[4] = "Amazon未登録";
      // この行自体に在庫がある場合のみ色を付ける
      if (thisRowHasStock) {
        rowColor = AMZ_COLOR_WARN;
      }
      unregRows.push([amazonSku, yahooQty, yahooMode, "Amazon未登録"]);
      stat.unreg++;
    } else {
      stat.hasAmz++;
      rowData[1] = amzQty;
      var diff = amzQty - yahooQty;
      rowData[3] = diff;

      if (yahooQty === 0 && amzQty > 0) {
        rowData[4] = "Amazonだけ在庫あり";
        // Yahoo=0なので色なし
        rowColor = null;
        diffRows.push([amazonSku, amzQty, yahooQty, yahooMode, "Amazonだけ在庫あり"]);
        stat.warnAmazonOnly++;
      } else if (diff === 0) {
        rowData[4] = "OK（" + yahooMode + "：" + yahooQty + "）";
        stat.ok++;
      } else if (diff > 0) {
        rowData[4] = "Amazonが多い";
        // この行自体に在庫がある場合のみ色を付ける
        if (thisRowHasStock) {
          rowColor = AMZ_COLOR_MORE;
        }
        diffRows.push([amazonSku, amzQty, yahooQty, yahooMode, "Amazonが多い"]);
        stat.more++;
      } else {
        rowData[4] = "Amazonが少ない";
        // この行自体に在庫がある場合のみ色を付ける
        if (thisRowHasStock) {
          rowColor = AMZ_COLOR_LESS;
        }
        diffRows.push([amazonSku, amzQty, yahooQty, yahooMode, "Amazonが少ない"]);
        stat.less++;
      }
    }
    
    // ★デバッグ用：在庫0なのに色が付く場合のログ
    if (rowColor !== null && !thisRowHasStock) {
      amzLog_("警告: 在庫0なのに色付け", {
        row: DATA_START_ROW + displayRowIndex,
        sku: amazonSku,
        rowSoku: rowSoku,
        rowOtori: rowOtori,
        totalSoku: totalSoku,
        totalOtori: totalOtori
      });
    }

    sparseData.push({
      row: DATA_START_ROW + displayRowIndex,
      values: rowData,
      color: rowColor
    });
  });

  amzLog_("Sparse data", { count: sparseData.length, stat: stat });

  // --- 7) スパース書き込み（Sheets API batchUpdate）---
  var tW = new Date().getTime();
  
  if (sparseData.length > 0) {
    // 連続する行をグループ化して効率的に書き込み
    sparseData.sort(function(a, b) { return a.row - b.row; });
    
    var valueRequests = [];
    var colorRequests = [];
    
    var i = 0;
    while (i < sparseData.length) {
      var startIdx = i;
      var startRow = sparseData[i].row;
      
      // 連続する行を探す
      while (i < sparseData.length - 1 && sparseData[i + 1].row === sparseData[i].row + 1) {
        i++;
      }
      
      var endIdx = i;
      var numRows = endIdx - startIdx + 1;
      
      // 値の配列を作成
      var values = [];
      for (var j = startIdx; j <= endIdx; j++) {
        values.push(sparseData[j].values);
      }
      
      // 値書き込みリクエスト
      var rangeStr = "'Yahoo在庫ビュー'!I" + startRow + ":M" + (startRow + numRows - 1);
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
    
    // 色付け（色がある行だけ）
    var coloredRows = sparseData.filter(function(d) { return d.color !== null; });
    
    if (coloredRows.length > 0) {
      var colorBatchRequests = [];
      
      // 色ごとにグループ化
      var colorGroups = {};
      for (var k = 0; k < coloredRows.length; k++) {
        var c = coloredRows[k].color;
        if (!colorGroups[c]) colorGroups[c] = [];
        colorGroups[c].push(coloredRows[k].row);
      }
      
      // 各色の連続行をマージしてリクエスト作成
      for (var color in colorGroups) {
        var rows = colorGroups[color].sort(function(a, b) { return a - b; });
        var ranges = mergeConsecutiveRows_AMZ_(rows);
        
        var rgb = hexToRgb_(color);
        
        for (var m = 0; m < ranges.length; m++) {
          colorBatchRequests.push({
            repeatCell: {
              range: {
                sheetId: shYId,
                startRowIndex: ranges[m].start - 1,
                endRowIndex: ranges[m].end - 1,
                startColumnIndex: AMZ_COL.CHECK - 1,
                endColumnIndex: AMZ_COL.JUDGE + 1  // exclusive なので +1 必要
              },
              cell: {
                userEnteredFormat: {
                  backgroundColor: rgb
                }
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
  
  amzLog_("Write done", { took: amzMs_(tW), sparseRows: sparseData.length });

  // --- 8) レポート更新 ---
  var tR = new Date().getTime();
  Amazonレポートシート更新_(diffRows, unregRows);
  amzLog_("Report done", { took: amzMs_(tR) });

  SpreadsheetApp.flush();
  amzLog_("END", { total: amzMs_(t0) });
  ss.toast("Amazon在庫照合 完了", "Amazon在庫照合", 5);
}


/**
 * 連続する行番号をマージ
 */
function mergeConsecutiveRows_AMZ_(rows) {
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
function hexToRgb_(hex) {
  var result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  return result ? {
    red: parseInt(result[1], 16) / 255,
    green: parseInt(result[2], 16) / 255,
    blue: parseInt(result[3], 16) / 255
  } : { red: 1, green: 1, blue: 1 };
}


/******************************************************
 * チェックボックス操作
 ******************************************************/
function Amazon_チェック全選択() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shY = ss.getSheetByName("Yahoo在庫ビュー");
  var rows = getLastDataRowByCol_AMZ_(shY, 1, 3) - 2;
  if (rows <= 0) return;

  var v = shY.getRange(3, AMZ_COL.SKU, rows, 1).getValues();
  shY.getRange(3, AMZ_COL.CHECK, rows, 1).setValues(v.map(function (r) {
    return [!!r[0]];
  }));
}

function Amazon_チェック全解除() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shY = ss.getSheetByName("Yahoo在庫ビュー");
  var rows = getLastDataRowByCol_AMZ_(shY, 1, 3) - 2;
  if (rows <= 0) return;

  shY.getRange(3, AMZ_COL.CHECK, rows, 1).setValues(new Array(rows).fill([false]));
}

function Amazon差異_全選択() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shY = ss.getSheetByName("Yahoo在庫ビュー");
  var rows = getLastDataRowByCol_AMZ_(shY, 1, 3) - 2;
  if (rows <= 0) return;

  var v = shY.getRange(3, AMZ_COL.JUDGE, rows, 1).getValues();
  shY.getRange(3, AMZ_COL.CHECK, rows, 1).setValues(v.map(function (r) {
    var j = String(r[0] || "");
    return [j !== "" && j.indexOf("OK") === -1 && j.indexOf("未登録") === -1];
  }));
  ss.toast("完了");
}

function Amazon差異_逆選択() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shY = ss.getSheetByName("Yahoo在庫ビュー");
  var rows = getLastDataRowByCol_AMZ_(shY, 1, 3) - 2;
  if (rows <= 0) return;

  var rng = shY.getRange(3, AMZ_COL.CHECK, rows, 1);
  rng.setValues(rng.getValues().map(function (r) {
    return [!r[0]];
  }));
  ss.toast("完了");
}


/******************************************************
 * ヘルパー関数群
 ******************************************************/
function getLastDataRowByCol_AMZ_(sheet, col, minRow) {
  var lastRow = sheet.getLastRow();
  if (lastRow < minRow) return minRow - 1;

  var vals = sheet.getRange(minRow, col, lastRow - minRow + 1, 1).getValues();
  for (var i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0] || "").trim() !== "") return minRow + i;
  }
  return minRow - 1;
}

function normalizeAmazonSku_AMZ_(skuRaw) {
  var s = String(skuRaw || "").trim();
  if (!s) return "";
  s = s.replace(/\s+/g, "");
  s = s.replace(/[-_]?AMZ$/i, "");
  return s;
}

function stripLowerAB_(s) {
  s = String(s || "").trim();
  return s.replace(/[ab]$/, "");
}

function hasSmallA_(subCode) {
  return /a$/.test(String(subCode || "").trim());
}

function hasSmallB_(subCode) {
  return /b$/.test(String(subCode || "").trim());
}

function columnToLetter_AMZ_(column) {
  var temp, letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}


/******************************************************
 * レポート更新
 ******************************************************/
function Amazonレポートシート更新_(diffRows, unregRows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ===== 1) 差異一覧 =====
  var shDiff = ss.getSheetByName("Amazon在庫差異一覧");
  if (!shDiff) shDiff = ss.insertSheet("Amazon在庫差異一覧");

  var diffHeader = ["AmazonSKU", "Amazon在庫", "Yahoo在庫", "Yahoo区分", "メモ"];

  shDiff.getRange(1, 1, 1, diffHeader.length)
    .clearContent()
    .setValues([diffHeader])
    .setFontWeight("bold");

  var lastRowDiff = shDiff.getLastRow();
  if (lastRowDiff > 1) {
    shDiff.getRange(2, 1, lastRowDiff - 1, shDiff.getLastColumn() || diffHeader.length)
      .clearContent()
      .setBackground(null);
  }

  if (diffRows && diffRows.length > 0) {
    shDiff.getRange(2, 1, diffRows.length, diffHeader.length).setValues(diffRows);

    var rowColors = [];
    for (var r = 0; r < diffRows.length; r++) {
      var amz = Number(diffRows[r][1] || 0);
      var yaho = Number(diffRows[r][2] || 0);
      var memo = String(diffRows[r][4] || "");
      var c = null;

      if (memo.indexOf("Amazonだけ在庫あり") !== -1) c = AMZ_COLOR_WARN;
      else if (memo.indexOf("多い") !== -1) c = AMZ_COLOR_MORE;
      else if (memo.indexOf("少ない") !== -1) c = AMZ_COLOR_LESS;

      rowColors.push(new Array(diffHeader.length).fill(c));
    }
    shDiff.getRange(2, 1, diffRows.length, diffHeader.length).setBackgrounds(rowColors);
  }

  // ===== 2) 未登録一覧 =====
  var shUnreg = ss.getSheetByName("未登録一覧_Amazon（要登録）");
  if (!shUnreg) shUnreg = ss.insertSheet("未登録一覧_Amazon（要登録）");

  var unregHeader = ["AmazonSKU", "Yahoo現在庫", "Yahoo状態", "メモ"];

  shUnreg.getRange(1, 1, 1, unregHeader.length)
    .clearContent()
    .setValues([unregHeader])
    .setFontWeight("bold");

  var lastRowUnreg = shUnreg.getLastRow();
  if (lastRowUnreg > 1) {
    shUnreg.getRange(2, 1, lastRowUnreg - 1, 6)
      .clearContent()
      .setBackground(null);
  }

  if (unregRows && unregRows.length > 0) {
    shUnreg.getRange(2, 1, unregRows.length, unregHeader.length).setValues(unregRows);

    var totalCols = 6;
    var rowColors2 = [];

    for (var i2 = 0; i2 < unregRows.length; i2++) {
      var mode2 = String(unregRows[i2][2] || "").trim();
      var c2 = AMZ_COLOR_WARN;
      if (mode2 === "即納") c2 = AMZ_COLOR_SOKUNO;
      else if (mode2 === "お取り寄せ") c2 = AMZ_COLOR_OTORI;

      rowColors2.push(new Array(totalCols).fill(c2));
    }
    shUnreg.getRange(2, 1, unregRows.length, totalCols).setBackgrounds(rowColors2);
  }
}


/******************************************************
 * Amazon未登録_子SKUをE列へ
 ******************************************************/
function Amazon未登録_子SKUをE列へ() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('未登録一覧_Amazon（要登録）');
  if (!sh) {
    ss.toast('未登録一覧_Amazon（要登録）が見つかりません', 'Amazon未登録', 5);
    return;
  }

  var HEADER = 1;
  var startRow = HEADER + 1;
  var lastRow = sh.getLastRow();
  if (lastRow < startRow) {
    ss.toast('データがありません', 'Amazon未登録', 5);
    return;
  }

  var numRows = lastRow - HEADER;
  var vals = sh.getRange(startRow, 1, numRows, 3).getValues();

  var outE = [];
  for (var i = 0; i < numRows; i++) {
    var amazonSku = String(vals[i][0] || '').trim();
    var mode = String(vals[i][2] || '').trim();

    var child = '';
    if (amazonSku) {
      if (mode === '即納') child = amazonSku + 'a';
      else if (mode === 'お取り寄せ') child = amazonSku + 'b';
    }
    outE.push([child]);
  }

  sh.getRange(startRow, 5, numRows, 1).setValues(outE);
  ss.toast('E列に子SKUを反映しました', 'Amazon未登録', 5);
}


/******************************************************
 * Amazon在庫変更 用の出力先シート名
 ******************************************************/
var AMZ_DEST_SKU_SHEET = 'Amazon価格在庫変更';


/******************************************************
 * SKU送信
 ******************************************************/
function Amazon_SKU送信() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shY = ss.getSheetByName('Yahoo在庫ビュー');
  var rows = getLastDataRowByCol_AMZ_(shY, 1, 3) - 2;
  if (rows <= 0) return;

  var vals = shY.getRange(3, 1, rows, AMZ_COL.MODE).getValues();
  var out = [];

  for (var i = 0; i < vals.length; i++) {
    if (vals[i][AMZ_COL.CHECK - 1] === true && vals[i][AMZ_COL.SKU - 1]) {
      var sku = vals[i][AMZ_COL.SKU - 1];
      var mode = vals[i][AMZ_COL.MODE - 1];
      var soku = Number(vals[i][3]) || 0;
      var otori = Number(vals[i][4]) || 0;
      var q = (mode === '即納') ? soku : otori;
      out.push([sku, '', q]);
    }
  }

  if (out.length === 0) {
    ss.toast('なし');
    return;
  }

  var shD = ss.getSheetByName(AMZ_DEST_SKU_SHEET);
  if (!shD) shD = ss.insertSheet(AMZ_DEST_SKU_SHEET);

  var dL = shD.getLastRow();
  if (dL > 3) shD.getRange(4, 1, dL - 3, 3).clearContent();

  shD.getRange(4, 1, out.length, 3).setValues(out);
  ss.toast('送信完了');
}


/******************************************************
 * 親コード取得ヘルパー
 ******************************************************/
function getYahooParentCodeForAmazon_AMZ_(code, sub) {
  var c = String(code || '').trim();
  var s = String(sub || '').trim();
  if (!c) return '';
  if (s) return c;
  var m = c.match(/^(.+?)([ab])$/i);
  return m ? m[1] : c;
}


/******************************************************
 * AmazonSKUからYahoo在庫数反映_C列
 ******************************************************/
function AmazonSKUからYahoo在庫数反映_C列() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shY = ss.getSheetByName('Yahoo在庫ビュー');
  var shD = ss.getSheetByName(AMZ_DEST_SKU_SHEET);
  if (!shY || !shD) return;

  var DATA = 3;
  var last = getLastDataRowByCol_AMZ_(shY, 1, DATA);
  if (last < DATA) return;

  var rowCount = last - DATA + 1;
  var vals = shY.getRange(DATA, 1, rowCount, 6).getValues();
  var map = new Map();

  for (var i = 0; i < vals.length; i++) {
    var code = String(vals[i][0] || '').trim();
    var sub = String(vals[i][2] || '').trim();
    var soku = Number(vals[i][3] || 0) || 0;
    var otori = Number(vals[i][4] || 0) || 0;
    var uni = String(vals[i][5] || '').trim();

    if (!code && !uni) continue;

    var key = '';
    if (uni) {
      key = normalizeAmazonSku_AMZ_(uni).replace(/(a|b)$/, '');
    } else {
      key = getYahooParentCodeForAmazon_AMZ_(code, sub);
      key = normalizeAmazonSku_AMZ_(key);
    }

    if (!key) continue;

    var ag = map.get(key);
    if (!ag) {
      ag = { s: 0, o: 0 };
      map.set(key, ag);
    }
    ag.s += soku;
    ag.o += otori;
  }

  var H = 3;
  var dL = shD.getLastRow();
  if (dL <= H) return;

  var skuVals = shD.getRange(H + 1, 1, dL - H, 1).getValues();
  var out = new Array(skuVals.length);

  for (var j = 0; j < skuVals.length; j++) {
    var rawSku = String(skuVals[j][0] || '').trim();
    var q = 0;
    if (rawSku) {
      var k = normalizeAmazonSku_AMZ_(rawSku);
      var ag2 = map.get(k);

      if (ag2) {
        var hasA = ag2.s > 0;
        var hasB = ag2.o > 0;
        if (hasA && !hasB) q = ag2.s;
        else if (!hasA && hasB) q = ag2.o;
        else if (hasA && hasB) q = ag2.s;
        else q = 0;
      }
    }
    out[j] = [q];
  }

  shD.getRange(H + 1, 3, out.length, 1).setValues(out);
  ss.toast('Yahoo在庫数 反映完了');
}


/******************************************************
 * AmazonSKUから発送区分反映_D列
 ******************************************************/
function AmazonSKUから発送区分反映_D列() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shY = ss.getSheetByName('Yahoo在庫ビュー');
  var shD = ss.getSheetByName(AMZ_DEST_SKU_SHEET);
  if (!shY || !shD) return;

  var DATA = 3;
  var last = getLastDataRowByCol_AMZ_(shY, 1, DATA);
  if (last < DATA) return;

  var vals = shY.getRange(DATA, 1, last - 2, 5).getValues();
  var map = new Map();

  for (var i = 0; i < vals.length; i++) {
    var c = String(vals[i][0] || '').trim();
    if (!c) continue;
    var p = getYahooParentCodeForAmazon_AMZ_(c, String(vals[i][2]));
    if (!map.has(p)) map.set(p, { s: 0, o: 0 });
    var ag = map.get(p);
    ag.s += (Number(vals[i][3]) || 0);
    ag.o += (Number(vals[i][4]) || 0);
  }

  var H = 3;
  var dL = shD.getLastRow();
  if (dL <= H) return;

  var sVs = shD.getRange(H + 1, 1, dL - H, 1).getValues();
  var o = [];

  for (var i = 0; i < sVs.length; i++) {
    var k = normalizeAmazonSku_AMZ_(sVs[i][0]);
    var m = '在庫なし';
    if (k) {
      var ag = map.get(k);
      if (ag) {
        if (ag.s > 0) m = '即納';
        else if (ag.o > 0) m = 'お取り寄せ';
      }
    }
    o.push([m]);
  }

  shD.getRange(H + 1, 4, o.length, 1).setValues(o);
  ss.toast('発送区分 反映完了');
}


/******************************************************
 * Amazon発送区分クリア
 ******************************************************/
function Amazon発送区分クリア() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(AMZ_DEST_SKU_SHEET);
  if (sh) {
    var H = 3;
    var L = sh.getLastRow();
    if (L > H) sh.getRange(H + 1, 4, L - H, 1).clearContent();
  }
  ss.toast('クリア完了');
}


/******************************************************
 * Amazon送信先_クリア
 ******************************************************/
function Amazon送信先_クリア() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(AMZ_DEST_SKU_SHEET);
  if (sh) {
    var H = 3;
    var L = sh.getLastRow();
    if (L > H) sh.getRange(H + 1, 1, L - H, sh.getLastColumn()).clearContent();
  }
  ss.toast('クリア完了');
}


/******************************************************
 * TSVダウンロード
 ******************************************************/
function downloadAmazonSheetAsTsv() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Amazon価格在庫変更');
  if (!sh) {
    ss.toast('Amazon価格在庫変更シートが見つかりません', 'TSV出力', 5);
    return;
  }
  var startRow = 4;
  var lastRow = sh.getLastRow();
  if (lastRow < startRow) {
    ss.toast('データがありません', 'TSV出力', 5);
    return;
  }
  var rows = lastRow - startRow + 1;
  var vals = sh.getRange(startRow, 1, rows, 3).getValues();

  var tsv = 'sku\tprice\tquantity\n';
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    if (r[0]) {
      tsv += r[0] + '\t' + (r[1] || '') + '\t' + (r[2] || '') + '\n';
    }
  }

  var fileName = 'amazon_update_' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyyMMdd_HHmmss") + '.tsv';

  var html = HtmlService.createHtmlOutput(
    '<html><head><script>' +
    'function download() {' +
    '  var data = ' + JSON.stringify(tsv) + ';' +
    '  var fileName = ' + JSON.stringify(fileName) + ';' +
    '  var blob = new Blob([data], {type: "text/tab-separated-values"});' +
    '  var a = document.createElement("a");' +
    '  a.href = URL.createObjectURL(blob);' +
    '  a.download = fileName;' +
    '  document.body.appendChild(a);' +
    '  a.click();' +
    '  document.body.removeChild(a);' +
    '  google.script.host.close();' +
    '}' +
    '</script></head>' +
    '<body onload="download()">' +
    '<p>ダウンロード中...</p>' +
    '</body></html>'
  )
  .setWidth(200)
  .setHeight(80);

  SpreadsheetApp.getUi().showModalDialog(html, 'TSVダウンロード');
}
