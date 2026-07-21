// ============================================================
// 台湾CN → Yahoo商品登録テンプレ 送信
// ============================================================

var YAHOO_TRANSFER_CFG = {
  destId: '1jN-3aLsLg_wRjLQU-mDqaImVVBgdpvwJWsHegxSJRF0',
  destSheet: '①商品入力シート',
  destHeaderRows: 2,        // 1〜2行目がヘッダー
  destDataStartRow: 3,      // データは3行目から
  srcSheets: ['台湾まんが', '台湾書籍その他', '台湾グッズ', '台湾雑誌'],
  checkboxCol: 1,           // A列
  // Yahoo列名 ← ソース候補名（左優先）
  fieldCandidates: {
    '商品名': ['タイトル', '商品名（出品用）', '雑誌名'],
    '商品コード': ['親コード', '商品コード（SKU）'],
    '発売日': ['発売日'],
    'JANコード': ['ISBN'],
    '配送グループ管理番号': ['配送パターン'],
    '商品情報': ['商品説明', '商品情報'],
    '値段': ['売価'],
  },
};

function yt_ヘッダー正規化_(s) {
  return String(s == null ? '' : s)
    .normalize('NFKC')               // 全角英数/括弧などを半角化
    .replace(/[\s　]+/g, '')      // 空白除去
    .trim();
}

function yt_ヘッダー名マップ_(headerRows) {
  var map = {};
  for (var r = 0; r < headerRows.length; r++) {
    var row = headerRows[r] || [];
    for (var c = 0; c < row.length; c++) {
      var key = yt_ヘッダー正規化_(row[c]);
      if (!key) continue;
      if (!(key in map)) map[key] = c;  // 最初に出た列を採用
    }
  }
  return map;
}

function yt_列を解決_(headerMap, candidates) {
  for (var i = 0; i < (candidates || []).length; i++) {
    var key = yt_ヘッダー正規化_(candidates[i]);
    if (key in headerMap) return headerMap[key];
  }
  return -1;
}

function yt_商品コードキー_(v) {
  return String(v == null ? '' : v)
    .normalize('NFKC')
    .replace(/[‐‑‒–—―−ー]/g, '-')   // 各種ダッシュ/長音をハイフンに
    .replace(/[\s　]+/g, '')
    .toUpperCase()
    .trim();
}

function yt_行から送信値_(rowValues, headerMap) {
  var out = {};
  var fields = YAHOO_TRANSFER_CFG.fieldCandidates;
  Object.keys(fields).forEach(function (yahooField) {
    var col = yt_列を解決_(headerMap, fields[yahooField]);
    out[yahooField] = (col >= 0 && col < rowValues.length) ? rowValues[col] : '';
  });
  return out;
}

// 既存商品コード(→送信先行番号)のmapを使い、append（新規）と update（既存行の空セル補充）に振り分ける。
// existingCodeRowMap: { 商品コードキー: 送信先行番号 }
function yt_送信計画を作る_(collected, existingCodeRowMap) {
  var map = {};
  Object.keys(existingCodeRowMap || {}).forEach(function (k) {
    var key = yt_商品コードキー_(k);
    if (key) map[key] = existingCodeRowMap[k];
  });
  var seen = {};
  var toAppend = [], toUpdate = [], skipDup = 0, skipNoCode = 0;
  for (var i = 0; i < collected.length; i++) {
    var c = collected[i];
    var mapped = c.mapped;
    var codeKey = yt_商品コードキー_(mapped['商品コード']);
    if (!codeKey) { skipNoCode++; continue; }
    if (seen[codeKey]) { skipDup++; continue; }   // 同一実行内の重複
    seen[codeKey] = true;
    if (Object.prototype.hasOwnProperty.call(map, codeKey)) {
      toUpdate.push({ sheet: c.sheet, rowIndex: c.rowIndex, mapped: mapped, codeKey: codeKey, destRow: map[codeKey] });
    } else {
      toAppend.push({ sheet: c.sheet, rowIndex: c.rowIndex, mapped: mapped, codeKey: codeKey });
    }
  }
  return { toAppend: toAppend, toUpdate: toUpdate, skipDup: skipDup, skipNoCode: skipNoCode };
}

// 構造確認: 各ソースタブと送信先の検出列をLoggerに出す（候補名が実ヘッダーに当たっているか確認用）
function 台湾_Yahoo送信_構造を確認() {
  var ss = SpreadsheetApp.getActive();
  var cfg = YAHOO_TRANSFER_CFG;
  var lines = [];

  cfg.srcSheets.forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) { lines.push('[ソース欠落] ' + name); return; }
    var headerRows = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues();
    var map = yt_ヘッダー名マップ_(headerRows);
    var detail = Object.keys(cfg.fieldCandidates).map(function (f) {
      var c = yt_列を解決_(map, cfg.fieldCandidates[f]);
      return f + '=' + (c >= 0 ? '列' + (c + 1) : '×なし');
    }).join(' / ');
    lines.push('[ソース] ' + name + ' : ' + detail);
  });

  var dss = SpreadsheetApp.openById(cfg.destId);
  var dsh = dss.getSheetByName(cfg.destSheet);
  if (!dsh) {
    lines.push('[送信先欠落] ' + cfg.destSheet);
  } else {
    var dHeader = dsh.getRange(1, 1, cfg.destHeaderRows, dsh.getLastColumn()).getDisplayValues();
    var dMap = yt_ヘッダー名マップ_(dHeader);
    var dDetail = Object.keys(cfg.fieldCandidates).map(function (f) {
      var c = yt_列を解決_(dMap, [f]);
      return f + '=' + (c >= 0 ? '列' + (c + 1) : '×なし');
    }).join(' / ');
    lines.push('[送信先] ' + cfg.destSheet + ' : ' + dDetail);
  }

  Logger.log(lines.join('\n'));
  SpreadsheetApp.getActive().toast('構造確認をログ出力しました（実行ログを確認）');
}

// 図形ボタン割当用エントリポイント: 台湾系タブのA列チェック行を Yahoo①商品入力シートへ追記
function 台湾_Yahooテンプレへ送信() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var cfg = YAHOO_TRANSFER_CFG;

  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(2000)) { ss.toast('他の処理が実行中です。'); return; }
  try {
    // 1) 送信先を開く
    var dss = SpreadsheetApp.openById(cfg.destId);
    var dsh = dss.getSheetByName(cfg.destSheet);
    if (!dsh) { ui.alert('送信先「' + cfg.destSheet + '」が見つかりません。'); return; }
    var dHeaderRows = dsh.getRange(1, 1, cfg.destHeaderRows, dsh.getLastColumn()).getDisplayValues();
    var dMap = yt_ヘッダー名マップ_(dHeaderRows);
    var destCols = {};
    var missing = [];
    Object.keys(cfg.fieldCandidates).forEach(function (f) {
      var c = yt_列を解決_(dMap, [f]);
      if (c < 0) missing.push(f); else destCols[f] = c; // 0始まり
    });
    if (!('商品コード' in destCols) || !('商品名' in destCols)) {
      ui.alert('送信先に必須列が見つかりません（不足: ' + missing.join(', ') + '）。'); return;
    }

    // 2) 送信先の既存行を読む（商品コードキー→行番号 と、各マッピング列の現在表示値）
    var dLast = dsh.getLastRow();
    var fieldList = Object.keys(destCols);
    var existingCodeRow = {};   // codeKey -> 送信先行番号
    var existingCellVals = {};  // codeKey -> { field: 現在の表示値 }
    if (dLast >= cfg.destDataStartRow) {
      var colIdxList = fieldList.map(function (f) { return destCols[f]; });
      var minCol = Math.min.apply(null, colIdxList);
      var maxCol = Math.max.apply(null, colIdxList);
      var block = dsh.getRange(cfg.destDataStartRow, minCol + 1, dLast - cfg.destDataStartRow + 1, maxCol - minCol + 1).getDisplayValues();
      for (var r = 0; r < block.length; r++) {
        var brow = block[r];
        var codeKey = yt_商品コードキー_(brow[destCols['商品コード'] - minCol]);
        if (!codeKey || (codeKey in existingCodeRow)) continue; // 最初の行を採用
        existingCodeRow[codeKey] = cfg.destDataStartRow + r;
        var vals = {};
        fieldList.forEach(function (f) { vals[f] = brow[destCols[f] - minCol]; });
        existingCellVals[codeKey] = vals;
      }
    }

    // 3) 台湾系タブを横断してチェックON行を収集
    var collected = [];
    cfg.srcSheets.forEach(function (name) {
      var sh = ss.getSheetByName(name);
      if (!sh) return;
      var last = sh.getLastRow();
      if (last < 2) return;
      var width = sh.getLastColumn();
      var headerMap = yt_ヘッダー名マップ_(sh.getRange(1, 1, 1, width).getDisplayValues());
      var checks = sh.getRange(2, cfg.checkboxCol, last - 1, 1).getValues();
      for (var i = 0; i < checks.length; i++) {
        if (checks[i][0] !== true) continue;
        var rowIndex = i + 2;
        var values = sh.getRange(rowIndex, 1, 1, width).getDisplayValues()[0];
        collected.push({
          sheet: name, rowIndex: rowIndex,
          mapped: yt_行から送信値_(values, headerMap),
        });
      }
    });

    if (collected.length === 0) { ss.toast('チェックON行がありません。'); return; }

    // 4) 送信計画（既存=空セル補充 / 新規=追記）
    var plan = yt_送信計画を作る_(collected, existingCodeRow);
    if (plan.toAppend.length === 0 && plan.toUpdate.length === 0) {
      ss.toast('送信対象なし（重複' + plan.skipDup + ' / コード無し' + plan.skipNoCode + '）。'); return;
    }

    // 5) 確認ダイアログ
    var previewItems = plan.toAppend.concat(plan.toUpdate);
    var preview = previewItems.slice(0, 10).map(function (o) {
      return (o.destRow ? '[補充]' : '[新規]') + ' ' + o.codeKey + ' / ' + String(o.mapped['商品名'] || '').slice(0, 28);
    }).join('\n');
    var res = ui.alert('Yahooテンプレへ送信',
      '新規追記 ' + plan.toAppend.length + '件 / 既存の空セル補充 ' + plan.toUpdate.length + '件 を「' + cfg.destSheet + '」に書き込みます。\n' +
      '（重複スキップ ' + plan.skipDup + ' / コード無しスキップ ' + plan.skipNoCode + '）\n\n' +
      preview + (previewItems.length > 10 ? '\n…ほか' : '') + '\n\n実行する？',
      ui.ButtonSet.YES_NO);
    if (res !== ui.Button.YES) { ui.alert('やめました。'); return; }

    // 6a) 既存行: 空セルだけ補充（値が入っているセルは触らない）
    var filledCells = 0;
    plan.toUpdate.forEach(function (o) {
      var cur = existingCellVals[o.codeKey] || {};
      fieldList.forEach(function (f) {
        var existingVal = String(cur[f] == null ? '' : cur[f]).trim();
        var newVal = o.mapped[f];
        if (existingVal === '' && newVal != null && String(newVal).trim() !== '') {
          dsh.getRange(o.destRow, destCols[f] + 1).setValue(newVal);
          filledCells++;
        }
      });
    });

    // 6b) 新規: 最終データ行の下へ追記
    if (plan.toAppend.length > 0) {
      var startRow = Math.max(dLast + 1, cfg.destDataStartRow);
      var n = plan.toAppend.length;
      fieldList.forEach(function (f) {
        var colVals = plan.toAppend.map(function (o) { return [o.mapped[f] != null ? o.mapped[f] : '']; });
        dsh.getRange(startRow, destCols[f] + 1, n, 1).setValues(colVals);
      });
    }
    SpreadsheetApp.flush();

    // 7) 処理した行のA列チェックをOFF（新規・補充とも）
    var bySheet = {};
    plan.toAppend.concat(plan.toUpdate).forEach(function (o) {
      (bySheet[o.sheet] = bySheet[o.sheet] || []).push(o.rowIndex);
    });
    Object.keys(bySheet).forEach(function (name) {
      var sh = ss.getSheetByName(name);
      if (!sh) return;
      bySheet[name].forEach(function (rowIndex) {
        sh.getRange(rowIndex, cfg.checkboxCol).setValue(false);
      });
    });

    ss.toast('Yahoo送信: 新規 ' + plan.toAppend.length + '件 / 補充 ' + plan.toUpdate.length + '件(' + filledCells + 'セル) / 重複' + plan.skipDup + ' / コード無し' + plan.skipNoCode);
  } finally {
    lock.releaseLock();
  }
}
