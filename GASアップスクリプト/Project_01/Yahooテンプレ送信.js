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

function yt_送信計画を作る_(collected, existingCodeKeys) {
  var seen = {};
  (existingCodeKeys || []).forEach(function (k) {
    var key = yt_商品コードキー_(k);
    if (key) seen[key] = true;
  });
  var toSend = [], skipDup = 0, skipNoCode = 0;
  for (var i = 0; i < collected.length; i++) {
    var c = collected[i];
    var mapped = c.mapped;
    var codeKey = yt_商品コードキー_(mapped['商品コード']);
    if (!codeKey) { skipNoCode++; continue; }
    if (seen[codeKey]) { skipDup++; continue; }
    seen[codeKey] = true;
    toSend.push({ sheet: c.sheet, rowIndex: c.rowIndex, mapped: mapped, codeKey: codeKey });
  }
  return { toSend: toSend, skipDup: skipDup, skipNoCode: skipNoCode };
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
