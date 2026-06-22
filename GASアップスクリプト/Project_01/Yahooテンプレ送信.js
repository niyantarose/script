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
