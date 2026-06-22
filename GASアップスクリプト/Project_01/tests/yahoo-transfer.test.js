const fs = require('fs');
const path = require('path');
const vm = require('vm');
const ctx = { console };
ctx.globalThis = ctx;
// SpreadsheetApp 等のGAS APIはダミー（純粋ヘルパーのみ検証するため未使用）
ctx.SpreadsheetApp = {};
ctx.LockService = {};
ctx.Logger = { log: () => {} };
vm.createContext(ctx);
vm.runInContext(
  fs.readFileSync(path.join(__dirname, '..', 'Yahooテンプレ送信.js'), 'utf8'),
  ctx,
  { filename: 'Yahooテンプレ送信.js' }
);

let failed = 0;
function eq(name, actual, expected) {
  const a = JSON.stringify(actual), e = JSON.stringify(expected);
  if (a !== e) { failed++; console.error(`[NG] ${name}: expected=${e} actual=${a}`); }
  else { console.log(`[OK] ${name}`); }
}

// ヘッダー2行を統合し、各列名→indexを解決できる
const headerRows = [
  ['発番発行', '登録状況', '商品コードステータス', '親コード', 'タイトル', '作者', '', 'リンク', '', '', '', '', '', '', '', '', '', '', '', '', '配送パターン', '', 'ISBN', '発売日'],
  [],
];
const map = ctx.yt_ヘッダー名マップ_(headerRows);
eq('タイトル列', ctx.yt_列を解決_(map, ['タイトル', '商品名（出品用）', '雑誌名']), 4);
eq('商品コード列(親コード優先)', ctx.yt_列を解決_(map, ['親コード', '商品コード（SKU）']), 3);
eq('発売日列', ctx.yt_列を解決_(map, ['発売日']), 23);
eq('ISBN列', ctx.yt_列を解決_(map, ['ISBN']), 22);
eq('配送パターン列', ctx.yt_列を解決_(map, ['配送パターン']), 20);
eq('無い候補は-1', ctx.yt_列を解決_(map, ['存在しない列']), -1);

// 括弧/空白ゆれを吸収（全角括弧・前後空白）
const map2 = ctx.yt_ヘッダー名マップ_([['  商品コード（ＳＫＵ） ']]);
eq('括弧空白ゆれ吸収', ctx.yt_列を解決_(map2, ['商品コード（SKU）']), 0);

// 商品コードキー: NFKC・空白除去・ハイフン統一・大文字化
eq('コードキー全角', ctx.yt_商品コードキー_('ＴＷＳ０１５２-ＣＭ-０４'), 'TWS0152-CM-04');
eq('コードキー空白/ダッシュ', ctx.yt_商品コードキー_(' tws0152‐cm‐04 '), 'TWS0152-CM-04');
eq('コードキー空', ctx.yt_商品コードキー_('  '), '');

// 行マッピング: ヘッダーに対応する列の値を取る。無い項目は空。
const mapM = ctx.yt_ヘッダー名マップ_([
  ['発番発行','登録状況','商品コードステータス','親コード','タイトル','作者','','','','','','','','','','','','','','','配送パターン','','ISBN','発売日'],
]);
const row = ['', true, '生成済み', 'TWS0152-CM-04', '台湾版まんが(特装版)『…』', '原作:…', '','','','','','','','','','','','','','','佐川','','9784000000000','2026/06/18'];
const mapped = ctx.yt_行から送信値_(row, mapM);
eq('商品名', mapped['商品名'], '台湾版まんが(特装版)『…』');
eq('商品コード', mapped['商品コード'], 'TWS0152-CM-04');
eq('発売日', mapped['発売日'], '2026/06/18');
eq('JANコード', mapped['JANコード'], '9784000000000');
eq('配送グループ管理番号', mapped['配送グループ管理番号'], '佐川');

process.exit(failed ? 1 : 0);
