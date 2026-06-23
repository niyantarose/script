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

// 送信計画: 既存=update(空セル補充対象) / 新規=append / 重複・コード空を除外
const mk = (code, title) => ({
  sheet:'台湾まんが', rowIndex:10,
  mapped: { '商品名':title, '商品コード':code, '発売日':'2026/06/18', 'JANコード':'', '配送グループ管理番号':'佐川' },
});
const collected = [ mk('TWS0001-CM-01','A'), mk('TWS0001-CM-01','A-dup'), mk('','B'), mk('TWS0002-CM-01','C') ];
// 送信先には TWS0001-CM-01 が 5行目に既存
const plan = ctx.yt_送信計画を作る_(collected, { 'TWS0001-CM-01': 5 });
// 1件目=既存→update(destRow5)、2件目=同一実行重複で除外、3件目=コード空、4件目=新規→append
eq('append件数', plan.toAppend.length, 1);
eq('append対象コード', plan.toAppend[0].codeKey, 'TWS0002-CM-01');
eq('update件数', plan.toUpdate.length, 1);
eq('update対象コード', plan.toUpdate[0].codeKey, 'TWS0001-CM-01');
eq('update対象destRow', plan.toUpdate[0].destRow, 5);
eq('重複スキップ件数', plan.skipDup, 1);
eq('コード空スキップ件数', plan.skipNoCode, 1);

process.exit(failed ? 1 : 0);
