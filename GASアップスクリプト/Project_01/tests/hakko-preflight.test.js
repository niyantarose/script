// 確定発行プリフライト検査と作品比較キー強化の検証（ネットワーク・GAS不要）。
// 実行: node tests/hakko-preflight.test.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ctx = { console };
ctx.globalThis = ctx;
ctx.SpreadsheetApp = {};
ctx.LockService = {};
ctx.PropertiesService = {};
ctx.Logger = { log: () => {} };
ctx.CacheService = {};
ctx.UrlFetchApp = {};
vm.createContext(ctx);
vm.runInContext(
  fs.readFileSync(path.join(__dirname, '..', '台湾書籍系_共通.js'), 'utf8'),
  ctx,
  { filename: '台湾書籍系_共通.js' }
);

let failed = 0;
function eq(name, actual, expected) {
  const a = JSON.stringify(actual), e = JSON.stringify(expected);
  if (a !== e) { failed++; console.error(`[NG] ${name}: expected=${e} actual=${a}`); }
  else { console.log(`[OK] ${name}`); }
}

// ---- 比較キー: 末尾の「<巻数>漫畫/小說」ノイズを吸収（0033/0128二重Worksの原因）----
eq('比較キー: 22漫畫を吸収',
   ctx.台湾書籍系_作品比較キー_('我獨自升級22漫畫'),
   ctx.台湾書籍系_作品比較キー_('我獨自升級'));
eq('比較キー: 21漫畫も同一',
   ctx.台湾書籍系_作品比較キー_('我獨自升級21漫畫'),
   ctx.台湾書籍系_作品比較キー_('我獨自升級'));
eq('比較キー: 末尾小說も吸収',
   ctx.台湾書籍系_作品比較キー_('某作品3小說'),
   ctx.台湾書籍系_作品比較キー_('某作品'));
// 中間に「漫畫」を含む題は変えない（末尾だけ）
eq('比較キー: 中間の漫畫は残す',
   ctx.台湾書籍系_作品比較キー_('漫畫學院物語') !== '', true);
eq('比較キー: 従来の巻数除去は不変',
   ctx.台湾書籍系_作品比較キー_('作品A 第3巻'),
   ctx.台湾書籍系_作品比較キー_('作品A'));

// ---- 確定発行プリフライト検査（純関数）----
// 対象行: {row, code, sku, status, 作品ID列, 原題, 作者}
const 全コード出現 = {
  'TWS0100-CM-01': [81, 130],   // シート内重複（実例）
  'TW0200-CM-01': [10],
  'TW0201-CM-02': [11],
  'TWF0025-CM-06': [12],
};
const works = [
  { id: '0033', 原題: '我獨自升級', 作者: 'DUBU' },
  { id: '0128', 原題: '我獨自升級22漫畫', 作者: 'DUBU' },     // 二重Works（同著者）
  { id: '0050', 原題: '同名タイトル', 作者: '作者X' },
  { id: '0051', 原題: '同名タイトル', 作者: '作者Y' },         // 著者違い→二重扱いしない
  { id: '0025', 原題: '惡靈剋星', 作者: 'ネオショコ' },
];

// 正常行: 問題なし
let r = ctx.台湾書籍系_確定発行検査_(
  [{ row: 10, code: 'TW0200-CM-01', sku: 'TW0200-CM-01', status: '', 作品ID列: '0200', 原題: '新作品', 作者: 'A' }],
  全コード出現, works);
eq('検査: 正常行は素通し', r, []);

// ⚠️ステータス持ち越し: onEditガードの警告が残る行は確定しない
r = ctx.台湾書籍系_確定発行検査_(
  [{ row: 11, code: 'TW0201-CM-02', sku: '', status: '⚠️ID不整合(親:0201/自動:0200/SKU:-)', 作品ID列: '0200', 原題: 'X', 作者: '' }],
  全コード出現, works);
eq('検査: ⚠️ステータス行はブロック', r.length === 1 && r[0].row === 11 && /ID不整合/.test(r[0].理由), true);

// 行内ID不整合（コード⇔作品ID列）
r = ctx.台湾書籍系_確定発行検査_(
  [{ row: 12, code: 'TWF0025-CM-06', sku: 'TWF0027-CM-06', status: '', 作品ID列: '0027', 原題: '惡靈剋星', 作者: '' }],
  全コード出現, works);
eq('検査: コードとSKUのID食い違いを検出', r.length === 1 && /不整合/.test(r[0].理由), true);

// シート内コード重複
r = ctx.台湾書籍系_確定発行検査_(
  [{ row: 81, code: 'TWS0100-CM-01', sku: '', status: '', 作品ID列: '0100', 原題: '鬼今天也等待著雨', 作者: '' }],
  全コード出現, works);
eq('検査: コード重複を検出(相手行付き)', r.length === 1 && /重複/.test(r[0].理由) && /130/.test(r[0].理由), true);

// Works二重（同著者・同比較キー）に属する作品ID
r = ctx.台湾書籍系_確定発行検査_(
  [{ row: 20, code: 'TW0128-CM-23', sku: '', status: '', 作品ID列: '0128', 原題: '我獨自升級23漫畫', 作者: 'DUBU' }],
  全コード出現, works);
eq('検査: Works二重(0033/0128型)を検出', r.length === 1 && /Works二重/.test(r[0].理由) && /0033/.test(r[0].理由), true);

// 著者が違う同名タイトルは二重扱いしない
r = ctx.台湾書籍系_確定発行検査_(
  [{ row: 21, code: 'TW0050-CM-01', sku: '', status: '', 作品ID列: '0050', 原題: '同名タイトル', 作者: '作者X' }],
  全コード出現, works);
eq('検査: 著者違い同名は素通し', r, []);

process.exitCode = failed ? 1 : 0;
console.log(failed ? `\n${failed} failed` : '\nall passed');
