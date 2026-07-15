// bookCodeWorkCompareKey_ の末尾媒体語吸収の検証（0033/0128型の二重Works防止）。
// 実行: node tests/compare-key.test.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ctx = { console };
ctx.globalThis = ctx;
ctx.SpreadsheetApp = {};
ctx.LockService = {};
ctx.PropertiesService = {};
ctx.Logger = { log: () => {} };
ctx.Utilities = {};
ctx.CacheService = {};
ctx.UrlFetchApp = {};
ctx.ContentService = {};
ctx.XmlService = {};
vm.createContext(ctx);
vm.runInContext(
  fs.readFileSync(path.join(__dirname, '..', '博客來ウェブアプリ.js'), 'utf8'),
  ctx,
  { filename: '博客來ウェブアプリ.js' }
);

let failed = 0;
function eq(name, actual, expected) {
  const a = JSON.stringify(actual), e = JSON.stringify(expected);
  if (a !== e) { failed++; console.error(`[NG] ${name}: expected=${e} actual=${a}`); }
  else { console.log(`[OK] ${name}`); }
}

eq('22漫畫を吸収',
   ctx.bookCodeWorkCompareKey_('我獨自升級22漫畫'),
   ctx.bookCodeWorkCompareKey_('我獨自升級'));
eq('21漫畫も同一',
   ctx.bookCodeWorkCompareKey_('我獨自升級21漫畫'),
   ctx.bookCodeWorkCompareKey_('我獨自升級'));
eq('末尾小說も吸収',
   ctx.bookCodeWorkCompareKey_('某作品3小說'),
   ctx.bookCodeWorkCompareKey_('某作品'));
eq('中間の漫畫は残す', ctx.bookCodeWorkCompareKey_('漫畫學院物語') !== '', true);
eq('第3巻の除去は従来どおり',
   ctx.bookCodeWorkCompareKey_('作品A 第3巻'),
   ctx.bookCodeWorkCompareKey_('作品A'));

process.exitCode = failed ? 1 : 0;
console.log(failed ? `\n${failed} failed` : '\nall passed');
