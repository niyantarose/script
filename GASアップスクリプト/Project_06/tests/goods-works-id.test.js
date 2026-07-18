// 韓国グッズWorksのID採番（getLastRow廃止→最大値+1+ハイウォーター）の検証。
// 実行: node tests/goods-works-id.test.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');

let hwStore = 0;
const ctx = { console };
ctx.globalThis = ctx;
ctx.SpreadsheetApp = { getActive: () => ({}), getActiveSpreadsheet: () => ({}), getUi: () => ({ alert: () => {} }), flush: () => {} };
ctx.PropertiesService = {};
ctx.LockService = {};
ctx.Logger = { log: () => {} };
ctx.CacheService = {};
ctx.UrlFetchApp = {};
ctx._kyoutuu = {
  採番ハイウォーター読取: () => hwStore,
  採番ハイウォーター更新: (cfg, v) => { hwStore = v; },
};
vm.createContext(ctx);
vm.runInContext(
  fs.readFileSync(path.join(__dirname, '..', '韓国グッズ.js'), 'utf8'),
  ctx,
  { filename: '韓国グッズ.js' }
);

let failed = 0;
function eq(name, actual, expected) {
  const a = JSON.stringify(actual), e = JSON.stringify(expected);
  if (a !== e) { failed++; console.error(`[NG] ${name}: expected=${e} actual=${a}`); }
  else { console.log(`[OK] ${name}`); }
}

function makeWorks(ids) {
  return {
    getLastRow: () => ids.length + 1,
    getRange: (r, c, nr, nc) => ({ getDisplayValues: () => ids.map(id => [id]) }),
  };
}

// 既存 KR-W-0003 / KR-W-0007 → 次は 0008（getLastRowの3ではない）
hwStore = 0;
eq('最大値+1で採番', ctx.韓国グッズ_次のWorksID_(makeWorks(['KR-W-0003', 'KR-W-0007'])), 'KR-W-0008');
eq('採番後にHWが更新される', hwStore, 8);

// 行削除で最大が下がってもHWが守る（旧実装はここで重複IDを発行していた）
hwStore = 9;
eq('HW優先: 削除後も9+1', ctx.韓国グッズ_次のWorksID_(makeWorks(['KR-W-0003'])), 'KR-W-0010');

// Worksが空でもHWから継続
hwStore = 5;
eq('空Works: HW+1', ctx.韓国グッズ_次のWorksID_(makeWorks([])), 'KR-W-0006');

// ライブラリ呼び出しが失敗しても採番は止まらない（縮退）
ctx._kyoutuu = { 採番ハイウォーター読取: () => { throw new Error('boom'); }, 採番ハイウォーター更新: () => { throw new Error('boom'); } };
eq('ライブラリ例外でも継続', ctx.韓国グッズ_次のWorksID_(makeWorks(['KR-W-0002'])), 'KR-W-0003');

process.exitCode = failed ? 1 : 0;
console.log(failed ? `\n${failed} failed` : '\nall passed');
