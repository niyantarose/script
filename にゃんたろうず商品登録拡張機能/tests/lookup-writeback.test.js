// mergeEnrichedItemsToStorageProducts の書き戻し規則の検証。
// 商品ページのタイトル直下から取った日本語タイトル（取得元=page_trusted）は
// 外部照会(MangaUpdates)で上書きしない、という不変条件を守ること。
// この不変条件は core/titleAnalysis.js と popup/taiwan/popup.js が宣言しているが、
// 書き戻し経路だけが破っていた（かつ取得元ラベルを付け替えず page_trusted のまま残っていた）。
// 実行: node tests/lookup-writeback.test.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');

let storedProducts = [];
const ctx = {
  console,
  fetch: () => Promise.reject(new Error('fetch is not stubbed')),
  chrome: {
    storage: {
      local: {
        get: (keys, cb) => { cb({ products: storedProducts }); },
        set: (obj, cb) => { if (obj.products) storedProducts = obj.products; if (cb) cb(); },
      },
    },
    runtime: {
      onMessage: { addListener: () => {} },
      onStartup: { addListener: () => {} },
      onInstalled: { addListener: () => {} },
      lastError: null,
      getURL: p => p,
    },
    alarms: { create: () => {}, onAlarm: { addListener: () => {} } },
    offscreen: { createDocument: () => {}, closeDocument: () => {} },
    scripting: { executeScript: () => {} },
    tabs: { query: () => {}, sendMessage: () => {} },
    downloads: { download: () => {} },
    notifications: { create: () => {} },
    action: { onClicked: { addListener: () => {} } },
  },
};
ctx.globalThis = ctx;
ctx.self = ctx;
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(root, 'backgrounds/taiwan.js'), 'utf8'), ctx, {
  filename: 'backgrounds/taiwan.js',
});

const merge = ctx.mergeEnrichedItemsToStorageProducts;
if (typeof merge !== 'function') {
  throw new Error('mergeEnrichedItemsToStorageProducts が読み込めない');
}

let failed = 0;
function eq(name, actual, expected) {
  const a = JSON.stringify(actual), e = JSON.stringify(expected);
  if (a !== e) { failed++; console.error(`[NG] ${name}: expected=${e} actual=${a}`); }
  else { console.log(`[OK] ${name}`); }
}

const lookup = {
  status: 'resolved',
  japaneseTitle: 'MU由来の日本語題',
  provider: 'mangaUpdates(extension)',
};

async function run() {
  // === page_trusted（ページのタイトル直下から取得）は外部照会で上書きしない ===
  storedProducts = [{
    商品コード: 'TWS0001-CM-01',
    日本語タイトル: 'How to melt',
    日本語タイトル取得元: 'page_trusted',
  }];
  await merge([{ 商品コード: 'TWS0001-CM-01', japaneseTitleLookup: lookup }]);
  eq('page_trusted は上書きされない', storedProducts[0]['日本語タイトル'], 'How to melt');
  eq('page_trusted のラベルは維持', storedProducts[0]['日本語タイトル取得元'], 'page_trusted');

  // === page_scan（離れた場所の走査＝要検証）は従来どおり上書きしてよい ===
  storedProducts = [{
    商品コード: 'TWS0002-CM-01',
    日本語タイトル: '怪しい値',
    日本語タイトル取得元: 'page_scan',
  }];
  await merge([{ 商品コード: 'TWS0002-CM-01', japaneseTitleLookup: lookup }]);
  eq('page_scan は上書きされる', storedProducts[0]['日本語タイトル'], 'MU由来の日本語題');
  eq('上書き時は取得元ラベルを外部照会に付け替える',
     storedProducts[0]['日本語タイトル取得元'], 'mangaupdates');

  // === 空欄なら当然埋める ===
  storedProducts = [{ 商品コード: 'TWS0003-CM-01', 日本語タイトル: '', 日本語タイトル取得元: '' }];
  await merge([{ 商品コード: 'TWS0003-CM-01', japaneseTitleLookup: lookup }]);
  eq('空欄は補充される', storedProducts[0]['日本語タイトル'], 'MU由来の日本語題');
  eq('補充時もラベルを付ける', storedProducts[0]['日本語タイトル取得元'], 'mangaupdates');

  // === page_trusted でも空欄なら補充する（既存テスト title-analysis.test.js と同じ方針）===
  storedProducts = [{
    商品コード: 'TWS0004-CM-01',
    日本語タイトル: '',
    日本語タイトル取得元: 'page_trusted',
  }];
  await merge([{ 商品コード: 'TWS0004-CM-01', japaneseTitleLookup: lookup }]);
  eq('page_trusted でも空欄なら補充', storedProducts[0]['日本語タイトル'], 'MU由来の日本語題');

  if (failed) {
    console.error(`\n${failed} failed`);
    process.exitCode = 1;
    return;
  }
  console.log('\nlookup-writeback.test.js: ok');
}

run();
