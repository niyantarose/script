// productForSheetRowBuild_ の日本語タイトル採用規則の検証。
// シート行を組み立てる直前の最後の分岐であり、ここで resolved を無検証で
// 採用していると、照会側のバグ由来の中文題がそのままシートへ流れる。
// 実行: node tests/sheet-row-japanese-title.test.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');

const ctx = { console };
ctx.globalThis = ctx;
ctx.chrome = {
  storage: {
    local: { get: () => {}, set: () => {} },
    onChanged: { addListener: () => {} },
  },
  runtime: { onMessage: { addListener: () => {} }, lastError: null, sendMessage: () => {} },
  tabs: { query: () => {}, sendMessage: () => {} },
  scripting: { executeScript: () => {} },
  downloads: { download: () => {} },
};
// popup.js は起動時に多数の DOM 要素へイベントを結び付けるので、
// どの id でも同じダミー要素を返すスタブにする。
const makeStubElement = () => ({
  addEventListener: () => {},
  removeEventListener: () => {},
  appendChild: () => {},
  removeChild: () => {},
  setAttribute: () => {},
  getAttribute: () => null,
  classList: { add: () => {}, remove: () => {}, toggle: () => {}, contains: () => false },
  style: {},
  dataset: {},
  value: '',
  textContent: '',
  innerHTML: '',
  checked: false,
  disabled: false,
  files: [],
  children: [],
});
ctx.document = {
  addEventListener: () => {},
  getElementById: () => makeStubElement(),
  querySelector: () => makeStubElement(),
  querySelectorAll: () => [],
  createElement: () => makeStubElement(),
  body: makeStubElement(),
};
ctx.window = { addEventListener: () => {} };
ctx.navigator = { userAgent: 'node-test' };
ctx.location = { href: '' };
ctx.alert = () => {};
ctx.confirm = () => true;
vm.createContext(ctx);

for (const f of [
  'core/titleAnalysis.js',
  'popup/taiwan/popup.shared.js',
  'popup/taiwan/popup.books.js',
  'popup/taiwan/popup.goods.js',
  'popup/taiwan/popup.magazines.js',
  'popup/taiwan/popup.js',
]) {
  vm.runInContext(fs.readFileSync(path.join(root, f), 'utf8'), ctx, { filename: f });
}

const build = ctx.productForSheetRowBuild_;
if (typeof build !== 'function') {
  throw new Error('productForSheetRowBuild_ が読み込めない');
}

let failed = 0;
// 未設定(undefined)と空文字はどちらも「シートに何も書かれない」で同じ意味なので正規化する
function eq(name, actual, expected) {
  const a = JSON.stringify(String(actual == null ? '' : actual));
  const e = JSON.stringify(String(expected == null ? '' : expected));
  if (a !== e) { failed++; console.error(`[NG] ${name}: expected=${e} actual=${a}`); }
  else { console.log(`[OK] ${name}`); }
}

const baseProduct = {
  商品コード: 'TWS0171-CM-0102',
  商品名: '台湾版 まんが 日昇之屋',
  原題タイトル: '日昇之屋',
  URL: 'https://www.books.com.tw/products/0011053964',
};

// === ページ由来(page_trusted)は最優先で維持 ===
eq('page_trusted は最優先',
   build({ ...baseProduct, 日本語タイトル: 'How to melt', 日本語タイトル取得元: 'page_trusted' }).日本語タイトル,
   'How to melt');

// === 正しい resolved（かな入り）はそのまま採用 ===
eq('かな入りの resolved は採用',
   build({
     ...baseProduct,
     japaneseTitleLookup: {
       status: 'resolved',
       japaneseTitle: '陽が昇る家〜田舎で出会った俺たち〜',
       provider: 'mangaUpdates(extension)',
       candidates: [{ title: '陽が昇る家〜田舎で出会った俺たち〜' }],
     },
   }).日本語タイトル,
   '陽が昇る家〜田舎で出会った俺たち〜');

// === 本題: 中文題の resolved は採用せず空欄に落とす ===
// storage に残っている過去の誤 resolved が、修正後もシートへ流れないこと。
eq('中文題の resolved はシートに書かない',
   build({
     ...baseProduct,
     japaneseTitleLookup: {
       status: 'resolved',
       japaneseTitle: '日出之家',
       provider: 'mangaUpdates(extension)',
       candidates: [
         { title: '日出之家' },
         { title: '日昇之屋' },
         { title: '陽が昇る家〜田舎で出会った俺たち〜' },
       ],
     },
   }).日本語タイトル,
   '');

// === K-9型（かな無し日本語題）は壊さない ===
eq('かな無しの日本語漢字題は維持',
   build({
     商品コード: 'TWF0001-CM-01',
     商品名: 'K-9 警視廳公安部公安第9課異能對策組',
     原題タイトル: 'K-9 警視廳公安部公安第9課異能對策組',
     japaneseTitleLookup: {
       status: 'resolved',
       japaneseTitle: 'K-9 警視庁公安部公安第9課異能対策係',
       provider: 'mangaUpdates(extension)',
       candidates: [{ title: 'K-9 警視庁公安部公安第9課異能対策係' }],
     },
   }).日本語タイトル,
   'K-9 警視庁公安部公安第9課異能対策係');

process.exitCode = failed ? 1 : 0;
console.log(failed ? `\n${failed} failed` : '\nsheet-row-japanese-title.test.js: ok');
