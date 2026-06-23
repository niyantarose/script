const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const ctx = {
  console,
  chrome: {
    storage: {
      local: {
        get: () => {},
        set: () => {},
      },
    },
  },
};
ctx.globalThis = ctx;
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(root, 'popup/taiwan/popup.shared.js'), 'utf8'), ctx, {
  filename: 'popup/taiwan/popup.shared.js',
});
vm.runInContext(fs.readFileSync(path.join(root, 'popup/taiwan/popup.books.js'), 'utf8'), ctx, {
  filename: 'popup/taiwan/popup.books.js',
});

const { detectEditionType, deriveTaiwanVolumeColumns_ } = ctx;

const editionCases = [
  {
    name: '特裝版 in 原題商品タイトル',
    product: { 原題商品タイトル: '綠蔭之冠1+2小說限量特裝版', 商品名: '綠蔭之冠' },
    expected: '特装版',
  },
  {
    name: '通常版 when no edition marker',
    product: { 原題商品タイトル: '全球高考', 商品名: '全球高考' },
    expected: '通常版',
  },
];

for (const tc of editionCases) {
  const actual = detectEditionType(tc.product);
  if (actual !== tc.expected) {
    throw new Error(`${tc.name}: expected ${tc.expected}, got ${actual}`);
  }
}

const volumeCases = [
  {
    name: '1+2 bundle novel set',
    product: { 原題商品タイトル: '綠蔭之冠1+2小說限量特裝版' },
    expected: { 単巻数: '', セット巻数開始番号: '1', セット巻数終了番号: '2' },
  },
  {
    name: '上下巻 set',
    product: { 原題商品タイトル: '某某漫畫 上下巻合售' },
    expected: { 単巻数: '', セット巻数開始番号: '1', セット巻数終了番号: '2' },
  },
  {
    name: '上+下 set',
    product: { 商品名: '作品名 上+下' },
    expected: { 単巻数: '', セット巻数開始番号: '1', セット巻数終了番号: '2' },
  },
  {
    name: 'single volume default',
    product: { 原題商品タイトル: '全球高考' },
    expected: { 単巻数: '1', セット巻数開始番号: '', セット巻数終了番号: '' },
  },
];

for (const tc of volumeCases) {
  const actual = deriveTaiwanVolumeColumns_(tc.product);
  for (const key of Object.keys(tc.expected)) {
    if (actual[key] !== tc.expected[key]) {
      throw new Error(`${tc.name}: ${key} expected ${tc.expected[key]}, got ${actual[key]}`);
    }
  }
}

console.log('volume-edition.test.js: ok');
