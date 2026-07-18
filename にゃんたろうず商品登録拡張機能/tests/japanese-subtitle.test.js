const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const ctx = {
  console,
  chrome: {
    runtime: {
      onMessage: {
        addListener: () => {},
      },
    },
  },
};
ctx.globalThis = ctx;
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(root, 'content/taiwan/content.js'), 'utf8'), ctx, {
  filename: 'content/taiwan/content.js',
});

const { acceptsTrustedOriginalSubtitle } = ctx;

if (typeof acceptsTrustedOriginalSubtitle !== 'function') {
  throw new Error('acceptsTrustedOriginalSubtitle が content.js のトップレベルに定義されていない');
}

// h1直下のh2（博客來の原文書名欄）テキストを日本語原題として採用してよいかの判定。
// 引数: (h2テキスト, h1テキスト)
const cases = [
  {
    name: '原題が英語表記のみの日本作品（How to melt）を採用する',
    subtitle: 'How to melt',
    title: 'How to melt(全)特裝版',
    expected: true,
  },
  {
    name: 'かな入り日本語題は従来どおり採用する',
    subtitle: 'ボーイッシュ彼女が可愛すぎる',
    title: '男孩子氣的女友超級可愛(02)',
    expected: true,
  },
  {
    name: '漢字のみ（中文副題の可能性）は採用しない',
    subtitle: '后宮的Ω王子 雪花之章',
    title: '后宮的Ω王子 雪花之章(全)',
    expected: false,
  },
  {
    name: 'ハングル（韓国語原題）は採用しない',
    subtitle: '이번 생은 가주가 되겠습니다',
    title: '今生我來當家主(01)',
    expected: false,
  },
  {
    name: '英字でも漢字が混在するものは採用しない',
    subtitle: 'How to melt 漫畫',
    title: 'How to melt(全)特裝版',
    expected: false,
  },
  {
    name: 'h1タイトルのエコー（同一文字列）は採用しない',
    subtitle: 'How to melt(全)特裝版',
    title: 'How to melt(全)特裝版',
    expected: false,
  },
  {
    name: '空白差だけのエコーも採用しない',
    subtitle: ' How  to melt(全)特裝版 ',
    title: 'How to melt(全)特裝版',
    expected: false,
  },
  {
    name: '英字1文字は短すぎるので採用しない',
    subtitle: 'A',
    title: '某作品(01)',
    expected: false,
  },
  {
    name: '数字のみ（英字レター無し）は採用しない',
    subtitle: '2026',
    title: '某作品(01)',
    expected: false,
  },
  {
    name: '空文字は採用しない',
    subtitle: '',
    title: '某作品(01)',
    expected: false,
  },
  {
    name: '記号入り英字題（SPY×FAMILY等の×）は採用する',
    subtitle: 'SPY×FAMILY',
    title: '間諜家家酒(01)',
    expected: true,
  },
];

for (const tc of cases) {
  const actual = acceptsTrustedOriginalSubtitle(tc.subtitle, tc.title);
  if (actual !== tc.expected) {
    throw new Error(`${tc.name}: expected ${tc.expected}, got ${actual} (subtitle=${JSON.stringify(tc.subtitle)})`);
  }
}

console.log('japanese-subtitle.test.js: ok');
