const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const ctx = { console };
ctx.globalThis = ctx;
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(root, 'core/titleAnalysis.js'), 'utf8'), ctx, {
  filename: 'core/titleAnalysis.js',
});

const tests = [
  {
    name: 'books light novel',
    item: {
      source: 'books_tw',
      title: '小説 葬送的芙莉蓮～前奏～（首刷限定版）2',
      categoryName: '輕小說',
      language: '繁體中文',
    },
    expected: {
      itemType: 'light_novel',
      extractedWorkTitle: '葬送的芙莉蓮',
      subtitle: '前奏',
      edition: '首刷限定版',
      volume: '2',
      normalizedSearchTitle: '葬送的芙莉蓮 前奏',
    },
  },
  {
    name: 'books tw goods',
    item: {
      source: 'books_tw',
      title: '全知讀者視角(漫畫) 炫彩壓克力色紙 D',
      categoryName: '動漫周邊',
    },
    expected: {
      itemType: 'goods',
      extractedWorkTitle: '全知讀者視角',
      goodsType: '炫彩壓克力色紙',
      optionType: 'D',
      normalizedSearchTitle: '全知讀者視角',
    },
  },
  {
    name: 'books tw goods folder 2-pack',
    item: {
      source: 'books_tw',
      title: '排球少年!! 資料夾2入',
      categoryName: '動漫周邊',
    },
    expected: {
      itemType: 'goods',
      extractedWorkTitle: '排球少年',
      goodsType: '資料夾',
      normalizedSearchTitle: '排球少年',
    },
  },
  {
    name: 'books tw goods folder 2-pack without space',
    item: {
      source: 'books_tw',
      title: '排球少年!!資料夾2入',
      categoryName: '動漫周邊',
    },
    expected: {
      itemType: 'goods',
      extractedWorkTitle: '排球少年',
      goodsType: '資料夾',
      normalizedSearchTitle: '排球少年',
    },
  },
  {
    name: 'books tw compact volume manga suffix',
    item: {
      source: 'books_tw',
      title: '我獨自升級21漫畫',
      categoryName: '漫畫',
    },
    expected: {
      itemType: 'manga',
      extractedWorkTitle: '我獨自升級',
      volume: '21',
      normalizedSearchTitle: '我獨自升級',
    },
  },
  {
    name: 'aladin manga',
    item: {
      source: 'aladin',
      title: '나의 히어로 아카데미아 42 한정판',
    },
    expected: {
      itemType: 'manga',
      extractedWorkTitle: '나의 히어로 아카데미아',
      volume: '42',
      edition: '한정판',
      normalizedSearchTitle: '나의 히어로 아카데미아',
    },
  },
  {
    name: 'magazine',
    item: {
      source: 'aladin',
      title: 'GQ KOREA 2026.03',
      categoryName: '잡지',
    },
    expected: {
      itemType: 'magazine',
      magazineName: 'GQ',
      year: '2026',
      month: '3',
      normalizedSearchTitle: 'GQ',
    },
  },
  {
    name: 'unknown overseas magazine no issue',
    item: {
      source: 'books_tw',
      title: 'Bella Style No.5',
      categoryName: '雜誌',
    },
    expected: {
      itemType: 'magazine',
      magazineName: 'Bella Style',
      issue: '5',
      extractedWorkTitle: 'Bella Style',
      normalizedSearchTitle: 'Bella Style',
    },
  },
  {
    name: 'korean goods',
    item: {
      source: 'korean_goods',
      title: '죽음의 교실 44교시 서바이벌 랜덤 SD 아크릴 키링',
    },
    expected: {
      itemType: 'goods',
      extractedWorkTitle: '죽음의 교실 44교시 서바이벌',
      goodsType: '랜덤 SD 아크릴 키링',
      normalizedSearchTitle: '죽음의 교실 44교시 서바이벌',
    },
  },
];

let failed = 0;

for (const test of tests) {
  const actual = ctx.analyzeProductTitle(test.item);
  for (const [key, expected] of Object.entries(test.expected)) {
    if (actual[key] !== expected) {
      failed += 1;
      console.error(`[NG] ${test.name} ${key}: expected=${expected} actual=${actual[key]}`);
    }
  }
  if (!failed) {
    console.log(`[OK] ${test.name}: ${actual.itemType} / ${actual.normalizedSearchTitle}`);
  }
}

const payload = ctx.buildPayloadForGas({
  action: 'updateAladinData',
  source: 'aladin',
  itemId: '123',
  title: '나의 히어로 아카데미아 42 한정판',
});

if (payload.action !== 'updateAladinData' || !payload.titleAnalysis) {
  failed += 1;
  console.error('[NG] payload compatibility');
} else {
  console.log(`[OK] payload titleAnalysis: ${payload.titleAnalysis.itemType}`);
}

const source = ctx.detectSourceFromUrl('https://www.books.com.tw/products/0011041344');
if (source !== 'books_tw') {
  failed += 1;
  console.error(`[NG] detectSourceFromUrl: expected=books_tw actual=${source}`);
} else {
  console.log('[OK] detectSourceFromUrl books_tw');
}

const lookupPayload = ctx.buildLookupPayload({}, ctx.analyzeProductTitle({
  source: 'books_tw',
  title: '排球少年SS-多功能手機包(2)',
}));
if (lookupPayload.action !== 'lookupJapaneseTitle' || lookupPayload.titleAnalysis.normalizedSearchTitle !== '排球少年') {
  failed += 1;
  console.error('[NG] lookup payload normalizedSearchTitle');
} else {
  console.log(`[OK] lookup payload: ${lookupPayload.titleAnalysis.normalizedSearchTitle}`);
}

const savePayload = ctx.buildSavePayload(
  { itemId: '001', title: '排球少年SS-多功能手機包(2)' },
  lookupPayload.titleAnalysis,
  { status: 'resolved', japaneseTitle: 'ハイキュー!!', provider: 'titleAliasDictionary' }
);
if (savePayload.action !== 'upsertProductWithLookup' || savePayload.japaneseTitleLookup.japaneseTitle !== 'ハイキュー!!') {
  failed += 1;
  console.error('[NG] save payload with japaneseTitleLookup');
} else {
  console.log('[OK] save payload with lookup');
}

process.exit(failed ? 1 : 0);
