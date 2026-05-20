const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const ctx = { console };
ctx.globalThis = ctx;
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(root, 'core/mangaUpdatesClient.js'), 'utf8'), ctx, {
  filename: 'core/mangaUpdatesClient.js',
});

const t = ctx.titleLookupMangaUpdates.__test;
const jpTitle = 'パニックテスト～これでも冷静でいられますか？～';
const queries = ['全球高考'];
const queryKeys = queries.map(t.normalizeTitleKey);

const detail = {
  title: 'Global Examination',
  associated: [
    { title: '全球高考' },
    { title: 'Quanqiu Gaokao' },
    { title: jpTitle },
  ],
};
const searchRow = {
  hit_title: '全球高考',
  record: {
    title: 'Global Examination',
    associated: [{ title: '全球高考' }],
  },
};

const titles = t.collectTitlesForMatch(detail, searchRow);
if (!titles.includes('全球高考')) {
  throw new Error('collectTitlesForMatch should include associated Chinese title');
}

if (!t.detailMatchesQueries(detail, queryKeys, queries, searchRow)) {
  throw new Error('detailMatchesQueries should match Chinese query via search row associated');
}

const withoutRow = t.detailMatchesQueries(
  { title: 'Global Examination', associated: [{ title: jpTitle }] },
  queryKeys,
  queries,
  null
);
if (withoutRow) {
  throw new Error('detailMatchesQueries should not match when associated lacks query');
}

const picked = t.pickJapaneseFromDetail(detail, queries, { seriesVerified: true });
if (picked !== jpTitle) {
  throw new Error(`pickJapaneseFromDetail expected JP license title, got: ${picked}`);
}

const blocked = t.pickJapaneseFromDetail(detail, queries, { seriesVerified: false });
if (blocked === jpTitle) {
  throw new Error('pickJapaneseFromDetail without seriesVerified should reject kana-only JP for han query');
}

const numericId = t.base36SlugToDecimalString('xvytm0b');
if (numericId !== '73766757035') {
  throw new Error(`xvytm0b should map to 73766757035, got ${numericId}`);
}

const resolved = t.tryResolveMatchedDetail_(detail, searchRow, queryKeys, queries, []);
if (resolved.japaneseTitle !== jpTitle) {
  throw new Error(`tryResolveMatchedDetail_ expected JP title, got ${resolved.japaneseTitle}`);
}

console.log('manga-updates-client.test.js: ok');
