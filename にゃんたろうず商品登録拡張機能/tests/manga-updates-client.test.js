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

// seriesVerified では中文原題エコー（= クエリ一致）を除外し日本語ライセンス題を採用する。
// 実フロー（tryResolveMatchedDetail_）と同様に echoQueries を渡す。
const picked = t.pickJapaneseFromDetail(detail, queries, { seriesVerified: true, echoQueries: queries });
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

const k9JpTitle = 'K-9 警視庁公安部公安第9課異能対策係';
const k9Detail = {
  title: 'K-9',
  associated: [
    { title: k9JpTitle },
    { title: 'K-9: Public Security Bureau, Division 9 - Special Abilities Countermeasure' },
  ],
};
const k9SearchRow = {
  hit_title: 'K9 警視庁公安部公安第9課異能特捜班',
  record: { title: 'K-9', associated: [{ title: k9JpTitle }] },
};
const k9Queries = ['K9 警視庁公安部公安第9課異能特捜班'];
const k9QueryKeys = k9Queries.map(t.normalizeTitleKey);
const k9Resolved = t.tryResolveMatchedDetail_(k9Detail, k9SearchRow, k9QueryKeys, k9Queries, []);
if (k9Resolved.japaneseTitle !== k9JpTitle) {
  throw new Error(`K-9 kanji-only JP title expected ${k9JpTitle}, got ${k9Resolved.japaneseTitle}`);
}

// === 繁体字クエリ（博客來原題）↔ 日本語題（字体差: 廳/庁, 對/対, 組/係）でも
//     強い漢字重なりで正解シリーズを採用し、Associated 先頭の日本語題を返す ===
const k9RealDetail = {
  title: 'K-9: Keishichou Kouanbu Kouan Dai-9 Ka Inou Taisaku-gakari',
  associated: [
    { title: 'K-9 警視庁公安部公安第9課異能対策係' },
    { title: 'K-9: Keishichou Kouanbu Kouan Dai-9-ka Inou Taisaku Gakari' },
    { title: 'K-9: Public Security Bureau, Division 9 - Special Abilities Countermeasure' },
    { title: 'K-9~警視庁公安部公安第9課異能対策係~' },
    { title: 'ケーナイン　警視庁公安部公安第９課異能対策係' },
    { title: 'Ｋ－９　警視庁公安部公安第９課異能対策係' },
  ],
};
const k9RealRow = {
  hit_title: 'K-9 警視庁公安部公安第9課異能対策係',
  record: { title: 'K-9: Keishichou Kouanbu Kouan Dai-9 Ka Inou Taisaku-gakari' },
};
const k9CnQueries = ['K-9 警視廳公安部公安第9課異能對策組'];
const k9CnKeys = k9CnQueries.map(t.normalizeTitleKey);
if (!t.detailMatchesQueries(k9RealDetail, k9CnKeys, k9CnQueries, k9RealRow)) {
  throw new Error('detailMatchesQueries should match correct K-9 series via strong Han overlap (traditional query)');
}
const k9RealResolved = t.tryResolveMatchedDetail_(k9RealDetail, k9RealRow, k9CnKeys, k9CnQueries, []);
if (k9RealResolved.japaneseTitle !== 'K-9 警視庁公安部公安第9課異能対策係') {
  throw new Error(`K-9 should resolve to first associated JP title, got ${k9RealResolved.japaneseTitle}`);
}

// === 中文原題エコー（クエリと一致）は除外し、日本語ライセンス題（カナ）を採用する ===
const manhuaDetail = {
  title: 'Some Chinese Manhua',
  associated: [
    { title: '螢幕情緣' },
    { title: 'モニターごしの恋' },
  ],
};
const manhuaRow = { hit_title: '螢幕情緣', record: { title: 'Some Chinese Manhua' } };
const manhuaQueries = ['螢幕情緣'];
const manhuaKeys = manhuaQueries.map(t.normalizeTitleKey);
const manhuaResolved = t.tryResolveMatchedDetail_(manhuaDetail, manhuaRow, manhuaKeys, manhuaQueries, []);
if (manhuaResolved.japaneseTitle !== 'モニターごしの恋') {
  throw new Error(`Chinese-original echo should be skipped for JP license title, got ${manhuaResolved.japaneseTitle}`);
}

// === 簡体字・繁体字が混在する中文原題もエコーとして除外する ===
// 博客來の商品名「今生我来當家主」に対し、MU は「今生我來當家主」と
// 日本語ライセンス題「今世は当主になります」を Associated Names に返す。
const matriarchDetail = {
  title: 'I Shall Master This Family',
  associated: [
    { title: '今生我來當家主' },
    { title: '今世は当主になります' },
    { title: '이번 생은 가주가 되겠습니다' },
    { title: 'I Shall Master This Family' },
  ],
};
const matriarchRow = {
  hit_title: '今生我來當家主',
  record: { title: 'I Shall Master This Family' },
};
const matriarchQueries = ['今生我来當家主'];
const matriarchKeys = matriarchQueries.map(t.normalizeTitleKey);
const matriarchResolved = t.tryResolveMatchedDetail_(
  matriarchDetail,
  matriarchRow,
  matriarchKeys,
  matriarchQueries,
  []
);
if (matriarchResolved.japaneseTitle !== '今世は当主になります') {
  throw new Error(`Mixed-script Chinese echo should be skipped, got: ${matriarchResolved.japaneseTitle}`);
}

// === 部分一致の中国語エコーを除外し、かな入り日本語原題を採用する ===
// クエリ「披著狼皮的羊公主」に対し、Associated 先頭の「披著狼皮的羊」を
// 日本語題として誤採用しないこと。
const sheepJp = '狼の皮をかぶった羊姫';
const sheepDetail = {
  title: "Sheep Princess in Wolf's Clothing",
  associated: [
    { title: '披著狼皮的羊' },
    { title: '披著狼皮的羊公主' },
    { title: "Sheep Princess in Wolf's Clothing" },
    { title: sheepJp },
  ],
};
const sheepRow = {
  hit_title: '披著狼皮的羊公主',
  record: {
    title: "Sheep Princess in Wolf's Clothing",
    associated: [{ title: '披著狼皮的羊公主' }],
  },
};
const sheepQueries = ['披著狼皮的羊公主'];
const sheepKeys = sheepQueries.map(t.normalizeTitleKey);
const sheepResolved = t.tryResolveMatchedDetail_(
  sheepDetail,
  sheepRow,
  sheepKeys,
  sheepQueries,
  []
);
if (sheepResolved.japaneseTitle !== sheepJp) {
  throw new Error(
    `partial Chinese echo should yield JP title ${sheepJp}, got: ${sheepResolved.japaneseTitle}`
  );
}

// 簡体字別名（披着狼皮的羊）が先頭でも同様に除外できること（著⇔着の字体差対応）。
// 実シートで誤採用された値そのもの。
const sheepSimplifiedDetail = {
  title: "Sheep Princess in Wolf's Clothing",
  associated: [
    { title: '披着狼皮的羊' },
    { title: '披著狼皮的羊公主' },
    { title: sheepJp },
  ],
};
const sheepSimplifiedResolved = t.tryResolveMatchedDetail_(
  sheepSimplifiedDetail,
  sheepRow,
  sheepKeys,
  sheepQueries,
  []
);
if (sheepSimplifiedResolved.japaneseTitle !== sheepJp) {
  throw new Error(
    `simplified Chinese echo should be skipped, got: ${sheepSimplifiedResolved.japaneseTitle}`
  );
}

console.log('manga-updates-client.test.js: ok');
