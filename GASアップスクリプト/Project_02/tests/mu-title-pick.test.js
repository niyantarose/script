// pickJapaneseMangaUpdatesTitle_ の中国語エコー除外の検証。
// 「披著狼皮的羊公主」のようなクエリに対し、MU Associated の中国語
// （簡体字違い・部分一致含む）を日本語タイトルとして採用しないこと。
// 実行: node tests/mu-title-pick.test.js
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

const query = '披著狼皮的羊公主';

// ケース1: かな入り日本語原題があれば、それを採用（中国語エコーは無視）
let detail = {
  title: "Sheep Princess in Wolf's Clothing",
  associated: [
    { title: '披着狼皮的羊' },        // 簡体字の部分エコー（実際にシートに誤採用された値）
    { title: '披著狼皮的羊公主' },    // 完全エコー
    { title: '狼の皮をかぶった羊姫' },
  ],
};
eq('かな入り日本語原題を採用',
   ctx.pickJapaneseMangaUpdatesTitle_({}, detail, [query]),
   '狼の皮をかぶった羊姫');

// ケース2: 中国語エコーしか無ければ空文字（中国語を「解決」にしない）
detail = {
  title: "Sheep Princess in Wolf's Clothing",
  associated: [
    { title: '披着狼皮的羊' },
    { title: '披著狼皮的羊公主' },
  ],
};
eq('エコーのみなら空文字',
   ctx.pickJapaneseMangaUpdatesTitle_({}, detail, [query]),
   '');

// ケース3: クエリと無関係な漢字のみ日本語題（K-9型）は残す
detail = {
  title: 'K-9',
  associated: [
    { title: 'K-9 警視庁公安部公安第9課異能対策係' },
  ],
};
eq('無関係の漢字題は従来どおり採用',
   ctx.pickJapaneseMangaUpdatesTitle_({}, detail, ['K-9 警視廳公安部公安第9課異能對策組']),
   'K-9 警視庁公安部公安第9課異能対策係');

// ケース4: echoQueriesなし（後方互換）は従来動作
detail = {
  title: 'X',
  associated: [{ title: '某中文題' }],
};
eq('後方互換: クエリ無しなら従来どおり先頭採用',
   ctx.pickJapaneseMangaUpdatesTitle_({}, detail),
   '某中文題');

// ケース5: エコーではない「別の中国語題」が先頭でも、かな入り日本語題を採る
// MangaUpdates 実データ（Our Sunny Days / series 53806507474）。
// 日出之家 は クエリ 日昇之屋 と包含関係が無いためエコー除外を通り抜ける。
detail = {
  title: 'Our Sunny Days',
  associated: [
    { title: 'When the Sun Rises' },
    { title: '日出之家' },
    { title: '日昇之屋' },
    { title: '陽が昇る家〜田舎で出会った俺たち〜' },
    { title: '해 뜨는 집' },
  ],
};
eq('非エコーの中国語題より、かな入り日本語題を優先',
   ctx.pickJapaneseMangaUpdatesTitle_({}, detail, ['日昇之屋']),
   '陽が昇る家〜田舎で出会った俺たち〜');

// ケース6: 同一題の表記違い（カナ読み別名）では、漢字の正式題を維持する。
// K-9 は Associated に「漢字正式題 → カナ読み別名」の順で並ぶため、
// 単純な「かな優先」だと正式題を差し置いてカナ別名を拾ってしまう。
detail = {
  title: 'K-9',
  associated: [
    { title: 'K-9 警視庁公安部公安第9課異能対策係' },
    { title: 'K-9~警視庁公安部公安第9課異能対策係~' },
    { title: 'ケーナイン　警視庁公安部公安第９課異能対策係' },
  ],
};
eq('同一題のカナ読み別名には乗り換えない',
   ctx.pickJapaneseMangaUpdatesTitle_({}, detail, ['K-9 警視廳公安部公安第9課異能對策組']),
   'K-9 警視庁公安部公安第9課異能対策係');

// ============================================================
// AniList 経路: native は「原語のタイトル」であって日本語とは限らない。
// 中国・韓国原作では native が中文題・ハングル題で、本物の日本語題は synonyms 側。
// ============================================================
const pickAni = ctx.pickAniListJapaneseFromMedia_;
if (typeof pickAni !== 'function') {
  console.error('[NG] pickAniListJapaneseFromMedia_ が未定義');
  failed++;
} else {
  // 実データ（AniList search=快把我哥带走）: media[0] は countryOfOrigin=CN
  eq('AniList: CN native はエコーなので採らず synonyms のかな題を採る',
     pickAni([{
       countryOfOrigin: 'CN',
       title: { native: '快把我哥带走', romaji: 'Kuai Ba Wo Ge Dai Zou', english: 'Please Take My Brother Away' },
       synonyms: ['Ani ni Tsukeru Kusuri wa Nai!', '兄に付ける薬はない！', 'Please Take My Brother Away'],
     }], ['快把我哥带走']),
     '兄に付ける薬はない！');

  eq('AniList: KR原作は synonyms のかな題を採る',
     pickAni([{
       countryOfOrigin: 'KR',
       title: { native: '日出之家', romaji: 'Our Sunny Days', english: 'Our Sunny Days' },
       synonyms: ['日昇之屋', '陽が昇る家〜田舎で出会った俺たち〜', '해 뜨는 집'],
     }], ['日昇之屋']),
     '陽が昇る家〜田舎で出会った俺たち〜');

  eq('AniList: かな入り native はそのまま採用',
     pickAni([{
       countryOfOrigin: 'JP',
       title: { native: 'アオのハコ', romaji: 'Ao no Hako', english: 'Blue Box' },
       synonyms: ['青春之箱'],
     }], ['青春之箱']),
     'アオのハコ');

  eq('AniList: K-9型はカナ読み別名に乗り換えない',
     pickAni([{
       countryOfOrigin: 'JP',
       title: { native: 'K-9 警視庁公安部公安第9課異能対策係', romaji: 'K-9', english: 'K-9' },
       synonyms: ['ケーナイン　警視庁公安部公安第９課異能対策係', 'K-9 警視廳公安部公安第9課異能對策組'],
     }], ['K-9 警視廳公安部公安第9課異能對策組']),
     'K-9 警視庁公安部公安第9課異能対策係');

  eq('AniList: 中文題しか無い作品は空文字',
     pickAni([{
       countryOfOrigin: 'CN',
       title: { native: '某中文漫画', romaji: 'Mou Chinese Manhua', english: '' },
       synonyms: ['某中文漫画別名'],
     }], ['某中文漫画']),
     '');

  eq('AniList: 無関係な作品は採用しない',
     pickAni([{
       countryOfOrigin: 'JP',
       title: { native: 'バトルファッカーB子', romaji: '', english: '' },
       synonyms: [],
     }], ['全球高考']),
     '');
}

// ============================================================
// pickBestCandidate_ : MUサイトスクレイプ／全プロバイダmiss時の候補列から選ぶ関数。
// 候補列は MU の Associated Names 掲載順そのままなので、中国語題が先に来る。
// ============================================================
const pickBest = ctx.pickBestCandidate_;
if (typeof pickBest !== 'function') {
  console.error('[NG] pickBestCandidate_ が未定義');
  failed++;
} else {
  // MU実データ（Our Sunny Days）の掲載順そのまま
  eq('候補列: 非エコーの中国語題より、かな入り日本語題を優先',
     pickBest(['Our Sunny Days', '日出之家', '日昇之屋', '陽が昇る家〜田舎で出会った俺たち〜'], ['日昇之屋']),
     '陽が昇る家〜田舎で出会った俺たち〜');

  eq('候補列: K-9型はカナ読み別名に乗り換えない',
     pickBest(['K-9 警視庁公安部公安第9課異能対策係', 'ケーナイン　警視庁公安部公安第９課異能対策係'],
              ['K-9 警視廳公安部公安第9課異能對策組']),
     'K-9 警視庁公安部公安第9課異能対策係');

  eq('候補列: クエリのエコーだけなら空文字',
     pickBest(['披著狼皮的羊公主', '披着狼皮的羊'], ['披著狼皮的羊公主']),
     '');

  eq('候補列: クエリ無し（後方互換）は従来どおり先頭の日本語シグナル候補',
     pickBest(['Our Sunny Days', '日出之家', '陽が昇る家〜田舎で出会った俺たち〜']),
     '陽が昇る家〜田舎で出会った俺たち〜');

  eq('候補列: 空なら空文字', pickBest([], ['なにか']), '');
}

// ============================================================
// pickMatchingMangaUpdatesSiteJapaneseTitle_ : MUサイト検索結果の題名列から選ぶ。
// クエリは台湾商品なので常に中国語。「クエリと一致する候補」を返す構造だと
// 原題エコーがそのまま日本語タイトルになる。
// ============================================================
const pickSite = ctx.pickMatchingMangaUpdatesSiteJapaneseTitle_;
if (typeof pickSite !== 'function') {
  console.error('[NG] pickMatchingMangaUpdatesSiteJapaneseTitle_ が未定義');
  failed++;
} else {
  const siteKeys = ctx.buildMangaUpdatesTitleKeys_('日昇之屋', '台湾', 'まんが', '日昇之屋');

  eq('サイト検索: クエリのエコーを日本語タイトルにしない',
     pickSite(['日昇之屋'], siteKeys, '台湾', 'まんが', '日昇之屋'),
     '');

  eq('サイト検索: かな入り日本語題があればそれを採る',
     pickSite(['日昇之屋', '日出之家', '陽が昇る家〜田舎で出会った俺たち〜'],
              siteKeys, '台湾', 'まんが', '日昇之屋'),
     '陽が昇る家〜田舎で出会った俺たち〜');

  // 字体差でキー一致しない題は、この関数では従来から採用されない（''）。
  // 別経路（MU API の associated / 漢字重なり判定）で拾う設計なので、
  // 今回のガード追加でこの挙動が変わっていないことだけ確認する。
  const k9Keys = ctx.buildMangaUpdatesTitleKeys_(
    'K-9 警視廳公安部公安第9課異能對策組', '台湾', 'まんが', 'K-9 警視廳公安部公安第9課異能對策組');
  eq('サイト検索: 字体差でキー不一致なら従来どおり採用しない',
     pickSite(['K-9 警視庁公安部公安第9課異能対策係'],
              k9Keys, '台湾', 'まんが', 'K-9 警視廳公安部公安第9課異能對策組'),
     '');
}

// ============================================================
// validateClientJapaneseTitleLookup_ : 拡張機能が送ってきた照会結果をGAS側で検算する
// 唯一の関門。漢字重なり率だけで判定していると、中文題（原題と重なりが高い）を止められない。
// ============================================================
const validateClient = ctx.validateClientJapaneseTitleLookup_;
if (typeof validateClient !== 'function') {
  console.error('[NG] validateClientJapaneseTitleLookup_ が未定義');
  failed++;
} else {
  const rowData = { 原題タイトル: '日昇之屋', 商品名: '台湾版 まんが 日昇之屋' };
  const analysis = { extractedWorkTitle: '日昇之屋', normalizedSearchTitle: '日昇之屋' };

  eq('クライアント検算: 中文題は候補にかな入りがあるので却下',
     validateClient({
       japaneseTitle: '日出之家',
       provider: 'mangaUpdates',
       candidates: ['日出之家', '日昇之屋', '陽が昇る家〜田舎で出会った俺たち〜'],
     }, analysis, rowData),
     false);

  eq('クライアント検算: 正解のかな入り題は通す',
     validateClient({
       japaneseTitle: '陽が昇る家〜田舎で出会った俺たち〜',
       provider: 'mangaUpdates',
       candidates: ['日出之家', '日昇之屋', '陽が昇る家〜田舎で出会った俺たち〜'],
     }, analysis, rowData),
     true);

  eq('クライアント検算: かな入り候補が無ければ漢字のみでも通す(K-9型)',
     validateClient({
       japaneseTitle: 'K-9 警視庁公安部公安第9課異能対策係',
       provider: 'mangaUpdates',
       candidates: ['K-9 警視庁公安部公安第9課異能対策係'],
     },
     { extractedWorkTitle: 'K-9 警視廳公安部公安第9課異能對策組' },
     { 原題タイトル: 'K-9 警視廳公安部公安第9課異能對策組' }),
     true);

  eq('クライアント検算: 候補列が無ければ従来判定（後方互換）',
     validateClient({ japaneseTitle: '日出之家', provider: 'mangaUpdates' }, analysis, rowData),
     true);
}

process.exitCode = failed ? 1 : 0;
console.log(failed ? `\n${failed} failed` : '\nall passed');
