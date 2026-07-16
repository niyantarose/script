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

process.exitCode = failed ? 1 : 0;
console.log(failed ? `\n${failed} failed` : '\nall passed');
