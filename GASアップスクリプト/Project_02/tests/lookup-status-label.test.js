// buildJapaneseTitleLookupStatusLabel_ の検証。
// 書き込みパスで外部照会をスキップした回を「登録なし(全サイト)」と誤確定しないこと。
// 実行: node tests/lookup-status-label.test.js
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

// 書き込みパス: 辞書だけ確認して外部は全部skip → 未照会（再照会対象に残す）
eq('skip_during_writeは未照会',
   ctx.buildJapaneseTitleLookupStatusLabel_({
     checkedSources: ['titleAliasDictionary'],
     skippedSources: ['mangaUpdates', 'aniList'],
     failedSources: [],
     trace: ['titleAliasDictionary:miss', 'mangaUpdates:skipped(skip_during_write)', 'aniList:skipped(skip_during_write)'],
   }),
   '未照会');

// 全ソースを実際に確認してmiss → 登録なし(全サイト)
eq('全ソース確認済みは登録なし(全サイト)',
   ctx.buildJapaneseTitleLookupStatusLabel_({
     checkedSources: ['titleAliasDictionary', 'mangaUpdates', 'aniList'],
     skippedSources: [],
     failedSources: [],
     trace: ['titleAliasDictionary:miss', 'mangaUpdates:miss', 'aniList:miss'],
   }),
   '登録なし(全サイト)');

// 失敗ソースあり → 照会失敗(ソース名)
eq('失敗は照会失敗ラベル',
   ctx.buildJapaneseTitleLookupStatusLabel_({
     checkedSources: ['mangaUpdates'],
     skippedSources: [],
     failedSources: ['mangaUpdates'],
     trace: ['mangaUpdates:failed(MU API search 403)'],
   }),
   '照会失敗(mangaUpdates)');

// 何も確認していない → 登録なし
eq('未確認は登録なし',
   ctx.buildJapaneseTitleLookupStatusLabel_({ checkedSources: [], skippedSources: [], failedSources: [], trace: [] }),
   '登録なし');

process.exitCode = failed ? 1 : 0;
console.log(failed ? `\n${failed} failed` : '\nall passed');
