// bookCodeNextUnusedWorkId_ の共有ハイウォーターマーク（採番管理!B1）対応の検証。
// ScriptProperties（Webアプリ専用）とシート側 DocumentProperties が分裂していても、
// スプレッドシート上の共有セルを介して両系統が同じ最大値を見られること。
const fs = require('fs');
const path = require('path');
const vm = require('vm');

// ---- GASモック ----
const scriptProps = {};
let sharedCellValue = ''; // 採番管理!B1
let sheetExists = true;
let insertedSheet = false;

const fakeRange = {
  getDisplayValue: () => String(sharedCellValue),
  setValue: (v) => { sharedCellValue = String(v); return fakeRange; },
};
const fakeNumberingSheet = {
  getRange: () => fakeRange,
  hideSheet: () => {},
};
const fakeSS = {
  getSheetByName: (name) => (name === '採番管理' && sheetExists ? fakeNumberingSheet : null),
  insertSheet: (name) => {
    if (name !== '採番管理') throw new Error('unexpected sheet: ' + name);
    sheetExists = true;
    insertedSheet = true;
    return fakeNumberingSheet;
  },
};

const ctx = { console };
ctx.globalThis = ctx;
ctx.SpreadsheetApp = { flush: () => {} };
ctx.PropertiesService = {
  getScriptProperties: () => ({
    getProperty: (k) => (k in scriptProps ? scriptProps[k] : null),
    setProperty: (k, v) => { scriptProps[k] = String(v); },
  }),
};
ctx.LockService = {};
ctx.Logger = { log: () => {} };
ctx.Utilities = {};
ctx.CacheService = {};
ctx.UrlFetchApp = {};
ctx.MailApp = {};
ctx.HtmlService = {};
ctx.ScriptApp = {};
ctx.XmlService = {};
ctx.ContentService = {};
ctx.Session = {};
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

// runtime.usedIds をvm内Setで用意（シート走査をスキップ）
function makeRuntime(ids) {
  return vm.runInContext(`({ usedIds: new Set(${JSON.stringify(ids)}) })`, ctx);
}

// ケース1: 共有セルが空 → 走査max(165)+1。共有セルとpropsの両方に166が書き戻る
scriptProps['台湾書籍系_作品ID_ハイウォーター'] = '150';
sharedCellValue = '';
eq('セル空: max+1を採番', ctx.bookCodeNextUnusedWorkId_(fakeSS, makeRuntime(['0005', '0165'])), '0166');
eq('セル空: 共有セルへ書き戻し', sharedCellValue, '166');
eq('セル空: propsへも書き戻し', scriptProps['台湾書籍系_作品ID_ハイウォーター'], '166');

// ケース2: 共有セル(200)が走査max(165)とprops(150)より大きい
// → シート側が発行済みの200を尊重して0201（削除済み最上位IDの再発行を防ぐ）
scriptProps['台湾書籍系_作品ID_ハイウォーター'] = '150';
sharedCellValue = '200';
eq('セル優先: 共有最大+1を採番', ctx.bookCodeNextUnusedWorkId_(fakeSS, makeRuntime(['0005', '0165'])), '0201');
eq('セル優先: 共有セル更新', sharedCellValue, '201');

// ケース3: 採番管理シート未作成 → 自動作成して従来どおり採番
scriptProps['台湾書籍系_作品ID_ハイウォーター'] = '150';
sheetExists = false;
insertedSheet = false;
sharedCellValue = '';
eq('シート無し: 採番は継続', ctx.bookCodeNextUnusedWorkId_(fakeSS, makeRuntime(['0005', '0165'])), '0166');
eq('シート無し: シートを自動作成', insertedSheet, true);

// ケース4: 共有セルアクセスが例外でも採番は止まらない（従来動作に縮退）
scriptProps['台湾書籍系_作品ID_ハイウォーター'] = '170';
const origGetSheet = fakeSS.getSheetByName;
fakeSS.getSheetByName = () => { throw new Error('boom'); };
eq('セル例外: props最大+1に縮退', ctx.bookCodeNextUnusedWorkId_(fakeSS, makeRuntime(['0005', '0165'])), '0171');
fakeSS.getSheetByName = origGetSheet;

process.exitCode = failed ? 1 : 0;
console.log(failed ? `\n${failed} failed` : '\nall passed');
