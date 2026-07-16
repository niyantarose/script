// 台湾書籍系_次の未使用作品ID_ の共有ハイウォーターマーク（採番管理!B1）対応の検証。
// DocumentProperties（Project_01専用）と ScriptProperties（Project_02専用）が
// 分裂していても、シート上の共有セルを介して両系統が同じ最大値を見られること。
const fs = require('fs');
const path = require('path');
const vm = require('vm');

// ---- GASモック ----
const docProps = {};
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
ctx.SpreadsheetApp = {
  getActiveSpreadsheet: () => fakeSS,
  getActive: () => fakeSS,
  flush: () => {},
};
ctx.PropertiesService = {
  getDocumentProperties: () => ({
    getProperty: (k) => (k in docProps ? docProps[k] : null),
    setProperty: (k, v) => { docProps[k] = String(v); },
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
ctx.Session = {};
vm.createContext(ctx);
vm.runInContext(
  fs.readFileSync(path.join(__dirname, '..', '台湾書籍系_共通.js'), 'utf8'),
  ctx,
  { filename: '台湾書籍系_共通.js' }
);

let failed = 0;
function eq(name, actual, expected) {
  const a = JSON.stringify(actual), e = JSON.stringify(expected);
  if (a !== e) { failed++; console.error(`[NG] ${name}: expected=${e} actual=${a}`); }
  else { console.log(`[OK] ${name}`); }
}

// 使用済みIDキャッシュをvm内Setで直接セット（シート走査をスキップ）
function seedUsed(ids) {
  vm.runInContext(
    `_台湾書籍系_使用済みIDキャッシュ_ = new Set(${JSON.stringify(ids)});`,
    ctx
  );
}

// ケース1: 共有セルが空 → シート走査max(165)+1。共有セルとpropsの両方に166が書き戻る
seedUsed(['0005', '0165']);
docProps['台湾書籍系_作品ID_ハイウォーター'] = '150';
sharedCellValue = '';
eq('セル空: max+1を採番', ctx.台湾書籍系_次の未使用作品ID_(), '0166');
eq('セル空: 共有セルへ書き戻し', sharedCellValue, '166');
eq('セル空: propsへも書き戻し', docProps['台湾書籍系_作品ID_ハイウォーター'], '166');

// ケース2: 共有セル(200)が走査max(165)とprops(150)より大きい
// → Webアプリ側が発行済みの200を尊重して0201（削除済み最上位IDの再発行を防ぐ）
seedUsed(['0005', '0165']);
docProps['台湾書籍系_作品ID_ハイウォーター'] = '150';
sharedCellValue = '200';
eq('セル優先: 共有最大+1を採番', ctx.台湾書籍系_次の未使用作品ID_(), '0201');
eq('セル優先: 共有セル更新', sharedCellValue, '201');

// ケース3: 採番管理シート未作成 → 自動作成して従来どおり採番
seedUsed(['0005', '0165']);
docProps['台湾書籍系_作品ID_ハイウォーター'] = '150';
sheetExists = false;
insertedSheet = false;
sharedCellValue = '';
eq('シート無し: 採番は継続', ctx.台湾書籍系_次の未使用作品ID_(), '0166');
eq('シート無し: シートを自動作成', insertedSheet, true);

// ケース4: 共有セルアクセスが例外でも採番は止まらない（従来動作に縮退）
seedUsed(['0005', '0165']);
docProps['台湾書籍系_作品ID_ハイウォーター'] = '170';
const origGetSheet = fakeSS.getSheetByName;
fakeSS.getSheetByName = () => { throw new Error('boom'); };
eq('セル例外: props最大+1に縮退', ctx.台湾書籍系_次の未使用作品ID_(), '0171');
fakeSS.getSheetByName = origGetSheet;

// ケース5: 巻き戻し運用 — セルに正の値があればpropsより低くてもセルを正とする
// （「採番を巻き戻す」メニューがセルとpropsを下げた後、片方のprops残骸で戻りが無効化されない）
seedUsed(['0005', '0165']);
docProps['台湾書籍系_作品ID_ハイウォーター'] = '170';
sheetExists = true;
sharedCellValue = '167';
eq('巻き戻し: セル(167)がprops(170)より優先', ctx.台湾書籍系_次の未使用作品ID_(), '0168');

// ケース6: セルを走査max未満まで下げても、シート上の使用中IDは絶対に越えない
seedUsed(['0005', '0165']);
docProps['台湾書籍系_作品ID_ハイウォーター'] = '0';
sharedCellValue = '100';
eq('巻き戻し: 走査max(165)が下限', ctx.台湾書籍系_次の未使用作品ID_(), '0166');

process.exitCode = failed ? 1 : 0;
console.log(failed ? `\n${failed} failed` : '\nall passed');
