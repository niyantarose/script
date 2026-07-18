// 韓国系（_kyoutuu ライブラリ）の作品ID採番ハイウォーター対応の検証。
// 実行: node tests/works-id-highwater.test.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');

// ---- GASモック ----
const docProps = {};
let 採番管理データ = []; // [[Worksシート名, 値], ...] 2行目以降
let 採番管理セル値 = {}; // シート名 -> 値（setValue反映先）

function makeCellRange(name) {
  return {
    getDisplayValue: () => String(採番管理セル値[name] == null ? '' : 採番管理セル値[name]),
    setValue: (v) => { 採番管理セル値[name] = String(v); },
  };
}
const fake採番管理 = {
  getLastRow: () => 採番管理データ.length + 1,
  getRange: function (row, col, numRows, numCols) {
    if (typeof row === 'string') return makeCellRange('__header__');
    if (numRows != null) {
      // 名前一覧の読み出し (2,1,n,1)
      return { getDisplayValues: () => 採番管理データ.map(r => [r[0]]), setValue: () => {} };
    }
    // 単一セル: (i+2, 2) → i行目の値セル / (行,1) → 名前書き込み
    const idx = row - 2;
    if (col === 1) {
      return { setValue: (v) => { 採番管理データ[idx] = [String(v), '']; },
               getDisplayValue: () => (採番管理データ[idx] || [''])[0] };
    }
    const name = (採番管理データ[idx] || [''])[0];
    return makeCellRange(name);
  },
  hideSheet: () => {},
};
const fakeSS = {
  getSheetByName: (name) => (name === '採番管理' ? fake採番管理 : null),
  insertSheet: (name) => { if (name !== '採番管理') throw new Error('unexpected: ' + name); return fake採番管理; },
};

const ctx = { console };
ctx.globalThis = ctx;
ctx.SpreadsheetApp = {
  getActiveSpreadsheet: () => fakeSS,
  getActive: () => fakeSS,
  getUi: () => ({ alert: () => {} }),
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
ctx.Session = {};
vm.createContext(ctx);
vm.runInContext(
  fs.readFileSync(path.join(__dirname, '..', '★共通.js'), 'utf8'),
  ctx,
  { filename: '★共通.js' }
);

let failed = 0;
function eq(name, actual, expected) {
  const a = JSON.stringify(actual), e = JSON.stringify(expected);
  if (a !== e) { failed++; console.error(`[NG] ${name}: expected=${e} actual=${a}`); }
  else { console.log(`[OK] ${name}`); }
}

// Worksシートのフェイク（A=WorksKey, B=作品ID, C=日題, D=作者, E=原題, F=巻, ...）
function makeWorksSheet(rows) {
  return {
    getLastRow: () => rows.length + 1,
    getLastColumn: () => 10,
    getRange: (r, c, nr, nc) => ({ getValues: () => rows.slice(r - 2, r - 2 + nr) }),
  };
}

const cfg = { 作品シート名: 'Works（韓国マンガ）', 作品列数: 10 };

// ケース1: HWなし → 走査max
docProps['作品ID_ハイウォーター_Works（韓国マンガ）'] = '0';
採番管理データ = []; 採番管理セル値 = {};
let d = ctx.全作品データを読み込み(makeWorksSheet([
  ['key1', '0001', '作品A', '作者A', '原題A', '', '', '', ''],
  ['key2', '0005', '作品B', '作者B', '原題B', '', '', '', ''],
]), cfg);
eq('HWなし: maxIdは走査max', d.maxId, 5);

// ケース2: 上位行が削除されてもHW（props）が守る
docProps['作品ID_ハイウォーター_Works（韓国マンガ）'] = '9';
d = ctx.全作品データを読み込み(makeWorksSheet([
  ['key1', '0001', '作品A', '作者A', '原題A', '', '', '', ''],
]), cfg);
eq('props HW: maxIdが9に引き上がる', d.maxId, 9);

// ケース3: 採番管理セルの値も反映（別プロジェクトが発行した最大値）
docProps['作品ID_ハイウォーター_Works（韓国マンガ）'] = '0';
採番管理データ = [['Works（韓国マンガ）', '']];
採番管理セル値 = { 'Works（韓国マンガ）': '12' };
d = ctx.全作品データを読み込み(makeWorksSheet([
  ['key1', '0001', '作品A', '作者A', '原題A', '', '', '', ''],
]), cfg);
eq('セルHW: maxIdが12に引き上がる', d.maxId, 12);

// ケース4: Worksが空でもHWは効く（全行削除後の再発行防止＝過去事故の核心）
docProps['作品ID_ハイウォーター_Works（韓国マンガ）'] = '7';
採番管理データ = []; 採番管理セル値 = {};
d = ctx.全作品データを読み込み(makeWorksSheet([]), cfg);
eq('空Works: maxIdはHWの7', d.maxId, 7);

// ケース5: 更新関数で props とセルの両方に書き戻る
採番管理データ = []; 採番管理セル値 = {};
ctx.採番ハイウォーター更新(cfg, 13);
eq('更新: propsへ保存', docProps['作品ID_ハイウォーター_Works（韓国マンガ）'], '13');
eq('更新: セルへ保存', 採番管理セル値['Works（韓国マンガ）'], '13');
eq('読取: max(セル,props)', ctx.採番ハイウォーター読取(cfg), 13);

// ケース6: WorksID振り直しは封印（シートに一切触らず空の結果を返す）
let touched = false;
const renumberSheet = {
  getLastRow: () => 5,
  getRange: () => ({
    getValues: () => { touched = true; return []; },
    setValues: () => { touched = true; },
  }),
};
eq('振り直し封印: 空結果', ctx.WorksID振り直しを実行(renumberSheet, cfg), { 変更数: 0, 旧新マップ: {} });
eq('振り直し封印: シート不干渉', touched, false);

process.exitCode = failed ? 1 : 0;
console.log(failed ? `\n${failed} failed` : '\nall passed');
