// 確定発行バッチ文脈キャッシュの検証（GAS非依存・node実行）。
// 実行: node tests/kakutei-batch-cache.test.js
//
// 対象:
//   _TWBK_確定Ctx設定_ / _TWBK_確定Ctxクリア_（モジュール文脈の set/clear）
//   台湾書籍系_Worksから取得_            … 文脈の works を消費（点3）
//   台湾書籍系_同一作品行から情報を取得_  … 文脈の sheetValues を消費（点2）
//   台湾書籍系_Worksを取得または作成_    … 新規Works作成時に文脈へ push（点4）
//
// 計画書「テスト」節の観点:
//   (a) 同一新規作品の2行 → 同一作品ID・Works追記1回
//   (b) 別媒体同名(CM/NV) → 別作品ID（媒体分離維持）
//   (c) 文脈OFF → 従来経路（シート読み）で同結果
//   (d) sheetValues 差し替え → 同一作品の後続行が最新IDを引ける（ID一貫性の要）
//
// 純関数の抽出はこのコードでは非現実的（GAS APIに密結合）なため、
// 既存 works-id-numbering.test.js と同じ vm+GASモック方式で「文脈キャッシュの整合」を
// 実コードのまま単体検証する（その旨はレポートに明記）。
const fs = require('fs');
const path = require('path');
const vm = require('vm');

// ---- 汎用フェイクシート（2D grid、1行目=ヘッダー）----
function makeFakeSheet(name, header) {
  const grid = [header.slice()]; // grid[r-1] = 行r（1始まり）
  return {
    _grid: grid,
    getName: () => name,
    getLastRow: () => grid.length,
    getLastColumn: () => header.length,
    getMaxRows: () => Math.max(grid.length, 2),
    getRange: (r, c, nr = 1, nc = 1) => ({
      setNumberFormat: () => {},
      setValue: (v) => {
        if (!grid[r - 1]) grid[r - 1] = new Array(header.length).fill('');
        grid[r - 1][c - 1] = v;
      },
      setValues: (vals) => {
        for (let i = 0; i < vals.length; i++) {
          const rr = r - 1 + i;
          if (!grid[rr]) grid[rr] = new Array(header.length).fill('');
          for (let j = 0; j < vals[i].length; j++) grid[rr][c - 1 + j] = vals[i][j];
        }
      },
      getValues: () => {
        const out = [];
        for (let i = 0; i < nr; i++) {
          const rr = r - 1 + i;
          const rowArr = grid[rr] || new Array(header.length).fill('');
          out.push(rowArr.slice(c - 1, c - 1 + nc));
        }
        return out;
      },
      getDisplayValues: function () {
        return this.getValues().map(row => row.map(v => String(v == null ? '' : v)));
      },
    }),
  };
}

const WORKS_HEADER = ['WorksKey', '作品ID', '日本語タイトル', '作者', '原題タイトル', '登録済み巻', '最新巻', '更新日時', '最新巻(予約込み)', '予約更新日時'];
const 設定 = { 作品シート名: 'Works（書籍専用）', 作品列数: 10, 作品ヘッダー: WORKS_HEADER.slice(), マスターシート名: undefined };

// ---- GASモック ----
let worksSheet = makeFakeSheet(設定.作品シート名, WORKS_HEADER);
let 採番Shared = ''; // 採番管理!B1（共有ハイウォーター）
let 採番Pool = '';   // 採番管理!B2（解放プール）
const docProps = {};
const fakeSharedRange = { getDisplayValue: () => String(採番Shared), setValue: (v) => { 採番Shared = String(v); return fakeSharedRange; } };
const fakePoolRange = { getDisplayValue: () => String(採番Pool), setValue: (v) => { 採番Pool = String(v); return fakePoolRange; } };
const fakeNumbering = { getRange: (...a) => (a[0] === 'B2' ? fakePoolRange : fakeSharedRange), hideSheet: () => {} };
const fakeSS = {
  getSheetByName: (n) => (n === 設定.作品シート名 ? worksSheet : (n === '採番管理' ? fakeNumbering : null)),
  insertSheet: (n) => { if (n === '採番管理') return fakeNumbering; throw new Error('unexpected insertSheet ' + n); },
};

const ctx = { console };
ctx.globalThis = ctx;
ctx.SpreadsheetApp = { getActive: () => fakeSS, getActiveSpreadsheet: () => fakeSS, flush: () => {}, getUi: () => ({ alert: () => {} }) };
ctx.PropertiesService = {
  getDocumentProperties: () => ({
    getProperty: (k) => (k in docProps ? docProps[k] : null),
    setProperty: (k, v) => { docProps[k] = String(v); },
  }),
};
ctx.LockService = { getDocumentLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) };
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

// カテゴリ／媒体コード集合キャッシュを先に温める（以降マスター読込は走らない＝挙動安定）。
ctx.台湾書籍系_媒体コード集合_();
ctx.台湾書籍系_カテゴリ名前マップ_();

// scenario間のモジュール状態リセット（汚染防止）。
function resetState(usedIds) {
  vm.runInContext(
    `_台湾書籍系_作品キャッシュ = null; _台湾書籍系_ID列書式済み_ = {}; _台湾書籍系_使用済みIDキャッシュ_ = new Set(${JSON.stringify(usedIds || [])});`,
    ctx
  );
  ctx._TWBK_確定Ctxクリア_();
  worksSheet = makeFakeSheet(設定.作品シート名, WORKS_HEADER); // fresh Works grid（fakeSSはクロージャで最新を参照）
  採番Shared = '';
  採番Pool = '';
  for (const k of Object.keys(docProps)) delete docProps[k];
}

const worksCol = {};
WORKS_HEADER.forEach((h, i) => { if (h) worksCol[h] = i; });
const CMRow = (id) => ['原題||CM||劍鬼花', id, '雪雲の花', 'Snob', '劍鬼花', '', '', '', '', ''];

/* ============================================================
 * 点3: 台湾書籍系_Worksから取得_ が文脈のworksを消費し、シートより優先する。媒体分離も守る。
 * ============================================================ */
resetState([]);
// シートには別ID(9999)を置く。文脈が使われれば 0300、シートが誤って読まれれば 9999。
worksSheet.getRange(2, 1, 1, 10).setValues([CMRow('9999')]);
ctx._TWBK_確定Ctx設定_({
  sheetName: '台湾まんが', lastCol: 7, sheetValues: [],
  worksSheetName: 設定.作品シート名, worksCol, works: [CMRow('0300')],
});
eq('点3: 文脈のworksを優先(シートの9999でなく0300)',
   (ctx.台湾書籍系_Worksから取得_(設定, '劍鬼花', '雪雲の花', 'Snob', 'CM') || {}).作品ID, '0300');
eq('点3: 別媒体NVはCM行にマッチせずnull(媒体分離)',
   ctx.台湾書籍系_Worksから取得_(設定, '劍鬼花', '雪雲の花', 'Snob', 'NV'), null);

/* ============================================================
 * 点c: 文脈OFF → 従来経路（Worksシート読み）で同結果。
 * ============================================================ */
resetState([]);
ctx._TWBK_確定Ctxクリア_();
worksSheet.getRange(2, 1, 1, 10).setValues([CMRow('0300')]);
eq('点c: 文脈OFFでもシート読みで作品IDを取得',
   (ctx.台湾書籍系_Worksから取得_(設定, '劍鬼花', '雪雲の花', 'Snob', 'CM') || {}).作品ID, '0300');

/* ============================================================
 * 点2 & 点d: 台湾書籍系_同一作品行から情報を取得_ が文脈のsheetValuesを消費し、
 *   兄弟行への「差し替え」を反映（＝行N採番後に行N+1が同じIDを引ける）。媒体分離も守る。
 * ============================================================ */
resetState([]);
const shProd = { getName: () => '台湾まんが' };
const 商品列 = { '原題タイトル': 1, '作者': 2, '日本語タイトル': 3, 'カテゴリ': 4, '作品ID(W)（自動）': 5, '親コード': 6, 'SKU(自動)': 7 };
const 商品実列名 = { 原題: '原題タイトル', 作者: '作者', 日本語タイトル: '日本語タイトル', カテゴリ: 'カテゴリ', 作品ID: '作品ID(W)（自動）', 商品コード: '親コード', SKU自動: 'SKU(自動)', 言語: '', 原題商品タイトル: '' };
// sheetValues[0]=行2(兄弟), [1]=行3(現在行)。列は 商品列-1 の0始まりで格納。
const sheetVals = [
  ['劍鬼花', 'Snob', '雪雲の花', 'まんが', '', '', ''], // 行2: まだID未採番
  ['劍鬼花', 'Snob', '雪雲の花', 'まんが', '', '', ''], // 行3: 現在行
];
ctx._TWBK_確定Ctx設定_({
  sheetName: '台湾まんが', lastCol: 7, sheetValues: sheetVals,
  worksSheetName: 設定.作品シート名, worksCol: {}, works: [],
});
const 値現在 = { 原題タイトル: '劍鬼花', 日本語タイトル: '雪雲の花', 作者: 'Snob', カテゴリ: 'まんが', 作品ID: '' };

const r未採番 = ctx.台湾書籍系_同一作品行から情報を取得_(shProd, 3, 商品列, 値現在, 商品実列名);
eq('点d: 兄弟行がID未採番なら作品IDは空', r未採番 && r未採番.作品ID, '');

sheetVals[0][4] = '0300'; // ← バッチの「行処理後 sheetValues[row-2] 差し替え」を模擬（兄弟行2に採番）
const r採番後 = ctx.台湾書籍系_同一作品行から情報を取得_(shProd, 3, 商品列, 値現在, 商品実列名);
eq('点d: 差し替え後は兄弟行の最新IDを引く', r採番後 && r採番後.作品ID, '0300');

sheetVals[0][3] = '小説'; // 兄弟行2を別媒体(NV)へ。現在行はまんが(CM)。
const r別媒体 = ctx.台湾書籍系_同一作品行から情報を取得_(shProd, 3, 商品列, 値現在, 商品実列名);
eq('点2: 別媒体の同名兄弟行は無視(媒体分離)', r別媒体, null);

/* ============================================================
 * 点a & 点b: 台湾書籍系_Worksを取得または作成_ を実コードで駆動（点3+点4を統合検証）。
 *   同一新規作品を2回 → 同一ID・追記1回 / 別媒体同名 → 別ID。
 * ============================================================ */
resetState(['0199']); // 使用済み最大199 → 次採番は0200
const 文脈I = {
  sheetName: '台湾まんが', lastCol: 7, sheetValues: [],
  worksSheetName: 設定.作品シート名, worksCol: Object.assign({}, worksCol), works: [],
};
ctx._TWBK_確定Ctx設定_(文脈I);

const 値CM1 = { 原題タイトル: '劍鬼花', 日本語タイトル: '雪雲の花', 作者: 'Snob', カテゴリ: 'まんが' };
const c1 = ctx.台湾書籍系_Worksを取得または作成_(設定, 値CM1);
eq('点a: 新規作品に0200採番', c1 && c1.作品ID, '0200');
eq('点a: 文脈worksへ1回追記', 文脈I.works.length, 1);
eq('点a: Worksシートも1行追記(header+1)', worksSheet._grid.length, 2);

const 値CM2 = { 原題タイトル: '劍鬼花', 日本語タイトル: '雪雲の花', 作者: 'Snob', カテゴリ: 'まんが' };
const c2 = ctx.台湾書籍系_Worksを取得または作成_(設定, 値CM2);
eq('点a: 同一新規作品の2回目は同一0200', c2 && c2.作品ID, '0200');
eq('点a: 2回目は文脈worksへ追記しない', 文脈I.works.length, 1);
eq('点a: 2回目はシートへも追記しない', worksSheet._grid.length, 2);

const 値NV1 = { 原題タイトル: '劍鬼花', 日本語タイトル: '雪雲の花', 作者: 'Snob', カテゴリ: '小説' };
const c3 = ctx.台湾書籍系_Worksを取得または作成_(設定, 値NV1);
eq('点b: 別媒体NV同名は別ID(0201)', c3 && c3.作品ID, '0201');
eq('点b: NVは新規なので文脈worksへ追記', 文脈I.works.length, 2);
eq('点b: NVはシートへも追記(header+2)', worksSheet._grid.length, 3);

/* ============================================================
 * set/clear: クリア後は文脈OFF＝従来経路（シート読み）へ戻る。
 * ============================================================ */
resetState([]);
worksSheet.getRange(2, 1, 1, 10).setValues([CMRow('0300')]);
ctx._TWBK_確定Ctx設定_({
  sheetName: '台湾まんが', lastCol: 7, sheetValues: [],
  worksSheetName: 設定.作品シート名, worksCol, works: [CMRow('7777')], // 文脈は7777
});
eq('set: 文脈ONは文脈値(7777)',
   (ctx.台湾書籍系_Worksから取得_(設定, '劍鬼花', '雪雲の花', 'Snob', 'CM') || {}).作品ID, '7777');
ctx._TWBK_確定Ctxクリア_();
eq('clear: 文脈OFFはシート値(0300)へ復帰',
   (ctx.台湾書籍系_Worksから取得_(設定, '劍鬼花', '雪雲の花', 'Snob', 'CM') || {}).作品ID, '0300');

process.exitCode = failed ? 1 : 0;
console.log(failed ? `\n${failed} failed` : '\nall passed');
