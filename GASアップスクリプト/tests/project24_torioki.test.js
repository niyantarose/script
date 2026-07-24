// 取り置き（開始前の手元在庫）まわりの純粋ロジックのテスト
// 実行: GASアップスクリプト直下で node tests/project24_torioki.test.js
// 背景: 取り置きで足りている注文が今回便から再引当され、取り置きから出た出荷済みが
//       今回の箱を推測消費して「帳尻は合うが引当先が違う」状態になる問題への対策。
const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

const context = {
  console,
  Date,
  Set,
  SpreadsheetApp: { openById: () => ({ getSheetByName: () => null }) },
  Utilities: { formatDate: () => '' },
  Logger: { log: () => {} }
};
vm.createContext(context);
[
  'Project_24/引当.js',
  'Project_24/P列自動記入.js',
  'Project_24/消込台帳.js',
  'Project_24/ダニエル余り.js'
].forEach(f => vm.runInContext(fs.readFileSync(f, 'utf8'), context));

let failures = 0;
function test(name, fn) {
  try {
    fn();
    console.log(`PASS ${name}`);
  } catch (error) {
    failures++;
    console.error(`FAIL ${name}: ${error.message}`);
  }
}

// ===== 残必要計算_: 現在の取り置き中数量だけを差し引く =====

test('取り置き中数量が注文数を満たす行は今回便からの必要数0', () => {
  assert.strictEqual(context.残必要計算_({qty:1, 取り置き中数量:1}), 0);
});

test('取り置き中数量が一部だけなら残りだけ必要', () => {
  assert.strictEqual(context.残必要計算_({qty:3, 取り置き中数量:1}), 2);
});

test('現在台帳の取り置きなしでは旧履歴を必要数から差し引かない', () => {
  assert.strictEqual(context.残必要計算_({qty:2}), 2);
  assert.strictEqual(context.残必要計算_({qty:2, alloc:1}), 1);
  assert.strictEqual(context.残必要計算_({qty:2, alloc:1, 履歴Alloc:1}), 1);
});

test('引当・取り置きの合計が注文数を超えても負にならない', () => {
  assert.strictEqual(context.残必要計算_({qty:1, alloc:1, 取り置き中数量:1}), 0);
});

// ===== P列需要: 行単位で発送済みの分割を除外する =====

test('P列需要は発送済みの分割行を除外し未発送行だけ残す', () => {
  const M = { 出荷日: 0, 出荷日毎: 1 };
  assert.strictEqual(context.P列需要対象行_(['', ''], M), true);
  assert.strictEqual(context.P列需要対象行_(['', '2026/07/14'], M), false);
  assert.strictEqual(context.P列需要対象行_(['2026-07-14', ''], M), false);
});

// ===== 取り置き出荷_: 台帳メモによる人為オーバーライド =====

test('台帳メモに「取り置き」がある出荷済み行を判定できる', () => {
  assert.strictEqual(context.取り置き出荷_({メモ:'取り置き'}), true);
  assert.strictEqual(context.取り置き出荷_({メモ:'取置分から出荷'}), true);
  assert.strictEqual(context.取り置き出荷_({メモ:'取り置きから'}), true);
  assert.strictEqual(context.取り置き出荷_({メモ:''}), false);
  assert.strictEqual(context.取り置き出荷_({メモ:'通常出荷'}), false);
  assert.strictEqual(context.取り置き出荷_({}), false);
});

// ===== ダニエル余り: 取り置き出荷はダニエル便も消費しない =====

test('取り置きメモ付きの出荷はダニエル余りからも差し引かない', () => {
  const rows = context.ダニエル余り集計_({
    記録: [{ems:'EE111KR', code:'AAA1', qty:5}],
    大邱ems: [],
    出荷済: [
      {ban:'1001', code:'AAA1', sku:'', qty:2, 入荷日:'', メモ:'取り置き'},
      {ban:'1002', code:'AAA1', sku:'', qty:1, 入荷日:'', メモ:''}
    ],
    受注: [], 反映: null
  });
  const r = rows.find(x => x.code === 'AAA1');
  assert.strictEqual(r.出荷済, 1, '取り置き出荷2個は除外され通常出荷1個だけ');
  assert.strictEqual(r.余り, 4);
});

process.exitCode = failures ? 1 : 0;
console.log(failures ? `FAILURES: ${failures}` : 'ALL TESTS PASSED');
