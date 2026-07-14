// ダニエル便の余り推定(ダニエル余り集計_)の純粋ロジックのテスト
// 実行: GASアップスクリプト直下で node tests/project24_daniel_amari.test.js
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

const 集計 = (src) => context.ダニエル余り集計_(src);
const row = (rows, code) => rows.find(r => r.code === code);

test('基本: 供給 − 出荷(推定) − 引当済み − Yahoo反映 = 余り', () => {
  const rows = 集計({
    記録: [{ems:'EE111KR', code:'AAA1', qty:10}],
    大邱ems: [],
    出荷済: [{ban:'1001', code:'AAA1', sku:'AAA1b', qty:3, 入荷日:''}],
    受注: [{code:'AAA1', sku:'AAA1b', qty:2, ems:'EE111KR'}],
    反映: {'AAA1': 1}
  });
  const r = row(rows, 'AAA1');
  assert.strictEqual(r.供給, 10);
  assert.strictEqual(r.出荷済, 3);
  assert.strictEqual(r.引当済, 2);
  assert.strictEqual(r.Yahoo反映, 1);
  assert.strictEqual(r.余り, 4);
});

test('大邱の箱に紐づく出荷(入荷日が到着日と一致)はダニエル消費に数えない', () => {
  const rows = 集計({
    記録: [{ems:'EE111KR', code:'BBB2', qty:5}],
    大邱ems: [{code:'BBB2', st:'到着済', qty:4, arrival:'2026-07-12', ems:'EG000111KR'}],
    出荷済: [
      {ban:'2001', code:'BBB2', sku:'', qty:2, 入荷日:'2026-07-12'}, // 大邱の箱由来
      {ban:'2002', code:'BBB2', sku:'', qty:1, 入荷日:''}            // GoQ内完結出荷=ダニエル消費
    ],
    受注: [], 反映: null
  });
  const r = row(rows, 'BBB2');
  assert.strictEqual(r.出荷済, 1, '大邱由来の2個は除外される');
  assert.strictEqual(r.余り, 4);
});

test('b枝番SKUは基底コードへ丸めて突き合わせる', () => {
  const rows = 集計({
    記録: [{ems:'EE111KR', code:'CCC3', qty:4}],
    大邱ems: [],
    出荷済: [{ban:'3001', code:'韓国語 何か', sku:'CCC3b', qty:1, 入荷日:''}],
    受注: [], 反映: null
  });
  assert.strictEqual(row(rows, 'CCC3').出荷済, 1);
});

test('ダニエル便以外のEMS番号が付いた受注は引当済みに数えない', () => {
  const rows = 集計({
    記録: [{ems:'EE111KR', code:'DDD4', qty:3}],
    大邱ems: [],
    出荷済: [],
    受注: [
      {code:'DDD4', sku:'', qty:1, ems:'EG999999KR'}, // 大邱の箱の引当
      {code:'DDD4', sku:'', qty:1, ems:''}            // 未引当
    ],
    反映: null
  });
  const r = row(rows, 'DDD4');
  assert.strictEqual(r.引当済, 0);
  assert.strictEqual(r.余り, 3);
});

test('記録に無いコードの出荷・受注は無視(供給のあるコードだけ出る)', () => {
  const rows = 集計({
    記録: [{ems:'EE111KR', code:'EEE5', qty:2}],
    大邱ems: [],
    出荷済: [{ban:'5001', code:'ZZZ9', sku:'', qty:5, 入荷日:''}],
    受注: [{code:'ZZZ9', sku:'', qty:5, ems:'EE111KR'}],
    反映: null
  });
  assert.strictEqual(rows.length, 1);
  assert.strictEqual(rows[0].code, 'EEE5');
  assert.strictEqual(rows[0].余り, 2);
});

test('消費が供給を超えたら余りはマイナスで出す(要調査のサイン)', () => {
  const rows = 集計({
    記録: [{ems:'EE111KR', code:'FFF6', qty:1}],
    大邱ems: [],
    出荷済: [{ban:'6001', code:'FFF6', sku:'', qty:3, 入荷日:''}],
    受注: [], 反映: null
  });
  assert.strictEqual(row(rows, 'FFF6').余り, -2);
});

test('同じ受注番号の台帳重複行は最大数量の1件にまとめる(出荷済み重複排除_)', () => {
  const rows = 集計({
    記録: [{ems:'EE111KR', code:'GGG7', qty:10}],
    大邱ems: [],
    出荷済: [
      {ban:'7001', code:'GGG7', sku:'GGG7b', qty:2, 入荷日:''},
      {ban:'7001', code:'GGG7（1011700）', sku:'', qty:1, 入荷日:''} // 表記ゆれの重複
    ],
    受注: [], 反映: null
  });
  assert.strictEqual(row(rows, 'GGG7').出荷済, 2, '同一受注は最大数量の1件だけ');
});

test('並びは余りの多い順', () => {
  const rows = 集計({
    記録: [
      {ems:'EE111KR', code:'HHH1', qty:1},
      {ems:'EE111KR', code:'HHH2', qty:5}
    ],
    大邱ems: [], 出荷済: [], 受注: [], 反映: null
  });
  assert.strictEqual(rows[0].code, 'HHH2');
});

process.exitCode = failures ? 1 : 0;
console.log(failures ? `FAILURES: ${failures}` : 'ALL TESTS PASSED');
