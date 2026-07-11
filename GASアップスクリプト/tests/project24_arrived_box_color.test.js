const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

const headers = ['ステータス列', 'EMS到着日', '商品コード', '数量', 'EMS番号', '注文番号'];
const rows = [
  ['到着済', '2026-07-10', 'MOFUN-AS-04', '4', 'EG049624664KR', '10117178:1'],
  ['在庫反映済み', '2026-07-09', 'TAROT10', '1', 'EG049624465KR', '10117173'],
  ['到着済', '2026-07-10', 'FAKE-01', '9', '棚卸20260710', '10117179']
];

const range = (values, displayValues = values) => ({
  getValues: () => values,
  getDisplayValues: () => displayValues
});
const sheet = {
  getLastRow: () => 9,
  getLastColumn: () => headers.length,
  getRange: row => row === 6 ? range([headers]) : range(rows)
};
const context = {
  console,
  Date,
  Set,
  SpreadsheetApp: {
    openById: () => ({ getSheetByName: () => sheet })
  }
};
vm.createContext(context);
vm.runInContext(fs.readFileSync('Project_24/引当.js', 'utf8'), context);
vm.runInContext(fs.readFileSync('Project_24/P列確定.js', 'utf8'), context);

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

test('到着済P列確定マップは実到着日とEMS番号を保持する', () => {
  const entry = context.P列確定マップ_()['10117178'][0];
  assert.strictEqual(entry.key, 'MOFUN-AS-04');
  assert.strictEqual(entry.qty, 1);
  assert.strictEqual(entry.arrival, '2026-07-10');
  assert.strictEqual(entry.ems, 'EG049624664KR');
  assert.strictEqual(context.P列確定マップ_()['10117173'], undefined);
  assert.strictEqual(context.P列確定マップ_()['10117179'], undefined);
});

test('EMS明細は実EMSだけを供給にし、棚卸と番号空欄は数量0にする', () => {
  const values = [
    ['状態', '到着日', '商品コード', '数量', 'EMS番号'],
    ['到着済', '2026-07-10', 'REAL-01', 2, 'EG049624664KR'],
    ['到着済', '2026-07-10', 'FAKE-01', 9, '棚卸20260710'],
    ['到着済', '2026-07-10', 'NOEMS-01', 4, '']
  ];
  const emv = { getDataRange: () => range(values) };
  const r = context.EMS明細_(emv);
  assert.strictEqual(r.rows.length, 3); // シート行との対応は維持する
  assert.strictEqual(r.rows[0][r.cols.コード], 'REAL-01');
  assert.strictEqual(r.rows[1][r.cols.コード], '');
  assert.strictEqual(r.rows[1][r.cols.数量], 0);
  assert.strictEqual(r.rows[2][r.cols.コード], '');
  assert.strictEqual(r.除外, 2);
});

test('EMS番号書戻しは在庫反映済みの実EMSを保持し棚卸だけ消す', () => {
  assert.strictEqual(context.EMS番号書戻し値_('EG049624664KR', { 入荷:true }), 'EG049624664KR');
  assert.strictEqual(context.EMS番号書戻し値_('棚卸20260710', { 入荷:true }), '');
  assert.strictEqual(context.EMS番号書戻し値_('EG049624664KR', { 入荷:false, 履歴成立:false }), '');
  assert.strictEqual(context.EMS番号書戻し値_('EG-OLD', { 入荷:true, 箱EMS:'EG-NEW' }), 'EG-NEW');
});

test('古い入荷日でも到着済P列で確定した行は今回便になる', () => {
  const line = {
    入荷: true,
    入荷日値: '2026-07-09',
    今回P: true,
    kbn: '取り寄せ'
  };
  assert.strictEqual(context.今回到着扱い_(line, () => false), true);
});

test('在庫反映済み履歴だけの行はラベンダーを維持する', () => {
  const cfg = { 色_グレー: 'gray', 色_水: 'aqua', 色_橙: 'orange', 色_黄: 'yellow', 色_着: 'lavender' };
  const line = {
    入荷: true,
    入荷日値: '2026-07-09',
    履歴成立: true,
    kbn: '取り寄せ',
    キャンセル: false
  };
  assert.strictEqual(context.引当行状態_(line, cfg, () => false).color, 'lavender');
});

test('ymd_はGoogleシリアル日(46212)を46212-01-01に誤読しない', () => {
  assert.strictEqual(context.ymd_(46212), '2026-07-09');
  assert.strictEqual(context.ymd_('46212'), '2026-07-09');
  assert.strictEqual(context.ymd_(46214), '2026-07-11');
  assert.strictEqual(context.ymd_('2026-07-09'), '2026-07-09');
  assert.strictEqual(context.ymd_('26/07/09(木)'), '2026-07-09');
});

if (failures) process.exit(1);
