const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

const headers = ['ステータス列', 'EMS到着日', '商品コード', '数量', 'EMS番号', '注文番号'];
const rows = [
  ['到着済', '2026-07-10', 'MOFUN-AS-04', '4', 'EG049624664KR', '10117178:1'],
  ['在庫反映済み', '2026-07-09', 'TAROT10', '1', 'EG049624465KR', '10117173']
];

const range = (values, displayValues = values) => ({
  getValues: () => values,
  getDisplayValues: () => displayValues
});
const sheet = {
  getLastRow: () => 8,
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

test('旧処理が作った西暦46213年の入荷日をシリアル46213として2026-07-10へ戻す', () => {
  const broken = new Date(0);
  broken.setFullYear(46213, 0, 1);
  broken.setHours(0, 0, 0, 0);

  assert.strictEqual(context.ymd_(broken), '2026-07-10');
  const corrected = context.入荷日シート値補正_(broken);
  assert.ok(corrected instanceof Date);
  assert.strictEqual(context.ymd_(corrected), '2026-07-10');
});

if (failures) process.exit(1);
