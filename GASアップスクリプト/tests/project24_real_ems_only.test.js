const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

const headers = ['ステータス列', 'EMS到着日', '商品コード', '数量', 'EMS番号', '注文番号'];
const pRows = [
  ['到着済', '2026-07-10', 'OFW304-1', '1', 'EG000000001KR', '10117001'],
  ['到着済', '2026-07-10', 'OFW304-2', '1', '棚卸20260710', '10117002'],
  ['到着済', '2026-07-10', 'OFW305-1', '2', '', '10117003'],
  ['在庫反映済み', '2026-07-09', 'OFW305-2', '2', 'EG000000002KR', '10117004']
];

const range = (values, displayValues = values) => ({
  getValues: () => values,
  getDisplayValues: () => displayValues
});

const externalSheet = {
  getLastRow: () => 6 + pRows.length,
  getLastColumn: () => headers.length,
  getRange: row => row === 6 ? range([headers]) : range(pRows)
};

const context = {
  console,
  Date,
  Set,
  SpreadsheetApp: {
    openById: () => ({ getSheetByName: () => externalSheet })
  }
};

vm.createContext(context);
[
  'Project_24/引当.js',
  'Project_24/P列確定.js',
  'Project_24/消込台帳.js',
  'Project_24/全件検算.js',
  'Project_24/入荷日チェック.js'
].forEach(file => vm.runInContext(fs.readFileSync(file, 'utf8'), context));

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

test('OFW304/305の数値枝番は4つの別コードとして残る', () => {
  const codes = ['OFW304-1', 'OFW304-2', 'OFW305-1', 'OFW305-2'];
  const normalized = codes.map(code => context.normCode_(code));
  assert.strictEqual(new Set(normalized).size, 4);
  codes.forEach(code => {
    assert.strictEqual(JSON.stringify(context.codeKeys_(code)), JSON.stringify([code]));
  });
  assert.strictEqual(context.normCode_('OFW305-1OFW304-2'), 'OFW305-1OFW304-2');
});

test('実EMS番号だけを有効にする', () => {
  assert.strictEqual(context.実EMS番号_('EG000000001KR'), true);
  assert.strictEqual(context.実EMS番号_('棚卸20260710'), false);
  assert.strictEqual(context.実EMS番号_(''), false);
  assert.strictEqual(context.実EMS番号_('   '), false);
});

test('EMS明細は行位置を保ったまま棚卸と番号空欄を数量0にする', () => {
  const values = [
    ['状態', '商品コード', '数量', 'EMS番号', 'EMS到着日'],
    ['到着済', 'OFW304-1', 1, 'EG000000001KR', '2026-07-10'],
    ['到着済', 'OFW304-2', 1, '棚卸20260710', '2026-07-10'],
    ['到着済', 'OFW305-1', 2, '', '2026-07-10']
  ];
  const sheet = { getDataRange: () => range(values) };
  const result = context.EMS明細_(sheet);
  assert.strictEqual(result.rows.length, 3);
  assert.strictEqual(result.rows[0][result.cols.コード], 'OFW304-1');
  assert.strictEqual(result.rows[0][result.cols.数量], 1);
  assert.strictEqual(result.rows[1][result.cols.コード], '');
  assert.strictEqual(result.rows[1][result.cols.数量], 0);
  assert.strictEqual(result.rows[2][result.cols.コード], '');
  assert.strictEqual(result.rows[2][result.cols.数量], 0);
  assert.strictEqual(result.除外, 2);
});

test('P列確定は到着済み実EMSだけを採用する', () => {
  const map = context.P列確定マップ_();
  assert.strictEqual(map['10117001'][0].key, 'OFW304-1');
  assert.strictEqual(map['10117002'], undefined);
  assert.strictEqual(map['10117003'], undefined);
  assert.strictEqual(map['10117004'], undefined);
});

test('全件検算は棚卸と番号空欄を供給へ加えない', () => {
  const result = context.全件検算_集計_({
    ems: [
      { code: 'OFW304-1', st: '到着済', qty: 1, arrival: '2026-07-10', ems: 'EG000000001KR' },
      { code: 'OFW304-1', st: '到着済', qty: 1, arrival: '2026-07-10', ems: '棚卸20260710' },
      { code: 'OFW305-1', st: '到着済', qty: 2, arrival: '2026-07-10', ems: '' }
    ],
    出荷済: [],
    受注: [],
    a在庫: null
  });
  assert.strictEqual(result.rows.length, 1);
  assert.strictEqual(result.rows[0].code, 'OFW304-1');
  assert.strictEqual(result.rows[0].到着済, 1);
});

test('EMS番号書戻しは実EMSを保持し棚卸番号を消す', () => {
  assert.strictEqual(context.EMS番号書戻し値_('EG000000001KR', { 入荷: true }), 'EG000000001KR');
  assert.strictEqual(context.EMS番号書戻し値_('棚卸20260710', { 入荷: true }), '');
  assert.strictEqual(context.EMS番号書戻し値_('', { 箱EMS: 'EG000000002KR' }), 'EG000000002KR');
});

test('P列・個別対応・履歴・入荷日チェックも実EMS判定へ接続されている', () => {
  const checks = [
    ['Project_24/P列自動記入.js', /実EMS番号_\(ev\[i\]\[0\]\)/],
    ['Project_24/P列確定.js', /実EMS番号_\(r\[cEms\]\)/],
    ['Project_24/個別対応.js', /実EMS番号_\(r\.EMS番号\)/],
    ['Project_24/引当履歴.js', /実EMS番号_\(r\[c\.EMS番号\]\)/],
    ['Project_24/入荷日チェック.js', /実EMS番号_\(r\[cE\]\)/],
    ['Project_24/全件検算.js', /実EMS番号_\(r\.ems\)/]
  ];
  checks.forEach(([file, pattern]) => {
    const source = fs.readFileSync(file, 'utf8');
    assert.ok(pattern.test(source), `${file} が実EMS番号_を通っていない`);
  });
});

test('旧棚卸解除は韓国ルートだけ入荷日を消し実EMSを保持する', () => {
  assert.strictEqual(
    JSON.stringify(context.旧棚卸割当解除値_('棚卸20260710', '2026-07-10', false)),
    JSON.stringify({ 対象: true, EMS番号: '', 入荷日: '' })
  );
  assert.strictEqual(
    JSON.stringify(context.旧棚卸割当解除値_('棚卸20260710', '2026-07-10', true)),
    JSON.stringify({ 対象: true, EMS番号: '', 入荷日: '2026-07-10' })
  );
  assert.strictEqual(
    JSON.stringify(context.旧棚卸割当解除値_('EG000000001KR', '2026-07-10', false)),
    JSON.stringify({ 対象: false, EMS番号: 'EG000000001KR', 入荷日: '2026-07-10' })
  );
});

test('旧棚卸解除後はP列書き直しから手動で進める', () => {
  assert.strictEqual(
    JSON.stringify(context.旧棚卸解除後手順_()),
    JSON.stringify(['P列を書き直す', 'EMS在庫を更新', '②引き当て実行', '全件検算レポート'])
  );
  const source=fs.readFileSync('Project_24/入荷日チェック.js','utf8');
  const start=source.indexOf('function 旧棚卸割当だけを解除して再引当本体_');
  const end=source.indexOf('// ===== 引当データの全リセット', start);
  const body=source.slice(start, end);
  assert.strictEqual(body.includes('引当実行_本体_();'), false);
});

if (failures) process.exit(1);
