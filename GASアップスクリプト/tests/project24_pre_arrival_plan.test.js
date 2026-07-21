// 未着便のP列先行記入（計画対象）と受注明細「入荷予定」の純粋ロジックテスト
// 実行: GASアップスクリプト直下で node tests/project24_pre_arrival_plan.test.js
const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

const activeSpreadsheet = { getSheetByName: () => null };
const context = {
  console,
  Date,
  Set,
  SpreadsheetApp: {
    openById: () => ({ getSheetByName: () => null }),
    getActive: () => activeSpreadsheet
  },
  LockService: { getDocumentLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
  PropertiesService: { getDocumentProperties: () => ({ getProperty: () => null, setProperty: () => {}, deleteProperty: () => {} }) },
  Utilities: { formatDate: () => '' },
  Logger: { log: () => {} }
};
vm.createContext(context);
[
  'Project_24/引当.js',
  'Project_24/引当履歴.js',
  'Project_24/取り置き計算.js',
  'Project_24/取り置き台帳.js',
  'Project_24/消込台帳.js',
  'Project_24/P列自動記入.js',
  'Project_24/全件再計算.js'
].filter(fs.existsSync).forEach(f => vm.runInContext(fs.readFileSync(f, 'utf8'), context));

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

test('計画対象判定: 到着済/在庫反映済み以外(空欄含む)を未着扱いにする', () => {
  assert.strictEqual(context.P列計画対象EMS_('到着済'), false);
  assert.strictEqual(context.P列計画対象EMS_('在庫反映済み'), false);
  assert.strictEqual(context.P列計画対象EMS_('発送済み'), true);
  assert.strictEqual(context.P列計画対象EMS_(''), true);
  assert.strictEqual(context.P列計画対象EMS_(null), true);
});

test('行順: 到着済が先・未着が後・ブロックSKUと対象外行は除外', () => {
  const rows = [
    {i: 0, 計画: true},
    {i: 1, 対象: true},
    {i: 2, 対象: true, 全件再計算ブロック: true},
    {i: 3} // 在庫反映済みなど(どちらでもない)
  ];
  const ordered = context.P列計画行順_(rows);
  assert.deepStrictEqual(ordered.map(r => r.i), [1, 0]);
});

test('未着FIFO: 到着済で足りない古い注文が未着便へ流れ、到着済の割当は単独時と同じ', () => {
  const arrived = {ems: 'EG-A', code: 'AAA', qty: 1, pOriginal: '', arrival: '2026-07-18', 対象: true};
  const future = {ems: 'EG-F', code: 'AAA', qty: 2, pOriginal: '', arrival: '2026-07-25', 計画: true};
  const orders = () => ([
    {ban: '101', code: 'AAA', sku: 'AAAb', qty: 2, need: 2, date: new Date('2026-07-01'), keys: ['AAA'], row: 10},
    {ban: '102', code: 'AAA', sku: 'AAAb', qty: 1, need: 1, date: new Date('2026-07-02'), keys: ['AAA'], row: 11}
  ]);
  const both = context.P列計画_純計算_(context.P列計画行順_([arrived, future]), orders(), {}, {});
  assert.strictEqual(both.rows[0].nextP, '101'); // 到着済: 全量101
  assert.strictEqual(both.rows[1].nextP, '101:1, 102:1'); // 未着: 101の残り1+102の1
  const alone = context.P列計画_純計算_(context.P列計画行順_([Object.assign({}, arrived)]), orders(), {}, {});
  assert.strictEqual(alone.rows[0].nextP, '101'); // 未着を混ぜても到着済の結果は不変
});

test('昇格: 同じ行が未着→到着済に変わっても割当テキストは同値', () => {
  const orders = () => ([{ban: '101', code: 'AAA', sku: 'AAAb', qty: 2, need: 2, date: new Date('2026-07-01'), keys: ['AAA'], row: 10}]);
  const asPlan = context.P列計画_純計算_(context.P列計画行順_([{ems: 'EG-F', code: 'AAA', qty: 2, pOriginal: '', arrival: '2026-07-25', 計画: true}]), orders(), {}, {});
  const asArrived = context.P列計画_純計算_(context.P列計画行順_([{ems: 'EG-F', code: 'AAA', qty: 2, pOriginal: '', arrival: '2026-07-25', 対象: true}]), orders(), {}, {});
  assert.strictEqual(asPlan.rows[0].nextP, asArrived.rows[0].nextP);
  assert.strictEqual(asPlan.rows[0].nextP, '101');
});

test('入荷予定表記: 単便・複数便の到着日順・日付なし・3便以上はほか', () => {
  assert.strictEqual(context.入荷予定表記_([{arrival: '2026-07-23', ems: 'EG050049766KR'}]), '7/23(…9766)');
  assert.strictEqual(
    context.入荷予定表記_([
      {arrival: '2026-07-26', ems: 'EG050099011KR'},
      {arrival: '2026-07-23', ems: 'EG050049766KR'}
    ]),
    '7/23(…9766)+7/26(…9011)');
  assert.strictEqual(context.入荷予定表記_([{arrival: '', ems: 'EG050049766KR'}]), '(…9766)');
  assert.strictEqual(
    context.入荷予定表記_([
      {arrival: '2026-07-23', ems: 'EG050049766KR'},
      {arrival: '2026-07-26', ems: 'EG050099011KR'},
      {arrival: '2026-07-30', ems: 'EG050111222KR'}
    ]),
    '7/23(…9766)+7/26(…9011)ほか');
});

test('入荷予定マップ: 計画行のentriesだけを受注行へ引き当て、到着済行は含めない', () => {
  const planRows = [
    {計画: true, code: 'AAA', arrival: '2026-07-25', ems: 'EG050049766KR', entries: [{ban: '101', qty: 1}]},
    {対象: true, code: 'AAA', arrival: '2026-07-18', ems: 'EG049827401KR', entries: [{ban: '102', qty: 1}]}
  ];
  const lines = [
    {ban: '101', row: 10, keys: ['AAA']},
    {ban: '102', row: 11, keys: ['AAA']}
  ];
  const map = context.入荷予定マップ_(planRows, lines);
  assert.strictEqual(map[10], '7/25(…9766)');
  assert.strictEqual(map[11], undefined);
});

process.exitCode = failures ? 1 : 0;
console.log(failures ? `${failures} FAILED` : 'ALL PASS');
