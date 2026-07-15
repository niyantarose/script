// 取り置き台帳の初回登録ワークフローに関する純粋ロジックのテスト
// 実行: GASアップスクリプト直下で node tests/project24_reservation_workflow.test.js
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
  'Project_24/取り置き計算.js',
  'Project_24/取り置き台帳.js'
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

test('部分在庫と希望日待ちの受注だけを初期候補にする', () => {
  const orders=[
    {ban:'101',code:'AAA',sku:'AAAb',qty:2,kbn:'取り寄せ'},
    {ban:'102',code:'BBB',sku:'BBBb',qty:1,kbn:'取り寄せ'},
    {ban:'103',code:'CCC',sku:'CCCb',qty:1,kbn:'取り寄せ'}
  ];
  const rows=context.取り置き_初期候補_(orders,new Set(['101']),new Set(['102']));
  assert.strictEqual(rows.length,2);
  assert.strictEqual(rows[0].取置ID,'INIT|101|AAA|AAAB');
  assert.strictEqual(rows[1].取置ID,'INIT|102|BBB|BBBB');
});

test('初回入力は空欄と0を除外し注文数量超過を止める', () => {
  const result=context.取り置き_初期確定計画_([
    {取置ID:'INIT|1',受注番号:'101',商品コード:'AAA',SKU:'AAAb',注文数量:2,現物取り置き数量:2},
    {取置ID:'INIT|2',受注番号:'102',商品コード:'BBB',SKU:'BBBb',注文数量:1,現物取り置き数量:0},
    {取置ID:'INIT|3',受注番号:'103',商品コード:'CCC',SKU:'CCCb',注文数量:1,現物取り置き数量:2}
  ],[],'2026-07-15 10:00:00');
  assert.ok(result.errors.some(e=>/101|102/.test(e))===false);
  assert.ok(result.errors.some(e=>/103/.test(e)));
  assert.strictEqual(result.rows.length,0, '1件でもエラーなら保存対象を返さない');
});

test('同じINITキーは追加せず目標数量へ洗い替える', () => {
  const existing=[{取置ID:'INIT|101|AAA|AAAB',状態:'取り置き中',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:1,取置元種別:'開始前在庫'}];
  const result=context.取り置き_初期確定計画_([
    {取置ID:'INIT|101|AAA|AAAB',受注番号:'101',商品コード:'AAA',SKU:'AAAb',注文数量:2,現物取り置き数量:2}
  ],existing,'2026-07-15 10:00:00');
  assert.strictEqual(JSON.stringify(result.errors),'[]');
  assert.strictEqual(result.rows.length,1);
  assert.strictEqual(result.rows[0].取り置き数量,2);
});

test('台帳I/Oと初回登録の公開入口を提供する', () => {
  [
    '取り置き台帳_読む_',
    '取り置き台帳_保存_',
    'EMS在庫移動台帳_読む_',
    'EMS在庫移動台帳_保存_',
    '取り置き初期登録を作成',
    '取り置き初期登録を確定'
  ].forEach(name=>assert.strictEqual(typeof context[name],'function',name));
});

test('受注共通メニューに初回登録の作成と確定を追加する', () => {
  const source=fs.readFileSync('Project_24/引当.js','utf8');
  assert.ok(source.includes(".addItem('📋 取り置き初期登録を作成', '取り置き初期登録を作成')"));
  assert.ok(source.includes(".addItem('✅ 取り置き初期登録を確定', '取り置き初期登録を確定')"));
});

process.exitCode = failures ? 1 : 0;
console.log(failures ? `FAILURES: ${failures}` : 'ALL TESTS PASSED');
