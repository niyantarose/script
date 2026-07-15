// 取り置き台帳の純粋ドメインロジックのテスト
// 実行: GASアップスクリプト直下で node tests/project24_reservation_domain.test.js
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
  'Project_24/取り置き計算.js'
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

test('数値枝番は別商品、販売末尾a/bだけを基底へ寄せる', () => {
  const orders = [
    {ban:'1', code:'OFW304-1', sku:'OFW304-1b'},
    {ban:'1', code:'OFW304-2', sku:'OFW304-2b'},
    {ban:'1', code:'OFW305-1', sku:'OFW305-1b'},
    {ban:'1', code:'OFW305-2', sku:'OFW305-2b'}
  ];
  const codes = orders.map(o => context.取り置き_商品コード_(o.sku, o.code));
  assert.strictEqual(JSON.stringify(codes), JSON.stringify(['OFW304-1','OFW304-2','OFW305-1','OFW305-2']));
  assert.strictEqual(new Set(codes).size, 4);
  assert.strictEqual(context.取り置き_商品コード_('', 'POEM65（10116569）'), 'POEM65');
});

test('取り置き中だけが注文必要数を減らす', () => {
  const key = context.取り置き_行キー_({ban:'101', code:'AAA-1', sku:'AAA-1b'});
  const rows = [
    {取置ID:'A', 状態:'取り置き中', 受注番号:'101', 商品コード:'AAA-1', SKU:'AAA-1b', 取り置き数量:2, 取置元種別:'開始前在庫'},
    {取置ID:'B', 状態:'発送済み', 受注番号:'101', 商品コード:'AAA-1', SKU:'AAA-1b', 取り置き数量:1, 取置元種別:'EMS', 元EMS番号:'EG1'},
    {取置ID:'C', 状態:'手動解除', 受注番号:'101', 商品コード:'AAA-1', SKU:'AAA-1b', 取り置き数量:1, 取置元種別:'開始前在庫'}
  ];
  const summary = context.取り置き_集計_(rows, []);
  assert.strictEqual(summary.activeByKey[key], 2);
  assert.strictEqual(context.取り置き_今回必要数_({ban:'101',code:'AAA-1',sku:'AAA-1b',qty:3}, summary), 1);
});

test('キャンセル戻し結果を供給使用区分へ正しく分類する', () => {
  const base = {状態:'キャンセル戻し',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:1,取置元種別:'EMS',元EMS番号:'EG1'};
  const rows = [
    Object.assign({取置ID:'A',戻し処理結果:'未確認'},base),
    Object.assign({取置ID:'B',戻し処理結果:'現物あり'},base),
    Object.assign({取置ID:'C',戻し処理結果:'在庫なし'},base),
    Object.assign({取置ID:'D',戻し処理結果:'再引当済み'},base),
    Object.assign({取置ID:'E',戻し処理結果:'Yahoo反映済み'},base)
  ];
  const summary = context.取り置き_集計_(rows, [{処理ID:'YAHOO|RETURN|E',EMS番号:'EG1',商品コード:'AAA',数量:1}]);
  const use = summary.usageBySupply['EG1|AAA'];
  assert.strictEqual(JSON.stringify(use), JSON.stringify({取り置き中:0,発送済み:0,戻し未処理:2,在庫なし確定:1,Yahoo移動済み:1}));
  assert.strictEqual(summary.confirmedReturns.length, 1);
  assert.strictEqual(summary.confirmedReturns[0].取置ID, 'B');
});

test('重複ID・非整数・注文数量超過をエラーにする', () => {
  const rows = [
    {取置ID:'X',状態:'取り置き中',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:1.5},
    {取置ID:'X',状態:'取り置き中',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:2}
  ];
  const summary = context.取り置き_集計_(rows, []);
  assert.ok(summary.errors.some(e => /重複/.test(e)));
  assert.ok(summary.errors.some(e => /正の整数/.test(e)));
});

test('現物ありキャンセル戻しを今回EMSより先に最古注文へ再引当する', () => {
  const result = context.取り置き_割当計算_({
    orders:[
      {ban:'100',code:'AAA',sku:'AAAb',qty:1,sortKey:100,i:0,keys:['AAA']},
      {ban:'200',code:'AAA',sku:'AAAb',qty:1,sortKey:200,i:1,keys:['AAA']}
    ],
    ledger:[{取置ID:'OLD',状態:'キャンセル戻し',受注番号:'050',商品コード:'AAA',SKU:'AAAb',取り置き数量:1,取置元種別:'EMS',元EMS番号:'EG0',戻し処理結果:'現物あり'}],
    movements:[],
    supplies:[{ems:'EG1',code:'AAA',qty:1,arrival:'2026-07-12'}],
    explicit:[]
  });
  assert.strictEqual(result.newRows[0].受注番号, '100');
  assert.strictEqual(result.newRows[0].取置元種別, 'キャンセル再引当');
  assert.strictEqual(result.newRows[0].元取置ID, 'OLD');
  assert.strictEqual(result.newRows[1].受注番号, '200');
  assert.strictEqual(result.newRows[1].元EMS番号, 'EG1');
});

test('YMNGD09は3個供給・1個確保なら余り2個', () => {
  const result = context.取り置き_割当計算_({
    orders:[{ban:'101',code:'YMNGD09',sku:'YMNGD09b',qty:1,sortKey:101,i:0,keys:['YMNGD09']}],
    ledger:[], movements:[],
    supplies:[{ems:'EG049827401KR',code:'YMNGD09',qty:3,arrival:'2026-07-12'}], explicit:[]
  });
  assert.strictEqual(result.newRows.reduce((s,r)=>s+r.取り置き数量,0),1);
  assert.strictEqual(result.surplus[0].qty,2);
});

test('JMEE167の10個を既存7個と新規3個に分けても超過しない', () => {
  const result = context.取り置き_割当計算_({
    orders:[
      {ban:'10117284',code:'JMEE167',sku:'JMEE167b',qty:7,sortKey:10117284,i:0,keys:['JMEE167']},
      {ban:'10117602',code:'JMEE167',sku:'JMEE167b',qty:3,sortKey:10117602,i:1,keys:['JMEE167']}
    ],
    ledger:[{取置ID:'EMS|EG1|10117284',状態:'取り置き中',受注番号:'10117284',商品コード:'JMEE167',SKU:'JMEE167b',取り置き数量:7,取置元種別:'EMS',元EMS番号:'EG1'}],
    movements:[], supplies:[{ems:'EG1',code:'JMEE167',qty:10,arrival:'2026-07-12'}], explicit:[]
  });
  assert.strictEqual(result.newRows.length,1);
  assert.strictEqual(result.newRows[0].受注番号,'10117602');
  assert.strictEqual(result.newRows[0].取り置き数量,3);
  assert.strictEqual(JSON.stringify(result.errors),'[]');
});

test('注文番号在庫10117375の2個を指定注文だけへ割り当てる', () => {
  const result = context.取り置き_割当計算_({
    orders:[{ban:'10117375',code:'KAGURA08W-PG',sku:'KAGURA08W-PGb',qty:2,sortKey:10117375,i:0,keys:['KAGURA08W-PG']}],
    ledger:[], movements:[],
    supplies:[{ems:'EG1',code:'10117375',qty:2,arrival:'2026-07-12',directBan:'10117375'}], explicit:[]
  });
  assert.strictEqual(result.newRows.length,1);
  assert.strictEqual(result.newRows[0].受注番号,'10117375');
  assert.strictEqual(result.newRows[0].取り置き数量,2);
  assert.strictEqual(result.newRows[0].商品コード,'KAGURA08W-PG');
  assert.strictEqual(result.newRows[0].元EMS商品コード,'10117375');
});

test('同じ入力で再計算しても既存EMS確保を追加しない', () => {
  const input={orders:[{ban:'101',code:'AAA',sku:'AAAb',qty:1,sortKey:101,i:0,keys:['AAA']}],movements:[],supplies:[{ems:'EG1',code:'AAA',qty:1,arrival:'2026-07-12'}],explicit:[]};
  const first=context.取り置き_割当計算_(Object.assign({ledger:[]},input));
  const second=context.取り置き_割当計算_(Object.assign({ledger:first.newRows},input));
  assert.strictEqual(first.newRows.length,1);
  assert.strictEqual(second.newRows.length,0);
  assert.strictEqual(second.surplus.length,0);
});

process.exitCode = failures ? 1 : 0;
console.log(failures ? `FAILURES: ${failures}` : 'ALL TESTS PASSED');
