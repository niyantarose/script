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
  'Project_24/取り置き計算.js',
  'Project_24/取り置き台帳.js'
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
    ledger:[
      {取置ID:'OLD',状態:'キャンセル戻し',受注番号:'050',商品コード:'AAA',SKU:'AAAb',取り置き数量:1,取置元種別:'EMS',元EMS番号:'EG0',戻し処理結果:'現物あり'},
      {取置ID:'MISS',状態:'キャンセル戻し',受注番号:'060',商品コード:'AAA',SKU:'AAAb',取り置き数量:1,取置元種別:'EMS',元EMS番号:'EGX',戻し処理結果:'在庫なし'}
    ],
    movements:[],
    supplies:[{ems:'EG1',code:'AAA',qty:1,arrival:'2026-07-12'}],
    explicit:[]
  });
  assert.strictEqual(result.newRows[0].受注番号, '100');
  assert.strictEqual(result.newRows[0].取置元種別, 'キャンセル再引当');
  assert.strictEqual(result.newRows[0].元取置ID, 'OLD');
  assert.strictEqual(result.newRows[1].受注番号, '200');
  assert.strictEqual(result.newRows[1].元EMS番号, 'EG1');
  assert.strictEqual(result.newRows.some(r=>r.元取置ID==='MISS'),false);
  assert.strictEqual(result.surplus.length,0);
});

test('開始前取り置き5商品を今回EMSへ再引当しない', () => {
  const codes=['KRBLCM16-02','OFW300','OFW301','MZBGD03','JNXGD01'];
  const orders=codes.map((code,index)=>({ban:String(10100+index),code,sku:code+'b',qty:1,sortKey:10100+index,i:index,keys:[code]}));
  const ledger=orders.map(order=>({取置ID:'INIT|'+context.取り置き_行キー_(order),状態:'取り置き中',受注番号:order.ban,
    商品コード:order.code,SKU:order.sku,取り置き数量:1,取置元種別:'開始前在庫'}));
  const supplies=codes.map(code=>({ems:'EGNEW',code,qty:1,arrival:'2026-07-12'}));
  const result=context.取り置き_割当計算_({orders,ledger,movements:[],supplies,explicit:[]});
  assert.strictEqual(result.newRows.length,0);
  assert.strictEqual(result.surplus.length,5);
  result.surplus.forEach(row=>assert.strictEqual(row.qty,1));
});

test('POEM65とRECIPE42の受注番号タグは指定注文へだけ入る', () => {
  const orders=[
    {ban:'10116569',code:'POEM65',sku:'POEM65b',qty:1,sortKey:1,i:0,keys:['POEM65']},
    {ban:'10117126',code:'RECIPE42',sku:'RECIPE42b',qty:1,sortKey:2,i:1,keys:['RECIPE42']},
    {ban:'10199999',code:'POEM65',sku:'POEM65b',qty:1,sortKey:3,i:2,keys:['POEM65']}
  ];
  const supplies=[
    {ems:'EG1',code:'POEM65',qty:1,directBan:'10116569'},
    {ems:'EG1',code:'RECIPE42',qty:1,directBan:'10117126'}
  ];
  const result=context.取り置き_割当計算_({orders,ledger:[],movements:[],supplies,explicit:[]});
  assert.strictEqual(JSON.stringify(result.newRows.map(r=>r.受注番号)),JSON.stringify(['10116569','10117126']));
  assert.strictEqual(result.orders.find(o=>o.ban==='10199999').need,1);
});

test('固定activeと同じP列表示は除外し、新規割当だけを計画へ渡す', () => {
  const ledger=[{取置ID:'EMS|EG1|AAA|101',状態:'取り置き中',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:1,
    取置元種別:'EMS',元EMS番号:'EG1',元EMS商品コード:'AAA'}];
  const pPlan={rows:[{ems:'EG1',code:'AAA',entries:[{ban:'101',qty:1},{ban:'102',qty:1}]}]};
  assert.strictEqual(JSON.stringify(context.P列計画_新規確定割当_(pPlan,ledger)),JSON.stringify([
    {ems:'EG1',code:'AAA',ban:'102',qty:1}
  ]));
  const result=context.取り置き_割当計算_({
    orders:[
      {ban:'101',code:'AAA',sku:'AAAb',qty:1,sortKey:1,i:0,keys:['AAA']},
      {ban:'102',code:'AAA',sku:'AAAb',qty:1,sortKey:2,i:1,keys:['AAA']}
    ],ledger,movements:[],supplies:[{ems:'EG1',code:'AAA',qty:2,arrival:'2026-07-12'}],
    explicit:context.P列計画_新規確定割当_(pPlan,ledger)
  });
  assert.strictEqual(JSON.stringify(result.errors),'[]');
  assert.strictEqual(result.newRows.length,1);
  assert.strictEqual(result.newRows[0].受注番号,'102');
});

test('割当計画反映は戻し全フィールドと決定ID upsertを1回保存する', () => {
  const registered=new Date('2026-07-01T00:00:00Z'), now=new Date('2026-07-15T00:00:00Z');
  const existing=[
    {取置ID:'RETURN',状態:'キャンセル戻し',受注番号:'100',商品コード:'AAA',SKU:'AAAb',取り置き数量:3,戻し処理結果:'現物あり',
      '終了理由・メモ':'old',登録日時:registered,更新日時:registered},
    {取置ID:'SAME',状態:'取り置き中',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:1,登録日時:registered,更新日時:registered}
  ];
  const plan={returnUpdates:[{取置ID:'RETURN',戻し処理結果:'現物あり',取り置き数量:1,'終了理由・メモ':'2個を再引当済み'}],newRows:[
    {取置ID:'SAME',状態:'取り置き中',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:2,取置元種別:'EMS',元EMS番号:'EG1'},
    {取置ID:'NEW',状態:'取り置き中',受注番号:'102',商品コード:'BBB',SKU:'BBBb',取り置き数量:1,取置元種別:'EMS',元EMS番号:'EG1'}
  ]};
  let saved=null,writes=0;
  context.取り置き台帳_保存_=rows=>{saved=rows;writes++;};
  context.取り置き台帳_割当計画を反映_(plan,existing,now);
  assert.strictEqual(writes,1);
  assert.strictEqual(saved.length,3);
  const returned=saved.find(r=>r.取置ID==='RETURN');
  assert.strictEqual(returned.取り置き数量,1);
  assert.strictEqual(returned['終了理由・メモ'],'2個を再引当済み');
  assert.strictEqual(returned.登録日時,registered);
  assert.strictEqual(returned.更新日時,now);
  assert.strictEqual(saved.find(r=>r.取置ID==='SAME').登録日時,registered);
  assert.strictEqual(saved.find(r=>r.取置ID==='NEW').登録日時,now);
});

test('今回EMS消費者は台帳の元EMSと一致するactive行、日本在庫はplan surplusだけ', () => {
  const supplies=[
    {ems:'EG1',code:'AAA',qty:3,arrival:'2026-07-12'},
    {ems:'EG2',code:'AAA',qty:1,arrival:'2026-07-13'}
  ];
  const projected=[
    {取置ID:'OLD',状態:'取り置き中',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:1,取置元種別:'EMS',元EMS番号:'EG1',元EMS商品コード:'AAA'},
    {取置ID:'NEW',状態:'取り置き中',受注番号:'102',商品コード:'AAA',SKU:'AAAb',取り置き数量:1,取置元種別:'EMS',元EMS番号:'EG1',元EMS商品コード:'AAA'},
    {取置ID:'OTHER',状態:'取り置き中',受注番号:'103',商品コード:'AAA',SKU:'AAAb',取り置き数量:1,取置元種別:'EMS',元EMS番号:'EG2',元EMS商品コード:'AAA'}
  ];
  const allocationPlan={newRows:[projected[1]],surplus:[{ems:'EG1',code:'AAA',qty:1,arrival:'2026-07-12'}]};
  const out=context.引当出力計画_(supplies,allocationPlan,projected);
  const eg1=out.supplies.find(r=>r.ems==='EG1');
  assert.strictEqual(JSON.stringify(eg1.consumers.map(c=>[c.ban,c.qty,c.current])),JSON.stringify([
    ['101',1,false],['102',1,true]
  ]));
  assert.strictEqual(eg1.surplus,1);
  assert.strictEqual(JSON.stringify(out.surplus),JSON.stringify(allocationPlan.surplus));
});

test('行状態は既存activeと今回allocを二重控除せず、計画後数量は表示用に分ける', () => {
  const line={ban:'101',code:'AAA',sku:'AAAb',qty:2,kbn:'取り寄せ',キャンセル:false};
  const key=context.取り置き_行キー_(line);
  const newRow={取置ID:'NEW',状態:'取り置き中',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:1,元EMS番号:'EG1',元EMS商品コード:'AAA'};
  context.引当計画_行へ反映_([line],{activeByKey:{[key]:0},activeRowsByKey:{}},{activeByKey:{[key]:1},activeRowsByKey:{[key]:[newRow]}},[newRow]);
  assert.strictEqual(line.取り置き中数量,0);
  assert.strictEqual(line.計画後取り置き中数量,1);
  assert.strictEqual(line.alloc,1);
  assert.strictEqual(context.残必要計算_(line),1);
  assert.strictEqual(line.引当成立,false);
});

test('引当切替差分は受注番号・商品コード・SKUをキーに追加更新削除を出す', () => {
  const planned=[
    {ban:'101',code:'AAA',sku:'AAAb',qty:1,state:'取り置き中',ems:'EG1'},
    {ban:'102',code:'BBB',sku:'BBBb',qty:1,state:'引当(今回)',ems:'EG2'}
  ];
  const current=[
    {ban:'101',code:'AAA',sku:'AAAb',qty:1,state:'在庫待ち',ems:''},
    {ban:'103',code:'CCC',sku:'CCCb',qty:1,state:'取り置き中',ems:'EG3'}
  ];
  const result=context.引当切替差分_純計算_(planned,current);
  assert.strictEqual(JSON.stringify(result.map(r=>r.change)),JSON.stringify(['更新','追加','削除']));
  assert.strictEqual(result[0].key,'101|AAA|AAAB');
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

test('EMS FIFOは入力順でなく到着日の古い供給を先に使う', () => {
  const result=context.取り置き_割当計算_({
    orders:[{ban:'101',code:'AAA',sku:'AAAb',qty:1,sortKey:101,i:0,keys:['AAA']}],
    ledger:[],movements:[],explicit:[],
    supplies:[
      {ems:'EG2',code:'AAA',qty:1,arrival:'2026-07-13'},
      {ems:'EG9',code:'AAA',qty:1,arrival:'2026-07-12'}
    ]
  });
  assert.strictEqual(result.newRows[0].元EMS番号,'EG9');
  assert.strictEqual(result.surplus[0].ems,'EG2');
});

test('EMS FIFOは同じ到着日ならEMS番号で決定する', () => {
  const result=context.取り置き_割当計算_({
    orders:[{ban:'101',code:'AAA',sku:'AAAb',qty:1,sortKey:101,i:0,keys:['AAA']}],
    ledger:[],movements:[],explicit:[],
    supplies:[
      {ems:'EG2',code:'AAA',qty:1,arrival:'2026-07-12'},
      {ems:'EG1',code:'AAA',qty:1,arrival:'2026-07-12'}
    ]
  });
  assert.strictEqual(result.newRows[0].元EMS番号,'EG1');
});

test('P列1個確定後に同じEMSから残り2個を割り当てても同じIDを1行へ集約する', () => {
  const result=context.取り置き_割当計算_({
    orders:[{ban:'101',code:'AAA',sku:'AAAb',qty:3,sortKey:101,i:0,keys:['AAA']}],
    ledger:[],movements:[],
    supplies:[{ems:'EG1',code:'AAA',qty:3,arrival:'2026-07-12'}],
    explicit:[{ems:'EG1',code:'AAA',ban:'101',qty:1}]
  });
  assert.strictEqual(result.newRows.length,1);
  assert.strictEqual(result.newRows[0].取り置き数量,3);
  assert.strictEqual(new Set(result.newRows.map(r=>r.取置ID)).size,1);
  assert.strictEqual(JSON.stringify(result.errors),'[]');
});

test('P列確定数量を注文必要数まで割り当てられなければエラーにする', () => {
  const result=context.取り置き_割当計算_({
    orders:[{ban:'101',code:'AAA',sku:'AAAb',qty:1,sortKey:101,i:0,keys:['AAA']}],
    ledger:[],movements:[],
    supplies:[{ems:'EG1',code:'AAA',qty:2,arrival:'2026-07-12'}],
    explicit:[{ems:'EG1',code:'AAA',ban:'101',qty:2}]
  });
  assert.ok(result.errors.some(e=>/P列確定数量を満たせない/.test(e)));
});

test('既存台帳エラーは割当計算結果で重複しない', () => {
  const result=context.取り置き_割当計算_({
    orders:[],movements:[],supplies:[],explicit:[],
    ledger:[{取置ID:'',状態:'取り置き中',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:0}]
  });
  assert.strictEqual(result.errors.length,new Set(result.errors).size);
});

process.exitCode = failures ? 1 : 0;
console.log(failures ? `FAILURES: ${failures}` : 'ALL TESTS PASSED');
