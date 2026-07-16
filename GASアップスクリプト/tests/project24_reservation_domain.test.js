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
  'Project_24/取り置き台帳.js',
  'Project_24/消込台帳.js' // 取り置き出荷_(メモ判定)を未台帳出荷計画が使う
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

test('同一EMS・同一基底コードでも異なるdirect受注の供給を混ぜない', () => {
  const orders=[
    {ban:'10116569',code:'POEM65',sku:'POEM65b',qty:1,sortKey:1,i:0,keys:['POEM65']},
    {ban:'10199999',code:'POEM65',sku:'POEM65b',qty:1,sortKey:2,i:1,keys:['POEM65']}
  ];
  const supplies=[
    {ems:'EG1',code:'POEM65',sourceCode:'POEM65（10116569）',qty:1,directBan:'10116569'},
    {ems:'EG1',code:'POEM65',sourceCode:'POEM65（10199999）',qty:1,directBan:'10199999'}
  ];
  const result=context.取り置き_割当計算_({orders,ledger:[],movements:[],supplies,explicit:[]});
  assert.strictEqual(JSON.stringify(result.errors),'[]');
  assert.strictEqual(JSON.stringify(result.newRows.map(r=>r.受注番号)),JSON.stringify(['10116569','10199999']));
  assert.strictEqual(JSON.stringify(result.orders.map(r=>r.need)),JSON.stringify([0,0]));
});

test('direct供給の元EMS商品コードはタグ付き原文を台帳へ保持する', () => {
  const sourceCode='PoEm65（10116569）';
  const result=context.取り置き_割当計算_({
    orders:[{ban:'10116569',code:'POEM65',sku:'POEM65b',qty:1,sortKey:1,i:0,keys:['POEM65']}],
    ledger:[],movements:[],
    supplies:[{ems:'EG1',code:'POEM65',sourceCode,qty:1,directBan:'10116569'}],explicit:[]
  });
  assert.strictEqual(JSON.stringify(result.errors),'[]');
  assert.strictEqual(result.newRows[0].元EMS商品コード,sourceCode);
});

test('同一source identityでもdirect ownerが異なる供給は別々に割り当てる', () => {
  const result=context.取り置き_割当計算_({
    orders:[
      {ban:'101',code:'POEM65',sku:'POEM65b',qty:1,sortKey:1,i:0,keys:['POEM65']},
      {ban:'102',code:'POEM65',sku:'POEM65b',qty:1,sortKey:2,i:1,keys:['POEM65']}
    ],ledger:[],movements:[],supplies:[
      {ems:'EG1',code:'POEM65',sourceCode:'POEM65',qty:1,directBan:'101'},
      {ems:'EG1',code:'POEM65',sourceCode:'POEM65',qty:1,directBan:'102'}
    ],explicit:[]
  });
  assert.strictEqual(JSON.stringify(result.errors),'[]');
  assert.strictEqual(JSON.stringify(result.newRows.map(r=>r.受注番号)),JSON.stringify(['101','102']));
});

test('direct供給の既存台帳使用と移動使用は一致するsource identityから一度だけ差し引く', () => {
  const orders=[
    {ban:'101',code:'POEM65',sku:'POEM65b',qty:1,sortKey:1,i:0,keys:['POEM65']},
    {ban:'102',code:'POEM65',sku:'POEM65b',qty:1,sortKey:2,i:1,keys:['POEM65']}
  ];
  const ledger=[{取置ID:'A',状態:'取り置き中',受注番号:'101',商品コード:'POEM65',SKU:'POEM65b',取り置き数量:1,
    取置元種別:'EMS',元EMS番号:'EG1',元EMS商品コード:'POEM65（101）'}];
  const movements=[{処理ID:'M1',EMS番号:'EG1',商品コード:'POEM65（102）',数量:1}];
  const supplies=[
    {ems:'EG1',code:'POEM65',sourceCode:'POEM65（101）',qty:1,directBan:'101'},
    {ems:'EG1',code:'POEM65',sourceCode:'POEM65（102）',qty:1,directBan:'102'}
  ];
  const result=context.取り置き_割当計算_({orders,ledger,movements,supplies,explicit:[]});
  assert.strictEqual(JSON.stringify(result.errors),'[]');
  assert.strictEqual(result.newRows.length,0);
  assert.strictEqual(JSON.stringify(result.orders.map(r=>r.need)),JSON.stringify([0,1]));
  assert.strictEqual(result.surplus.length,0);
});

test('同一受注番号に複数商品があってもdirectタグは基底コード一致行だけへ入る', () => {
  const result=context.取り置き_割当計算_({
    orders:[
      {ban:'10116569',code:'RECIPE42',sku:'RECIPE42b',qty:1,sortKey:1,i:0,keys:['RECIPE42']},
      {ban:'10116569',code:'POEM65',sku:'POEM65b',qty:1,sortKey:1,i:1,keys:['POEM65']}
    ],ledger:[],movements:[],supplies:[
      {ems:'EG1',code:'POEM65',sourceCode:'POEM65（10116569）',qty:1,directBan:'10116569'}
    ],explicit:[]
  });
  assert.strictEqual(JSON.stringify(result.errors),'[]');
  assert.strictEqual(result.newRows[0].商品コード,'POEM65');
  assert.strictEqual(JSON.stringify(result.orders.map(r=>r.need)),JSON.stringify([1,0]));
});

test('新引当照合は連結コードを推測分割せず販売SKU末尾a/bだけを正規化する', () => {
  const concatenated='OFW305-1OFW304-2';
  const keys=context.引当用照合キー一覧_(concatenated+'b',concatenated);
  assert.ok(keys.includes(concatenated+'B'));
  assert.ok(keys.includes(concatenated));
  assert.strictEqual(keys.includes('OFW305-1'),false);
  assert.strictEqual(keys.includes('OFW304-2'),false);
  assert.strictEqual(context.引当用照合キー一覧_('OFW3-0405-05b','OFW3-0405-05').includes('OFW3-05'),false);
});

test('固定activeと同じP列表示は除外し、新規割当だけを計画へ渡す', () => {
  const ledger=[{取置ID:'EMS|EG1|AAA|101',状態:'取り置き中',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:1,
    取置元種別:'EMS',元EMS番号:'EG1',元EMS商品コード:'AAA'}];
  const pPlan={rows:[{ems:'EG1',code:'AAA',entries:[{ban:'101',qty:1},{ban:'102',qty:1}]}]};
  assert.strictEqual(JSON.stringify(context.P列計画_新規確定割当_(pPlan,ledger)),JSON.stringify([
    {ems:'EG1',code:'AAA',sourceCode:'AAA',ban:'102',qty:1}
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

test('出力計画も同一source identityのdirect ownerを別供給として保つ', () => {
  const supplies=[
    {ems:'EG1',code:'POEM65',sourceCode:'POEM65',directBan:'101',qty:1,arrival:'2026-07-12'},
    {ems:'EG1',code:'POEM65',sourceCode:'POEM65',directBan:'102',qty:1,arrival:'2026-07-12'}
  ];
  const projected=[
    {取置ID:'A',状態:'取り置き中',受注番号:'101',商品コード:'POEM65',SKU:'POEM65b',取り置き数量:1,元EMS番号:'EG1',元EMS商品コード:'POEM65'},
    {取置ID:'B',状態:'取り置き中',受注番号:'102',商品コード:'POEM65',SKU:'POEM65b',取り置き数量:1,元EMS番号:'EG1',元EMS商品コード:'POEM65'}
  ];
  const output=context.引当出力計画_(supplies,{newRows:projected,surplus:[]},projected);
  assert.strictEqual(output.supplies.length,2);
  assert.strictEqual(JSON.stringify(output.supplies.map(s=>[s.directBan,s.consumers.map(c=>c.ban)])),JSON.stringify([
    ['101',['101']],['102',['102']]
  ]));
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

test('タグ付きdirect供給は実際に一致した受注枝番を表示し元コードを変更しない', () => {
  [
    ['10116569','POEM65-1','POEM65-1b','POEM65（10116569）'],
    ['10117126','RECIPE42-2','RECIPE42-2b','RECIPE42/10117126']
  ].forEach(([ban,code,sku,sourceCode])=>{
    const line={ban,code,sku,qty:1,kbn:'取り寄せ',キャンセル:false};
    const originalCode=line.code;
    const key=context.取り置き_行キー_(line);
    const newRow={取置ID:'NEW',状態:'取り置き中',受注番号:ban,商品コード:code,SKU:sku,取り置き数量:1,元EMS番号:'EG1',元EMS商品コード:sourceCode};
    context.引当計画_行へ反映_([line],{activeByKey:{},activeRowsByKey:{}},{activeByKey:{[key]:1},activeRowsByKey:{[key]:[newRow]}},[newRow]);
    assert.strictEqual(line.matchedKey,code);
    assert.strictEqual(context.注文一覧表示コード_(line,false,null),code);
    assert.strictEqual(line.code,originalCode);
  });
});

test('通常供給の表示はEMS照合コードを維持する', () => {
  const line={ban:'10116569',code:'POEM65-1',sku:'POEM65-1b',qty:1,kbn:'取り寄せ',キャンセル:false};
  const key=context.取り置き_行キー_(line);
  const newRow={取置ID:'NEW',状態:'取り置き中',受注番号:line.ban,商品コード:line.code,SKU:line.sku,取り置き数量:1,元EMS番号:'EG1',元EMS商品コード:'POEM65'};
  context.引当計画_行へ反映_([line],{activeByKey:{},activeRowsByKey:{}},{activeByKey:{[key]:1},activeRowsByKey:{[key]:[newRow]}},[newRow]);
  assert.strictEqual(line.matchedKey,'POEM65');
  assert.strictEqual(context.注文一覧表示コード_(line,false,null),'POEM65');
  assert.strictEqual(line.code,'POEM65-1');
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

test('現行出力にSKU列がなくても同一受注・同一コードの複数SKUを行順で復元する', () => {
  const planned=[
    {ban:'101',code:'AAA',sku:'AAA-RED',qty:1,state:'在庫待ち',ems:''},
    {ban:'101',code:'AAA',sku:'AAA-BLUE',qty:1,state:'在庫待ち',ems:''}
  ];
  const values=[
    ['受注番号','商品コード','個数','状態','EMS番号'],
    ['101','AAA',1,'在庫待ち',''],
    ['101','AAA',1,'在庫待ち','']
  ];
  const wait={getDataRange:()=>({getValues:()=>values})};
  const ss={getSheetByName:name=>name==='引当待ち'?wait:null};
  const current=context.引当切替_現行出力行_(ss,planned);
  assert.strictEqual(JSON.stringify(current.map(r=>r.sku)),JSON.stringify(['AAA-RED','AAA-BLUE']));
  assert.strictEqual(JSON.stringify(context.引当切替差分_純計算_(planned,current)),'[]');
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

// ===== タグ異常は警告して完走する(④全体を止めない・持ち主のない分は余りへ) =====
// キャンセル済み注文の名指しタグや多め買付は日常的に発生するため、全体中止にしない(ユーザー決定)。

test('名指し注文が不在のdirect供給は警告して余りへ回し、他の注文は完走する', () => {
  const orders=[{ban:'10200000',code:'AAA',sku:'AAAb',qty:1,sortKey:1,i:0,keys:['AAA']}];
  const supplies=[
    {ems:'EG1',code:'POEM65',sourceCode:'POEM65（10116569）',qty:2,directBan:'10116569'},
    {ems:'EG1',code:'AAA',sourceCode:'AAA',qty:1}
  ];
  const result=context.取り置き_割当計算_({orders,ledger:[],movements:[],supplies,explicit:[]});
  assert.strictEqual(result.errors.length,0,'エラーで中止しない: '+JSON.stringify(result.errors));
  assert.strictEqual(result.warnings.length,1);
  assert.ok(result.warnings[0].indexOf('10116569')>=0);
  assert.strictEqual(result.surplus.find(s=>s.directBan==='10116569').qty,2,'不在名指しの箱は全量余りへ');
  assert.strictEqual(result.newRows.filter(r=>r.受注番号==='10200000').length,1,'他の注文の引当は完走する');
});

test('多め買付(供給>注文数)のdirect供給は注文分だけ引き当て、残りを警告付きで余りへ', () => {
  const orders=[{ban:'10116569',code:'POEM65',sku:'POEM65b',qty:1,sortKey:1,i:0,keys:['POEM65']}];
  const supplies=[{ems:'EG1',code:'POEM65',sourceCode:'POEM65（10116569）',qty:2,directBan:'10116569'}];
  const result=context.取り置き_割当計算_({orders,ledger:[],movements:[],supplies,explicit:[]});
  assert.strictEqual(result.errors.length,0,'エラーで中止しない: '+JSON.stringify(result.errors));
  assert.strictEqual(result.warnings.length,1);
  assert.strictEqual(result.newRows.length,1);
  assert.strictEqual(result.newRows[0].取り置き数量,1);
  assert.strictEqual(result.surplus.find(s=>s.directBan==='10116569').qty,1,'多め分は余りへ');
});

test('正常なdirect割当では警告を出さない', () => {
  const orders=[{ban:'10116569',code:'POEM65',sku:'POEM65b',qty:2,sortKey:1,i:0,keys:['POEM65']}];
  const supplies=[{ems:'EG1',code:'POEM65',sourceCode:'POEM65（10116569）',qty:2,directBan:'10116569'}];
  const result=context.取り置き_割当計算_({orders,ledger:[],movements:[],supplies,explicit:[]});
  assert.strictEqual(result.errors.length,0);
  assert.strictEqual((result.warnings||[]).length,0);
});

// ===== 台帳に載らない出荷(週末GoQ等)の自動発送済み登録 =====
// 旧9a8f8d8「処理済CSVで事実の在庫差引き(売り越し防止)」の台帳版(ユーザー決定)。

test('台帳外出荷は発送済み行として自動登録し箱の残りを差し引く', () => {
  const supplies=[{ems:'EG1',code:'AAA',sourceCode:'AAA',qty:3,arrival:'2026-07-12'}];
  const r=context.取り置き_未台帳出荷計画_(
    [{ban:'10117000',code:'AAA',sku:'AAAb',qty:1,基準日:'2026-07-12',メモ:''}],[],[],supplies);
  assert.strictEqual(r.newRows.length,1);
  assert.strictEqual(r.newRows[0].状態,'発送済み');
  assert.strictEqual(r.newRows[0].元EMS番号,'EG1');
  assert.strictEqual(r.newRows[0].取り置き数量,1);
  assert.strictEqual(r.newRows[0].取置元種別,'出荷実績');
  assert.strictEqual(r.review.length,0);
});

test('台帳に行がある注文・メモ取り置き・発送日より後に着いた箱は自動登録しない', () => {
  const late=[{ems:'EG1',code:'AAA',sourceCode:'AAA',qty:3,arrival:'2026-07-14'}];
  const r1=context.取り置き_未台帳出荷計画_(
    [{ban:'1',code:'AAA',sku:'AAAb',qty:1,基準日:'2026-07-12',メモ:''}],[],[],late);
  assert.strictEqual(r1.newRows.length,0,'発送より後に着いた箱は出どころではない');
  const r2=context.取り置き_未台帳出荷計画_(
    [{ban:'1',code:'AAA',sku:'AAAb',qty:1,基準日:'2026-07-15',メモ:'取り置き'}],[],[],late);
  assert.strictEqual(r2.newRows.length,0,'メモ取り置きは人為オーバーライド');
  const ledger=[{取置ID:'X',状態:'発送済み',受注番号:'1',商品コード:'AAA',SKU:'AAAb',取り置き数量:1,取置元種別:'EMS',元EMS番号:'EG1',元EMS商品コード:'AAA'}];
  const r3=context.取り置き_未台帳出荷計画_(
    [{ban:'1',code:'AAA',sku:'AAAb',qty:1,基準日:'2026-07-15',メモ:''}],ledger,[],late);
  assert.strictEqual(r3.newRows.length,0,'台帳に行がある注文は遷移側で扱う');
});

test('既存使用を差し引いた残りまでしか自動登録せず、不足分は要確認へ', () => {
  const supplies=[{ems:'EG1',code:'AAA',sourceCode:'AAA',qty:2,arrival:'2026-07-12'}];
  const ledger=[{取置ID:'H',状態:'取り置き中',受注番号:'2',商品コード:'AAA',SKU:'AAAz',取り置き数量:1,取置元種別:'EMS',元EMS番号:'EG1',元EMS商品コード:'AAA'}];
  const r=context.取り置き_未台帳出荷計画_(
    [{ban:'9',code:'AAA',sku:'AAAb',qty:2,基準日:'2026-07-12',メモ:''}],ledger,[],supplies);
  assert.strictEqual(r.newRows.length,1);
  assert.strictEqual(r.newRows[0].取り置き数量,1);
  assert.strictEqual(r.review.length,1);
  assert.ok(r.review[0].理由.indexOf('1/2')>=0);
});

test('同じ出荷を二回渡しても決定IDで重複登録しない(再実行冪等)', () => {
  const supplies=[{ems:'EG1',code:'AAA',sourceCode:'AAA',qty:3,arrival:'2026-07-12'}];
  const shipped=[{ban:'10117000',code:'AAA',sku:'AAAb',qty:1,基準日:'2026-07-12',メモ:''}];
  const first=context.取り置き_未台帳出荷計画_(shipped,[],[],supplies);
  const second=context.取り置き_未台帳出荷計画_(shipped,first.newRows,[],supplies);
  assert.strictEqual(second.newRows.length,0,'登録済み(台帳へ保存済み)なら次回は作らない');
  assert.strictEqual(first.newRows[0].取置ID, context.取り置き_未台帳出荷計画_(shipped,[],[],supplies).newRows[0].取置ID,'IDは決定的');
});

// ===== 個別引当・引当キャンセルボタンの台帳直書きヘルパー =====
// P列は派生表示になったため、ボタンは台帳へ直接 取り置き中/手動解除 を書く(ユーザー決定)。

test('個別_台帳候補_: 到着済み箱から台帳使用を差し引いた残あり候補を返す', () => {
  const emsRows=[
    {状態:'到着済',EMS番号:'EG1',商品コード:'AAA',数量:2,EMS到着日:'2026-07-12'},
    {状態:'到着済',EMS番号:'EG2',商品コード:'AAA',数量:1,EMS到着日:'2026-07-14'},
    {状態:'在庫反映済み',EMS番号:'EG0',商品コード:'AAA',数量:5},
    {状態:'到着済',EMS番号:'棚卸',商品コード:'AAA',数量:9}
  ];
  const ledger=[{取置ID:'X',状態:'取り置き中',受注番号:'1',商品コード:'AAA',SKU:'AAAb',取り置き数量:2,取置元種別:'EMS',元EMS番号:'EG1',元EMS商品コード:'AAA'}];
  const summary=context.取り置き_集計_(ledger,[]);
  const out=context.個別_台帳候補_(emsRows,['AAA'],'',summary);
  assert.strictEqual(out.length,1,'EG1は使用済みで残0、反映済み・棚卸は対象外');
  assert.strictEqual(out[0].item.EMS番号,'EG2');
  assert.strictEqual(out[0].残,1);
});

test('個別_台帳候補_: 受注番号タグの箱はコード不一致でもこの受注の候補になる', () => {
  const emsRows=[{状態:'到着済',EMS番号:'EG1',商品コード:'REQBOOK01（10117000）',数量:1,EMS到着日:'2026-07-12'}];
  const summary=context.取り置き_集計_([],[]);
  assert.strictEqual(context.個別_台帳候補_(emsRows,['ZZZ'],'10117000',summary).length,1,'タグ一致で救済');
  assert.strictEqual(context.個別_台帳候補_(emsRows,['ZZZ'],'10119999',summary).length,0,'他人のタグは候補外');
});

test('個別_台帳引当行_: 新規は行を追加し、同じ決定IDなら数量を加算する', () => {
  const order={ban:'10117000',code:'AAA',sku:'AAAb'};
  const first=context.個別_台帳引当行_([],order,1,'EG1','AAA','2026-07-16 10:00');
  assert.strictEqual(first.error,'');
  assert.strictEqual(first.rows.length,1);
  assert.strictEqual(first.rows[0].状態,'取り置き中');
  const second=context.個別_台帳引当行_(first.rows,order,1,'EG1','AAA','2026-07-16 10:05');
  assert.strictEqual(second.rows.length,1,'同じ注文×箱は行を増やさない');
  assert.strictEqual(second.rows[0].取り置き数量,2);
});

test('個別_台帳解除計画_: 行キーの取り置き中だけを手動解除にする', () => {
  const ledger=[
    {取置ID:'A',状態:'取り置き中',受注番号:'1',商品コード:'AAA',SKU:'AAAb',取り置き数量:1,元EMS番号:'EG1',元EMS商品コード:'AAA'},
    {取置ID:'B',状態:'発送済み',受注番号:'1',商品コード:'AAA',SKU:'AAAb',取り置き数量:1,元EMS番号:'EG1',元EMS商品コード:'AAA'},
    {取置ID:'C',状態:'取り置き中',受注番号:'2',商品コード:'AAA',SKU:'AAAb',取り置き数量:1,元EMS番号:'EG1',元EMS商品コード:'AAA'}
  ];
  const key=context.取り置き_行キー_({ban:'1',code:'AAA',sku:'AAAb'});
  const r=context.個別_台帳解除計画_(ledger,key,'個別ボタンからキャンセル','2026-07-16 10:00');
  assert.strictEqual(r.released,1);
  assert.strictEqual(r.qty,1);
  assert.strictEqual(r.rows[0].状態,'手動解除');
  assert.strictEqual(r.rows[1].状態,'発送済み','発送済みは触らない');
  assert.strictEqual(r.rows[2].状態,'取り置き中','他注文は触らない');
});

// ===== コード不一致の名指し救済(説明文コード+P列手動名指し) =====
// 実例: 箱コード「핫 토픽 Hot Topik 2 읽기」+P列10117376 → 注文はHOTOPIK。
// コードがどの注文とも一致しない名指しは、その受注番号の唯一の取り寄せ行へ救済割当する。

test('direct注文: コードがどの行とも一致しなくても単一行の注文なら救済する', () => {
  const orders=[{ban:'10117376',code:'HOTOPIK',sku:'HOTOPIKb',qty:1,sortKey:1,i:0,keys:['HOTOPIKB','HOTOPIK']}];
  const supplies=[{ems:'EG1',code:'핫 토픽 Hot Topik 2 읽기',sourceCode:'핫 토픽 Hot Topik 2 읽기',qty:1,directBan:'10117376',arrival:'2026-07-12'}];
  const result=context.取り置き_割当計算_({orders,ledger:[],movements:[],supplies,explicit:[]});
  assert.strictEqual(result.errors.length,0,JSON.stringify(result.errors));
  assert.strictEqual((result.warnings||[]).length,0,JSON.stringify(result.warnings));
  assert.strictEqual(result.newRows.length,1);
  assert.strictEqual(result.newRows[0].受注番号,'10117376');
  assert.strictEqual(result.newRows[0].商品コード,'HOTOPIK','台帳の商品コードは注文側の実コード');
  assert.ok(result.newRows[0].元EMS商品コード.indexOf('Hot Topik')>=0,'元EMS商品コードは箱の原文を保持');
});

test('direct注文: コード不一致で取り寄せ行が複数ある注文は曖昧なので警告へ', () => {
  const orders=[
    {ban:'10117314',code:'KAGURA06W-PG',sku:'KAGURA06W-PGb',qty:2,sortKey:1,i:0,keys:['KAGURA06W-PG']},
    {ban:'10117314',code:'KAGURA03',sku:'KAGURA03b',qty:1,sortKey:1,i:1,keys:['KAGURA03']}
  ];
  const supplies=[{ems:'EG1',code:'★ポスター カグラバチ 06',sourceCode:'★ポスター カグラバチ 06',qty:2,directBan:'10117314',arrival:'2026-07-12'}];
  const result=context.取り置き_割当計算_({orders,ledger:[],movements:[],supplies,explicit:[]});
  assert.strictEqual(result.errors.length,0);
  assert.strictEqual(result.warnings.length,1,'どの行か特定できない名指しは警告して余りへ');
});

test('未台帳出荷: 名指し箱はコード不一致でもその注文の出荷を差し引ける', () => {
  const supplies=[{ems:'EG1',code:'핫 토픽 읽기',sourceCode:'핫 토픽 읽기',qty:1,directBan:'10117376',arrival:'2026-07-12'}];
  const r=context.取り置き_未台帳出荷計画_(
    [{ban:'10117376',code:'HOTOPIK',sku:'HOTOPIKb',qty:1,基準日:'2026-07-14',メモ:''}],[],[],supplies);
  assert.strictEqual(r.newRows.length,1);
  assert.strictEqual(r.newRows[0].元EMS番号,'EG1');
});

process.exitCode = failures ? 1 : 0;
console.log(failures ? `FAILURES: ${failures}` : 'ALL TESTS PASSED');
