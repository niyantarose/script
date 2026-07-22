// 三段階取り置き台帳ドメインのテスト
// 実行: GASアップスクリプト直下で node tests/project24_three_stage_allocation.test.js
const assert=require('assert');
const fs=require('fs');
const vm=require('vm');

const context={
  console,Date,Set,
  SpreadsheetApp:{openById:()=>({getSheetByName:()=>null})},
  Utilities:{formatDate:()=>''},Logger:{log:()=>{}}
};
vm.createContext(context);
[
  'Project_24/引当.js',
  'Project_24/取り置き計算.js',
  'Project_24/取り置き台帳.js'
].forEach(f=>vm.runInContext(fs.readFileSync(f,'utf8'),context));
const STAGE=vm.runInContext('TORIOKI_STAGE',context);
const LEDGER_HEADERS=vm.runInContext('TORIOKI_CFG.台帳HDR',context);

let failures=0;
function test(name,fn){
  try { fn(); console.log(`PASS ${name}`); }
  catch(error){ failures++; console.error(`FAIL ${name}: ${error.message}`); }
}
function active(id,stage,qty,extra){
  return Object.assign({取置ID:id,状態:'取り置き中',受注番号:'注文14',商品コード:'MRBLUE41',SKU:'MRBLUE41b',取り置き数量:qty,引当段階:stage,取置元種別:'EMS'},extra||{});
}

test('旧14列台帳を読み、保存用ヘッダーだけ三段階21列へ展開する',()=>{
  const additions=['引当段階','EMS到着予定日','現物確認日時','現物確認メモ','供給控除EMS','引当系譜ID','引当系譜数量','供給処理'];
  assert.deepStrictEqual(JSON.parse(JSON.stringify(LEDGER_HEADERS.slice(-8))),additions);
  const legacyHeaders=LEDGER_HEADERS.slice(0,14);
  const legacyValues=['OLD','取り置き中','注文14','MRBLUE41','MRBLUE41b',1,'EMS','EMS-1','MRBLUE41','','2026-07-20','2026-07-20','',''];
  context.SpreadsheetApp.getActive=()=>({getSheetByName:()=>({
    getLastRow:()=>2,getLastColumn:()=>14,
    getRange:(row)=>({getDisplayValues:()=>[legacyHeaders],getValues:()=>[legacyValues]})
  })});
  const row=context.取り置き_表を読む_('取り置き台帳',LEDGER_HEADERS)[0];
  additions.forEach(header=>assert.strictEqual(row[header],''));
});

test('先行から現物へは元EMSを残して供給解放し、後日到着しても消費しない',()=>{
  const ledger=[active('L',STAGE.PLANNED,2,{元EMS番号:'EMS-L'})];
  const plan=context.取り置き_現物確認変換計画_([{受注番号:'注文14',商品コード:'MRBLUE41',SKU:'MRBLUE41b',現物取り置き数量:2}],ledger,[{ems:'EMS-L',code:'MRBLUE41',qty:2,arrival:'2026-07-30'}],new Date('2026-07-22'));
  assert.strictEqual(plan.rows[0].元EMS番号,'EMS-L'); assert.strictEqual(plan.rows[0].供給処理,'供給解放');
  assert.strictEqual(context.取り置き_段階別集計_(plan.rows,[],{'EMS-L':'到着済'}).usageBySupply['EMS-L|MRBLUE41']||0,0);
});
test('到着済から現物へは元EMSを保持してEMS控除する',()=>{
  const ledger=[active('A',STAGE.ARRIVED,2,{元EMS番号:'EMS-1'})];
  const plan=context.取り置き_現物確認変換計画_([{受注番号:'注文14',商品コード:'MRBLUE41',SKU:'MRBLUE41b',現物取り置き数量:2}],ledger,[{ems:'EMS-1',code:'MRBLUE41',qty:2,arrival:'2026-07-20'}],new Date('2026-07-22'));
  assert.strictEqual(plan.rows[0].元EMS番号,'EMS-1');assert.strictEqual(plan.rows[0].供給処理,'EMS控除');
  assert.strictEqual(context.取り置き_段階別集計_(plan.rows,[],{}).usageBySupply['EMS-1|MRBLUE41'],2);
});
test('元EMS不明2個は最古2箱へ分割FIFO控除し、供給不足は元行を維持する',()=>{
  const ledger=[active('U',STAGE.PLANNED,2,{取置元種別:'手動',元EMS番号:''})], input=[{受注番号:'注文14',商品コード:'MRBLUE41',SKU:'MRBLUE41b',現物取り置き数量:2}];
  const plan=context.取り置き_現物確認変換計画_(input,ledger,[{ems:'OLD',code:'MRBLUE41',qty:1,arrival:'2026-07-20'},{ems:'NEW',code:'MRBLUE41',qty:1,arrival:'2026-07-21'}],new Date('2026-07-22'));
  assert.deepStrictEqual(JSON.parse(JSON.stringify(plan.rows.map(r=>[r.元EMS番号,r.供給控除EMS,r.取り置き数量]))),[['','OLD',1],['','NEW',1]]);
  const short=context.取り置き_現物確認変換計画_(input,ledger,[{ems:'OLD',code:'MRBLUE41',qty:1,arrival:'2026-07-20'}],new Date('2026-07-22'));
  assert.deepStrictEqual(JSON.parse(JSON.stringify(short.rows)),JSON.parse(JSON.stringify(ledger)));assert.ok(short.review.length);
});

test('有効行の実効EMS供給が不在なら物理行を維持して停止エラーにする',()=>{
  const row=active('P',STAGE.PHYSICAL,1,{元EMS番号:'MISSING'});
  const checked=context.取り置き_不変条件検証_([{ban:'注文14',code:'MRBLUE41',sku:'MRBLUE41b',qty:1}],[row],[]);
  assert.ok(checked.errors.some(e=>/EMS供給不在/.test(e)));
  assert.strictEqual(context.取り置き_段階別集計_([row],[],{'MISSING':'到着済'}).usageBySupply['MISSING|MRBLUE41'],1);
  assert.strictEqual(row.状態,'取り置き中');
});

test('同一系譜の到着2と現物1は元3以内なので段階重複にしない',()=>{
  const rows=[
    active('A',STAGE.ARRIVED,2,{元EMS番号:'EMS-1',引当系譜ID:'ROOT',引当系譜数量:3}),
    active('P',STAGE.PHYSICAL,1,{元EMS番号:'EMS-1',引当系譜ID:'ROOT',引当系譜数量:3})
  ];
  const result=context.取り置き_不変条件検証_([{ban:'注文14',code:'MRBLUE41',sku:'MRBLUE41b',qty:3}],rows,[{ems:'EMS-1',code:'MRBLUE41',qty:3,arrival:'2026-07-20'}]);
  assert.ok(!result.errors.some(e=>/段階重複/.test(e)));
});

test('同一系譜の有効数量が系譜数量を超えると段階重複で停止する',()=>{
  const rows=[
    active('A',STAGE.ARRIVED,2,{元EMS番号:'EMS-1',引当系譜ID:'ROOT',引当系譜数量:3}),
    active('P',STAGE.PHYSICAL,2,{元EMS番号:'EMS-1',引当系譜ID:'ROOT',引当系譜数量:3})
  ];
  const result=context.取り置き_不変条件検証_([{ban:'注文14',code:'MRBLUE41',sku:'MRBLUE41b',qty:4}],rows,[{ems:'EMS-1',code:'MRBLUE41',qty:4,arrival:'2026-07-20'}]);
  assert.ok(result.errors.some(e=>/段階重複/.test(e)));
});

test('異なる取置IDは未設定時に別系譜として扱う',()=>{
  const rows=[active('A',STAGE.ARRIVED,1,{元EMS番号:'EMS-1'}),active('B',STAGE.PHYSICAL,1,{元EMS番号:'EMS-1'})];
  const result=context.取り置き_不変条件検証_([{ban:'注文14',code:'MRBLUE41',sku:'MRBLUE41b',qty:2}],rows,[{ems:'EMS-1',code:'MRBLUE41',qty:2,arrival:'2026-07-20'}]);
  assert.ok(!result.errors.some(e=>/段階重複/.test(e)));
});

test('同じ注文SKUの現物確認入力が重複すると後勝ちにせずエラーにする',()=>{
  const ledger=[active('L',STAGE.PLANNED,2,{元EMS番号:'EMS-1'})];
  const plan=context.取り置き_現物確認変換計画_([
    {受注番号:'注文14',商品コード:'MRBLUE41',SKU:'MRBLUE41b',現物取り置き数量:1},
    {受注番号:'注文14',商品コード:'MRBLUE41',SKU:'MRBLUE41b',現物取り置き数量:2}
  ],ledger,[{ems:'EMS-1',code:'MRBLUE41',qty:2,arrival:'2026-07-30'}],new Date());
  assert.ok(plan.errors.some(e=>/重複入力/.test(e)));
  assert.deepStrictEqual(JSON.parse(JSON.stringify(plan.rows)),JSON.parse(JSON.stringify(ledger)));
});

test('同じ注文SKUの現物・到着済・先行を段階別に合計する',()=>{
  const rows=[
    active('P',STAGE.PHYSICAL,1),
    active('A',STAGE.ARRIVED,1,{元EMS番号:'EMS-A'}),
    active('L',STAGE.PLANNED,1,{元EMS番号:'EMS-L'})
  ];
  const key=context.取り置き_行キー_(rows[0]);
  const result=context.取り置き_段階別集計_(rows,[],{'EMS-A':'到着済','EMS-L':'未着'}).byKey[key];
  assert.strictEqual(result.現物確認済み数量,1);
  assert.strictEqual(result.到着済引当数量,1);
  assert.strictEqual(result.先行引当数量,1);
  assert.strictEqual(result.合計確保数量,3);
  assert.strictEqual(result.行内訳.length,3);
});

test('実例10117608は現物6と先行8を二重計上しない',()=>{
  const rows=[
    Object.assign(active('P',STAGE.PHYSICAL,6),{受注番号:'10117608'}),
    Object.assign(active('L',STAGE.PLANNED,8,{元EMS番号:'EMS-L'}),{受注番号:'10117608'})
  ];
  const result=context.取り置き_段階別集計_(rows,[],{'EMS-L':'未着'}).byKey[context.取り置き_行キー_(rows[0])];
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result)),{
    現物確認済み数量:6,到着済引当数量:0,先行引当数量:8,合計確保数量:14,行内訳:[
      {取置ID:'P',引当段階:'現物確認済み',数量:6},
      {取置ID:'L',引当段階:'先行',数量:8}
    ]
  });
});

test('既存EMS行の現物確認は数量を増やさず段階だけを変換する',()=>{
  const ledger=[active('EMS-1',STAGE.ARRIVED,2,{元EMS番号:'EMS-1'})];
  const plan=context.取り置き_現物確認変換計画_([
    {受注番号:'注文14',商品コード:'MRBLUE41',SKU:'MRBLUE41b',現物取り置き数量:2,現物確認メモ:'棚A'}
  ],ledger,[{ems:'EMS-1',code:'MRBLUE41',qty:2,arrival:'2026-07-20'}],new Date('2026-07-22T00:00:00Z'));
  assert.deepStrictEqual(JSON.parse(JSON.stringify(plan.errors)),[]);
  assert.strictEqual(plan.rows.length,1);
  assert.strictEqual(plan.rows[0].取り置き数量,2);
  assert.strictEqual(plan.rows[0].引当段階,STAGE.PHYSICAL);
  assert.strictEqual(plan.rows[0].元EMS番号,'EMS-1');
  assert.strictEqual(plan.rows[0].現物確認メモ,'棚A');
});

test('現物確認済み数量を減らす入力は解除フローへ誘導して台帳を変えない',()=>{
  const ledger=[active('P',STAGE.PHYSICAL,3)];
  const plan=context.取り置き_現物確認変換計画_([
    {受注番号:'注文14',商品コード:'MRBLUE41',SKU:'MRBLUE41b',現物取り置き数量:2}
  ],ledger,[],new Date());
  assert.ok(plan.errors.some(e=>/解除/.test(e)));
  assert.deepStrictEqual(JSON.parse(JSON.stringify(plan.rows)),JSON.parse(JSON.stringify(ledger)));
});

test('注文超過・同一取置ID重複・EMS供給超過は停止エラーにする',()=>{
  const rows=[
    active('DUP',STAGE.ARRIVED,2,{元EMS番号:'EMS-1'}),
    active('DUP',STAGE.ARRIVED,1,{元EMS番号:'EMS-1'})
  ];
  const result=context.取り置き_不変条件検証_(
    [{ban:'注文14',code:'MRBLUE41',sku:'MRBLUE41b',qty:2}],rows,[{ems:'EMS-1',code:'MRBLUE41',qty:2,arrival:'2026-07-20'}]
  );
  assert.ok(result.errors.some(e=>/注文数を超過/.test(e)));
  assert.ok(result.errors.some(e=>/取置ID重複/.test(e)));
  assert.ok(result.errors.some(e=>/EMS供給超過/.test(e)));
});

test('現物固定の注文消滅・数量減は自動解除せず停止エラーにする',()=>{
  const physical=active('P',STAGE.PHYSICAL,3);
  const missing=context.取り置き_不変条件検証_([], [physical], []);
  const reduced=context.取り置き_不変条件検証_([{ban:'注文14',code:'MRBLUE41',sku:'MRBLUE41b',qty:2}],[physical],[]);
  assert.ok(missing.errors.some(e=>/現物固定.*注文が消滅/.test(e)));
  assert.ok(reduced.errors.some(e=>/現物固定.*数量減/.test(e)));
});

test('旧開始前在庫は要移行となり通常の段階別集計へ混ぜない',()=>{
  const row={取置ID:'OLD',状態:'取り置き中',受注番号:'注文14',商品コード:'MRBLUE41',SKU:'MRBLUE41b',取り置き数量:1,取置元種別:'開始前在庫'};
  assert.strictEqual(context.取り置き_段階正規化_(row,{}).引当段階,'要移行');
  const result=context.取り置き_段階別集計_([row],[],{});
  assert.strictEqual(result.activeByKey[context.取り置き_行キー_(row)]||0,0);
  assert.strictEqual(result.要移行行.length,1);
});

test('先行2個を現物2個へ変換すると未着EMS供給を解放し台帳合計は維持する',()=>{
  const ledger=[active('L',STAGE.PLANNED,2,{元EMS番号:'EMS-L'})];
  const plan=context.取り置き_現物確認変換計画_([
    {受注番号:'注文14',商品コード:'MRBLUE41',SKU:'MRBLUE41b',現物取り置き数量:2}
  ],ledger,[{ems:'EMS-L',code:'MRBLUE41',qty:2,arrival:'2026-07-30'}],new Date());
  const summary=context.取り置き_段階別集計_(plan.rows,[],{'EMS-L':'未着'});
  assert.strictEqual(summary.activeByKey[context.取り置き_行キー_(ledger[0])],2);
  assert.strictEqual(summary.usageBySupply['EMS-L|MRBLUE41']||0,0);
});

test('元EMS不明の現物は同SKU最古到着済EMSから控除し、なければ要確認にする',()=>{
  const ledger=[active('UNKNOWN',STAGE.PLANNED,1,{取置元種別:'手動',元EMS番号:''})];
  const supplies=[
    {ems:'NEW',code:'MRBLUE41',qty:1,arrival:'2026-07-21'},
    {ems:'OLD',code:'MRBLUE41',qty:1,arrival:'2026-07-20'}
  ];
  const plan=context.取り置き_現物確認変換計画_([
    {受注番号:'注文14',商品コード:'MRBLUE41',SKU:'MRBLUE41b',現物取り置き数量:1}
  ],ledger,supplies,new Date());
  assert.strictEqual(plan.rows[0].元EMS番号,'');
  assert.strictEqual(plan.rows[0].供給控除EMS,'OLD');
  assert.strictEqual(context.取り置き_集計_(plan.rows,[],{'OLD':'到着済'}).usageBySupply['OLD|MRBLUE41'].取り置き中,1);
  const noSupply=context.取り置き_現物確認変換計画_([
    {受注番号:'注文14',商品コード:'MRBLUE41',SKU:'MRBLUE41b',現物取り置き数量:1}
  ],ledger,[],new Date());
  assert.ok(noSupply.review.some(r=>/供給控除EMS/.test(r.理由)));
});

if(failures){ console.error(`\n${failures} TEST(S) FAILED`); process.exit(1); }
console.log('\nALL PASS');
