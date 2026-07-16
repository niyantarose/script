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

test('台帳の取り置き中数量がある行はラベンダーで表示する', () => {
  const cfg = { 色_グレー: 'gray', 色_水: 'aqua', 色_橙: 'orange', 色_黄: 'yellow', 色_着: 'lavender' };
  const line = {
    取り置き中数量: 1,
    kbn: '取り寄せ',
    キャンセル: false
  };
  assert.strictEqual(context.引当行状態_(line, cfg, () => false).color, 'lavender');
});

test('実EMS番号・正数量・到着実績が揃った行だけを供給にする', () => {
  const cols={EMS番号:4,コード:2,数量:3};
  const input=[
    ['到着済','2026-07-10','AAA',2,'EG1'],
    ['発送済','',             'BBB',1,'EG2'],
    ['発送済','未到着',       'BBX',1,'EG9'],
    ['到着済','2026-07-10','CCC',1,'棚卸202607'],
    ['到着済','2026-07-10','DDD',0,'EG3']
  ];
  const result=context.EMS供給オブジェクト_(input,cols,row=>row[1]);
  assert.strictEqual(JSON.stringify(result),JSON.stringify([{ems:'EG1',code:'AAA',sourceCode:'AAA',qty:2,arrival:'2026-07-10',directBan:''}]));
});

test('供給のDate型到着日はFIFO比較できる年月日文字列へ正規化する', () => {
  const cols={EMS番号:4,コード:2,数量:3};
  const input=[['到着済',new Date('2026-07-10T00:00:00'),'AAA',2,'EG1']];
  const result=context.EMS供給オブジェクト_(input,cols,row=>row[1]);
  assert.strictEqual(result[0].arrival,'2026-07-10');
});

test('供給日は実在する暦日だけを許可しうるう日境界を正しく扱う', () => {
  const cols={EMS番号:4,コード:2,数量:3};
  const input=[
    ['到着済','2024-02-29','LEAP',1,'EG1'],
    ['到着済','2026-02-28','NORMAL',1,'EG2'],
    ['到着済','2026-02-29','NONLEAP',1,'EG3'],
    ['到着済','2026-02-31','INVALID',1,'EG4']
  ];
  const result=context.EMS供給オブジェクト_(input,cols,row=>row[1]);
  assert.strictEqual(JSON.stringify(result.map(r=>r.code)),JSON.stringify(['LEAP','NORMAL']));
});

test('供給adapterは照合コードとタグ付き元EMS商品コードを分離する', () => {
  const cols={EMS番号:4,コード:2,数量:3};
  const raw='PoEm65（10116569）';
  const result=context.EMS供給オブジェクト_([['到着済','2026-07-10',raw,1,'EG1']],cols,row=>row[1]);
  assert.strictEqual(result[0].code,'POEM65');
  assert.strictEqual(result[0].sourceCode,raw);
  assert.strictEqual(result[0].directBan,'10116569');
});

function 引当計画モック_(allocationPlan) {
  const writes=[];
  const head=['受注番号','氏名','お届け日','商品名','商品コード','商品SKU','項目・選択肢','個数','入金日','入荷日','取り置きメモ'];
  const recvRows=[head,['101','氏名','','AAA','AAA','AAAb','取り寄せ',1,'2026-07-01','','']];
  const recv={
    getDataRange:()=>({getValues:()=>recvRows}),getLastColumn:()=>head.length,getLastRow:()=>recvRows.length,
    getRange:(row,col,numRows,numCols)=>({
      getValues:()=>recvRows.slice(row-1,row-1+numRows).map(r=>r.slice(col-1,col-1+numCols)),
      setValues:()=>writes.push('受注明細'),setBackgrounds:()=>writes.push('受注明細'),setFontSize(){return this;},
      setVerticalAlignment(){return this;},setWrapStrategy(){return this;},setNumberFormat(){return this;}
    })
  };
  const emsRows=[['状態','EMS到着日','商品コード','数量','EMS番号'],['到着済','2026-07-10','AAA',1,'EG1']];
  const emv={getDataRange:()=>({getValues:()=>emsRows})};
  const ui={ButtonSet:{OK:'OK'},alert:()=>{}};
  context.SpreadsheetApp.getActive=()=>({
    getSheetByName:name=>name==='受注明細'?recv:name==='EMS在庫'?emv:null
  });
  context.SpreadsheetApp.getUi=()=>ui;
  context.PropertiesService={getDocumentProperties:()=>({setProperty:()=>writes.push('整合状態'),getProperty:()=>''})};
  context.列マップ_=()=>({hr:1,番号:0,氏名:1,届:2,商品名:3,コード:4,SKU:5,選択肢:6,個数:7,入金:8,入荷:9,EMS:-1,メモ:10,取置数:-1});
  context.取り置き台帳_読む_=()=>[];
  context.EMS在庫移動台帳_読む_=()=>[];
  context.取り置き_集計_=()=>({activeByKey:{},activeRowsByKey:{},usageBySupply:{},errors:[]});
  context.取り置き_行キー_=()=> '101|AAA|AAAB';
  context.取り置き_供給キー_=(ems,code)=>String(ems)+'|'+String(code);
  context.発注共有P列計画_=()=>({error:'',rows:[],writes:[],summary:{}});
  context.P列計画_新規確定割当_=()=>[];
  context.取り置き_割当計算_=()=>allocationPlan;
  context.取り置き台帳_割当計画後行_=()=>[];
  context.取り置き台帳_割当計画を反映_=()=>writes.push('台帳');
  context.発注共有P列計画を反映_=()=>writes.push('P列');
  context.消込台帳更新_=()=>writes.push('消込');
  context.発注共有P列記入_=()=>{writes.push('P列旧書込'); return {};};
  return {writes};
}

test('previewは計画を返し、最初の運用書込みより前に停止する', () => {
  const allocationPlan={orders:[],newRows:[],returnUpdates:[],surplus:[],errors:[]};
  const mock=引当計画モック_(allocationPlan);
  let result,thrown;
  try{ result=context.引当実行_本体_({preview:true}); }catch(e){thrown=e;}
  assert.ifError(thrown);
  assert.strictEqual(JSON.stringify(mock.writes),'[]');
  assert.strictEqual(result.allocationPlan,allocationPlan);
});

test('計画エラーは最初の運用書込みより前に引当を中止する', () => {
  const mock=引当計画モック_({orders:[],newRows:[],returnUpdates:[],surplus:[],errors:['NG']});
  try{ context.引当実行_本体_({preview:false}); }catch(e){}
  assert.strictEqual(JSON.stringify(mock.writes),'[]');
});

test('引当切替差分はpreview後に差分シートだけ書く', () => {
  let previewOption=null;
  const writes=[];
  const current={getDataRange:()=>({getValues:()=>[
    ['受注番号','商品コード','SKU','個数','状態','EMS番号'],
    ['101','AAA','AAAb',1,'在庫待ち','']
  ]})};
  const chain={setValues:values=>{writes.push({sheet:'引当切替差分',values});return chain;},setFontWeight:()=>chain,setBackground:()=>chain,setFontColor:()=>chain};
  const diff={getMaxRows:()=>10,getMaxColumns:()=>10,getRange:()=>chain,setFrozenRows:()=>{}};
  context.引当実行_本体_=options=>{previewOption=options;return {lines:[{
    ban:'101',code:'AAA',sku:'AAAb',qty:1,kbn:'取り寄せ',キャンセル:false,取り置き中数量:1,alloc:0,箱EMS:'EG1'
  }]};};
  context.SpreadsheetApp.getActive=()=>({
    getSheetByName:name=>name==='引当待ち'?current:name==='引当切替差分'?diff:null,
    insertSheet:name=>{writes.push({sheet:name,insert:true});return diff;}
  });
  context.引当切替差分を作成();
  assert.strictEqual(JSON.stringify(previewOption),JSON.stringify({preview:true}));
  assert.ok(writes.length>0);
  assert.ok(writes.every(w=>w.sheet==='引当切替差分'));
});

if (failures) process.exit(1);
