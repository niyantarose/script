// 一度きりの現物確認移行(旧開始前在庫の仕分け+消えた確保の復元)の純粋ロジックテスト
// 実行: GASアップスクリプト直下で node tests/project24_physical_migration.test.js
const assert=require('assert');
const fs=require('fs');
const vm=require('vm');

const context={
  console,Date,Set,Map,JSON,Math,
  SpreadsheetApp:{openById:()=>({getSheetByName:()=>null})},
  Utilities:{formatDate:()=>''},Logger:{log:()=>{}}
};
vm.createContext(context);
['Project_24/引当.js','Project_24/取り置き計算.js','Project_24/取り置き台帳.js','Project_24/現物確認移行.js']
  .forEach(f=>vm.runInContext(fs.readFileSync(f,'utf8'),context));

let failures=0;
function test(name,fn){
  try { fn(); console.log(`PASS ${name}`); }
  catch(error){ failures++; console.error(`FAIL ${name}: ${error.message}`); }
}
const json=v=>JSON.parse(JSON.stringify(v));

const 開始前行=(id,ban,code,sku,qty)=>({取置ID:id,状態:'取り置き中',受注番号:ban,商品コード:code,SKU:sku,
  取り置き数量:qty,取置元種別:'開始前在庫',元EMS番号:'',登録日時:'2026-07-17'});
const EMS行=(id,ban,code,sku,qty,stage)=>({取置ID:id,状態:'取り置き中',受注番号:ban,商品コード:code,SKU:sku,
  取り置き数量:qty,取置元種別:'EMS',引当段階:stage||'到着済',元EMS番号:'EG050152967KR'});

const currentLedger=[
  開始前行('INIT|500|HOLD-01|HOLD-01B','500','HOLD-01','HOLD-01b',2),
  EMS行('REBUILD|ACTIVE|10117428|3|EG050152967KR|2|1','10117428','MRBLUE42','MRBLUE42-6b',2,'先行')
];
const baselineLedger=[
  {取置ID:'EMS|OLD|10117428','状態':'取り置き中',受注番号:'10117428',商品コード:'MRBLUE42',SKU:'MRBLUE42-6b',取り置き数量:3,取置元種別:'EMS',元EMS番号:'EG049882819KR'},
  {取置ID:'EMS|OLD|900','状態':'取り置き中',受注番号:'900',商品コード:'GONE01',SKU:'GONE01b',取り置き数量:1,取置元種別:'EMS',元EMS番号:'EG0'}
];
const currentOrders=[
  {ban:'10117428',sku:'MRBLUE42-6b',code:'MRBLUE42',qty:3},
  {ban:'500',sku:'HOLD-01b',code:'HOLD-01',qty:2}
];

test('候補計算: 旧開始前在庫は現状のまま候補化し、数量を勝手に現物へ変えない', () => {
  const c=context.現物確認移行_候補計算_(currentLedger,baselineLedger,currentOrders);
  const hold=c.filter(x=>x.種別==='旧開始前在庫');
  assert.strictEqual(hold.length,1);
  assert.strictEqual(hold[0].受注番号,'500');
  assert.strictEqual(hold[0].現在,2);
  assert.strictEqual(hold[0].選択,'', '既定は未選択(自動で現物にしない)');
});

test('候補計算: 10117428/MRBLUE42-6b は以前3・現在2で差1を復元候補へ出す', () => {
  const c=context.現物確認移行_候補計算_(currentLedger,baselineLedger,currentOrders);
  const lost=c.find(x=>x.種別==='消えた確保'&&x.受注番号==='10117428');
  assert.ok(lost);
  assert.strictEqual(lost.以前,3);
  assert.strictEqual(lost.現在,2);
  assert.strictEqual(lost.差,1);
});

test('候補計算: 現役注文に無い受注の幽霊確保は候補にしない', () => {
  const c=context.現物確認移行_候補計算_(currentLedger,baselineLedger,currentOrders);
  assert.ok(!c.some(x=>x.受注番号==='900'), '注文900は現役でないため候補外');
});

test('反映計画: 現物確認済みにするは段階変換+供給解放で、注文数量を超える復元はその键だけエラー', () => {
  const candidates=context.現物確認移行_候補計算_(currentLedger,baselineLedger,currentOrders);
  const choices={};
  choices[context.取り置き_行キー_({ban:'500',sku:'HOLD-01b'})]={選択:'現物確認済みにする'};
  choices[context.取り置き_行キー_({ban:'10117428',sku:'MRBLUE42-6b'})]={選択:'現物確認済みにする',数量:2}; // 差1を超える
  const plan=context.現物確認移行_反映計画_(candidates,choices,currentLedger,currentOrders,[],new Date('2026-07-22'));
  // 500は適用: INIT行が現物確認済みへ変換され供給解放
  const conv=plan.rows.find(r=>r.取置ID==='INIT|500|HOLD-01|HOLD-01B');
  assert.strictEqual(conv.引当段階,'現物確認済み');
  assert.strictEqual(conv.供給処理,'供給解放');
  assert.ok(conv.現物確認日時);
  // 10117428は超過エラー(2+既存2=4>注文3)だが、500の適用は破棄されない
  assert.ok(plan.errors.some(e=>/10117428/.test(e)));
  assert.ok(!plan.rows.some(r=>String(r.取置ID||'').indexOf('MIG|')===0&&r.受注番号==='10117428'));
});

test('反映計画: 差1の復元は新規現物行(供給解放)を作り既存EMS行と二重加算しない', () => {
  const candidates=context.現物確認移行_候補計算_(currentLedger,baselineLedger,currentOrders);
  const choices={};
  choices[context.取り置き_行キー_({ban:'10117428',sku:'MRBLUE42-6b'})]={選択:'現物確認済みにする',数量:1};
  const plan=context.現物確認移行_反映計画_(candidates,choices,currentLedger,currentOrders,[],new Date('2026-07-22'));
  assert.deepStrictEqual(json(plan.errors),[]);
  const mig=plan.rows.filter(r=>String(r.取置ID||'').indexOf('MIG|')===0);
  assert.strictEqual(mig.length,1);
  assert.strictEqual(mig[0].取り置き数量,1);
  assert.strictEqual(mig[0].引当段階,'現物確認済み');
  assert.strictEqual(mig[0].供給処理,'供給解放');
  // 既存の先行EMS行2個はそのまま(合計3=注文数量)
  const ems=plan.rows.find(r=>String(r.取置ID||'').indexOf('REBUILD|ACTIVE|10117428')===0);
  assert.strictEqual(ems.取り置き数量,2);
});

test('反映計画: 解除は旧開始前在庫だけを手動解除にし、EMS由来は解除しない', () => {
  const candidates=context.現物確認移行_候補計算_(currentLedger,baselineLedger,currentOrders);
  const choices={};
  choices[context.取り置き_行キー_({ban:'500',sku:'HOLD-01b'})]={選択:'解除'};
  const plan=context.現物確認移行_反映計画_(candidates,choices,currentLedger,currentOrders,[],new Date('2026-07-22'));
  const released=plan.rows.find(r=>r.取置ID==='INIT|500|HOLD-01|HOLD-01B');
  assert.strictEqual(released.状態,'手動解除');
  const ems=plan.rows.find(r=>String(r.取置ID||'').indexOf('REBUILD|ACTIVE|10117428')===0);
  assert.strictEqual(ems.状態,'取り置き中', 'EMS行は触らない');
});

test('反映計画: 保留は台帳を変えず保留一覧として残る', () => {
  const candidates=context.現物確認移行_候補計算_(currentLedger,baselineLedger,currentOrders);
  const choices={};
  choices[context.取り置き_行キー_({ban:'500',sku:'HOLD-01b'})]={選択:'保留'};
  const plan=context.現物確認移行_反映計画_(candidates,choices,currentLedger,currentOrders,[],new Date('2026-07-22'));
  assert.deepStrictEqual(json(plan.rows),json(currentLedger));
  assert.strictEqual(plan.保留.length,1);
});

test('反映計画: 不正な選択値はその键だけエラーで他は適用される', () => {
  const candidates=context.現物確認移行_候補計算_(currentLedger,baselineLedger,currentOrders);
  const choices={};
  choices[context.取り置き_行キー_({ban:'500',sku:'HOLD-01b'})]={選択:'現物確認済みにする'};
  choices[context.取り置き_行キー_({ban:'10117428',sku:'MRBLUE42-6b'})]={選択:'???'};
  const plan=context.現物確認移行_反映計画_(candidates,choices,currentLedger,currentOrders,[],new Date('2026-07-22'));
  assert.ok(plan.errors.some(e=>/10117428/.test(e)));
  assert.strictEqual(plan.rows.find(r=>r.取置ID==='INIT|500|HOLD-01|HOLD-01B').引当段階,'現物確認済み');
});

test('基準シート(旧14列コピー)を21列ヘッダーで読める(見出し不足で落ちない)', () => {
  const HDR=vm.runInContext('TORIOKI_CFG.台帳HDR',context);
  const legacy=Array.prototype.slice.call(HDR,0,14);
  const data=[legacy,['OLD','取り置き中','900','GONE01','GONE01b',1,'EMS','EG0','GONE01','','2026-07-17','2026-07-17','','']];
  const sheet={getLastRow:()=>data.length,getLastColumn:()=>legacy.length,
    getRange:(r)=>({getDisplayValues:()=>[data[0]],getValues:()=>data.slice(1)})};
  const orig=context.SpreadsheetApp;
  context.SpreadsheetApp={getActive:()=>({getSheetByName:(n)=>n==='取り置き台帳_全件再計算前_20260721_123350'?sheet:null})};
  try{
    const rows=context.取り置き_表を読む_('取り置き台帳_全件再計算前_20260721_123350',HDR);
    assert.strictEqual(rows.length,1);
    assert.strictEqual(rows[0].取置ID,'OLD');
    assert.strictEqual(rows[0].引当段階,'','新列は空として返す');
  } finally { context.SpreadsheetApp=orig; }
});

if (failures) process.exit(1);
