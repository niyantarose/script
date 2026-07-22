// 引当状況一覧とGoQ差分の純粋ロジックテスト
// 実行: GASアップスクリプト直下で node tests/project24_allocation_status_list.test.js
const assert=require('assert');
const fs=require('fs');
const vm=require('vm');

const context={
  console,Date,Set,Map,JSON,Math,
  SpreadsheetApp:{openById:()=>({getSheetByName:()=>null})},
  Utilities:{formatDate:()=>''},Logger:{log:()=>{}}
};
vm.createContext(context);
['Project_24/引当.js','Project_24/取り置き計算.js','Project_24/引当状況一覧.js']
  .forEach(f=>vm.runInContext(fs.readFileSync(f,'utf8'),context));

let failures=0;
function test(name,fn){
  try { fn(); console.log(`PASS ${name}`); }
  catch(error){ failures++; console.error(`FAIL ${name}: ${error.message}`); }
}
function line(over){
  return Object.assign({kbn:'取り寄せ',qty:1,届:'',キャンセル:false,段階付与:true,
    現物確認済み数量:0,到着済引当数量:0,先行引当数量:0,別ルート済数量:0,氏名:'',code:'AAA',sku:'AAAb',商品名:'',メモ:''},over||{});
}
const FUTURE='2999/12/31';

test('全数到着済なのにGoQが取り寄せ中なら差異あり(到着済み全数)', () => {
  const arr=[line({到着済引当数量:1})];
  const rec=context.引当状況_推奨GoQ_(arr,true,false);
  assert.strictEqual(rec,'出荷可能');
  const d=context.引当状況_GoQ差分_('出荷待/取寄せ',rec,arr,true,false);
  assert.strictEqual(d.判定,'差異あり');
  assert.ok(d.理由.indexOf('到着済み全数')>=0);
});

test('全数先行は推奨に先行を含め到着済と誤表示しない', () => {
  const arr=[line({先行引当数量:1})];
  const rec=context.引当状況_推奨GoQ_(arr,true,false);
  assert.ok(/先行/.test(rec), rec);
  const d=context.引当状況_GoQ差分_('出荷待/取寄せ',rec,arr,true,false);
  assert.ok(d.判定==='差異あり'?d.理由.indexOf('未着先行のみ')>=0:true);
});

test('希望日・未入金・代引き・部分在庫は別々の推奨/理由になる', () => {
  assert.strictEqual(context.引当状況_推奨GoQ_([line({到着済引当数量:1,届:FUTURE})],true,false),'希望日待ち');
  assert.strictEqual(context.引当状況_推奨GoQ_([line({到着済引当数量:1})],false,false),'入金待ち');
  assert.strictEqual(context.引当状況_推奨GoQ_([line({到着済引当数量:1})],false,true),'出荷可能(代引き)');
  assert.ok(/^一部確保/.test(context.引当状況_推奨GoQ_([line({qty:2,到着済引当数量:1})],true,false)));
});

test('一覧行: 同じ注文は1分類で商品明細は全部出る', () => {
  const linesByBan={'700':[line({到着済引当数量:1}),line({code:'BBB',sku:'BBBb',先行引当数量:1})]};
  const rows=context.引当状況_一覧行_(linesByBan,{'700':true},{},{'700':'出荷待/取寄せ'});
  assert.strictEqual(rows.length,2);
  assert.strictEqual(rows[0].分類,rows[1].分類);
  assert.strictEqual(rows[0].受注番号,'700');
  assert.ok(rows[1].状態の理由);
});

test('一覧生成はGoQ更新や書き込みAPIを呼ばない(純粋)', () => {
  let touched=false;
  const spy=new Proxy({},{get(){touched=true;return()=>{};}});
  const orig=context.SpreadsheetApp; context.SpreadsheetApp=spy;
  try{ context.引当状況_一覧行_({'1':[line({})]},{},{},{}); }
  finally{ context.SpreadsheetApp=orig; }
  assert.strictEqual(touched,false);
});

if (failures) process.exit(1);
