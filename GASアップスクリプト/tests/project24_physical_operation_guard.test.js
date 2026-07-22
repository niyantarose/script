// 先行引当を物理オペ(納品書ピック・⑤箱締め・Yahoo出力)から遮断する境界の純粋ロジックテスト
// 実行: GASアップスクリプト直下で node tests/project24_physical_operation_guard.test.js
const assert=require('assert');
const fs=require('fs');
const vm=require('vm');

const context={
  console,Date,Set,Map,JSON,Math,
  SpreadsheetApp:{openById:()=>({getSheetByName:()=>null})},
  Utilities:{formatDate:()=>''},Logger:{log:()=>{}}
};
vm.createContext(context);
['Project_24/引当.js','Project_24/取り置き計算.js'].forEach(f=>vm.runInContext(fs.readFileSync(f,'utf8'),context));
const STAGE=vm.runInContext('TORIOKI_STAGE',context);

let failures=0;
function test(name,fn){
  try { fn(); console.log(`PASS ${name}`); }
  catch(error){ failures++; console.error(`FAIL ${name}: ${error.message}`); }
}
function row(over){
  return Object.assign({取置ID:'X',状態:'取り置き中',受注番号:'901',商品コード:'AAA',SKU:'AAAb',
    取り置き数量:1,取置元種別:'EMS',元EMS番号:'EMS-1'},over||{});
}

test('物理オペ対象行: 到着済・現物確認済み・要移行(旧開始前在庫)だけtrue', () => {
  assert.strictEqual(context.取り置き_物理オペ対象行_(row({引当段階:STAGE.ARRIVED}),{}), true);
  assert.strictEqual(context.取り置き_物理オペ対象行_(row({引当段階:STAGE.PHYSICAL}),{}), true);
  assert.strictEqual(context.取り置き_物理オペ対象行_(row({取置元種別:'開始前在庫',元EMS番号:''}),{}), true, '棚の現物(要移行)は物理');
  assert.strictEqual(context.取り置き_物理オペ対象行_(row({引当段階:STAGE.PLANNED}),{}), false, '先行は帳簿のみ');
  assert.strictEqual(context.取り置き_物理オペ対象行_(row({引当段階:STAGE.ARRIVED,状態:'発送済み'}),{}), false);
});

test('物理オペ対象行: 段階未記入の旧EMS行は未着証拠がある時だけ除外する', () => {
  assert.strictEqual(context.取り置き_物理オペ対象行_(row({}),{}), true, '状態不明の旧箱=物理側へ倒す');
  assert.strictEqual(context.取り置き_物理オペ対象行_(row({}),{'EMS-1':'未着'}), false, '箱が未着なら物理に無い');
});

test('便締め先行残検査: 対象EMSに先行行が残れば④昇格を促す理由を返し、無ければ空', () => {
  const ledger=[
    row({取置ID:'P1',引当段階:STAGE.PLANNED,元EMS番号:'EMS-CLOSE',受注番号:'10117428',商品コード:'MRBLUE42'}),
    row({取置ID:'A1',引当段階:STAGE.ARRIVED,元EMS番号:'EMS-CLOSE'}),
    row({取置ID:'P2',引当段階:STAGE.PLANNED,元EMS番号:'EMS-OTHER'})
  ];
  const reason=context.取り置き_便締め先行残検査_(ledger,new Set(['EMS-CLOSE']),{});
  assert.ok(/先行引当が1行/.test(reason), '対象便の先行だけ数える: '+reason);
  assert.ok(/④/.test(reason), '④実行での昇格を案内する');
  assert.ok(/10117428/.test(reason));
  assert.strictEqual(context.取り置き_便締め先行残検査_(ledger,new Set(['EMS-ARRIVED-ONLY']),{}), '');
  const arrivedOnly=[row({取置ID:'A2',引当段階:STAGE.ARRIVED,元EMS番号:'EMS-CLOSE'})];
  assert.strictEqual(context.取り置き_便締め先行残検査_(arrivedOnly,new Set(['EMS-CLOSE']),{}), '');
});

test('物理出荷対象注文: 全数先行はfalse・到着済+現物で全数はtrue(分類とは独立)', () => {
  const line=(over)=>Object.assign({kbn:'取り寄せ',qty:1,届:'',キャンセル:false,段階付与:true,
    現物確認済み数量:0,到着済引当数量:0,先行引当数量:0,別ルート済数量:0},over||{});
  assert.strictEqual(context.注文物理出荷可_([line({先行引当数量:1})]), false);
  assert.strictEqual(context.注文物理出荷可_([line({現物確認済み数量:1}),line({到着済引当数量:1})]), true);
});

if (failures) process.exit(1);
