// 取り置き登録の入力永続化(非表示シート保存)と表示対象判定の純粋ロジックテスト
// 実行: GASアップスクリプト直下で node tests/project24_reservation_work_screen.test.js
const assert=require('assert');
const fs=require('fs');
const vm=require('vm');

const context={
  console,Date,Set,Map,JSON,Math,
  SpreadsheetApp:{openById:()=>({getSheetByName:()=>null})},
  Utilities:{formatDate:()=>''},Logger:{log:()=>{}}
};
vm.createContext(context);
['Project_24/引当.js','Project_24/取り置き計算.js','Project_24/取り置き台帳.js']
  .forEach(f=>vm.runInContext(fs.readFileSync(f,'utf8'),context));

let failures=0;
function test(name,fn){
  try { fn(); console.log(`PASS ${name}`); }
  catch(error){ failures++; console.error(`FAIL ${name}: ${error.message}`); }
}
const json=v=>JSON.parse(JSON.stringify(v));
const gen=(over)=>Object.assign({取置ID:'INIT|700|AAA|AAAB',受注番号:'700',商品コード:'AAA',SKU:'AAAb',
  注文数量:1,棚確認:'',現物取り置き数量:'',メモ:''},over||{});

test('入力保存マージ: シート入力(棚確認/メモ/数量)をstoreへupsertし生成行へ復元する', () => {
  const sheetRows=[gen({棚確認:'部分在庫',現物取り置き数量:1,メモ:'棚A'})];
  const out=context.取り置き_入力保存マージ_([gen()],[],sheetRows,'2026-07-22 12:00:00');
  assert.strictEqual(out.rows[0].棚確認,'部分在庫');
  assert.strictEqual(String(out.rows[0].現物取り置き数量),'1');
  assert.strictEqual(out.rows[0].メモ,'棚A');
  const key=context.取り置き_入力キー_(sheetRows[0]);
  assert.strictEqual(String(out.store[key].未反映現物確認数量),'1');
  assert.strictEqual(out.store[key].棚確認,'部分在庫');
});

test('入力保存マージ: シートが空でも保存済み値から復元する(洗い替え耐性)', () => {
  const key=context.取り置き_入力キー_(gen());
  const saved=[{入力キー:key,受注番号:'700',SKU:'AAAb',商品コード:'AAA',棚確認:'未着',取り置きメモ:'保存済み',未反映現物確認数量:2}];
  const out=context.取り置き_入力保存マージ_([gen()],saved,[],'2026-07-22 12:00:00');
  assert.strictEqual(out.rows[0].棚確認,'未着');
  assert.strictEqual(out.rows[0].メモ,'保存済み');
  assert.strictEqual(String(out.rows[0].現物取り置き数量),'2');
});

test('入力保存マージ: 生成行に無いキーのstoreも消えない(注文が一時的に消えても保持)', () => {
  const saved=[{入力キー:'999|GONE',受注番号:'999',SKU:'GONEb',商品コード:'GONE',棚確認:'部分在庫'}];
  const out=context.取り置き_入力保存マージ_([gen()],saved,[],'2026-07-22 12:00:00');
  assert.ok(out.store['999|GONE']);
  assert.strictEqual(out.store['999|GONE'].棚確認,'部分在庫');
});

test('入力保存マージ: シートの新しい入力が保存済みの古い値を上書きする', () => {
  const key=context.取り置き_入力キー_(gen());
  const saved=[{入力キー:key,棚確認:'未着',取り置きメモ:'旧'}];
  const sheetRows=[gen({棚確認:'部分在庫',メモ:'新'})];
  const out=context.取り置き_入力保存マージ_([gen()],saved,sheetRows,'2026-07-22 12:00:00');
  assert.strictEqual(out.store[key].棚確認,'部分在庫');
  assert.strictEqual(out.store[key].取り置きメモ,'新');
  assert.strictEqual(out.rows[0].棚確認,'部分在庫');
});

test('作業対象判定: 部分在庫は要作業に出る・全未着の希望日待ちは出ない・すべては全部出る', () => {
  const 部分注文=[{現在の状態:'部分在庫',判定:'',要対応:'',台帳確保数:0}];
  const 未着希望=[{現在の状態:'希望日待ち',判定:'',要対応:'',台帳確保数:0,棚確認:'未着'}];
  const 現物希望=[{現在の状態:'希望日待ち',判定:'',要対応:'',台帳確保数:1}];
  assert.strictEqual(context.取り置き_作業対象判定_(部分注文,'要作業'),true);
  assert.strictEqual(context.取り置き_作業対象判定_(未着希望,'要作業'),false);
  assert.strictEqual(context.取り置き_作業対象判定_(現物希望,'要作業'),true,'現物ありの希望日待ちは棚を見る対象');
  assert.strictEqual(context.取り置き_作業対象判定_(未着希望,'すべて'),true);
  assert.strictEqual(context.取り置き_作業対象判定_(部分注文,'部分在庫'),true);
  assert.strictEqual(context.取り置き_作業対象判定_(未着希望,'部分在庫'),false);
});

test('作業対象判定: 要棚確認・棚戻し待ち・要対応ありは常に要作業', () => {
  assert.strictEqual(context.取り置き_作業対象判定_([{現在の状態:'出荷可能',判定:'要棚確認'}],'要作業'),true);
  assert.strictEqual(context.取り置き_作業対象判定_([{現在の状態:'キャンセル戻し',判定:'棚戻し待ち'}],'要作業'),true);
  assert.strictEqual(context.取り置き_作業対象判定_([{現在の状態:'出荷可能',判定:'',要対応:'棚へ戻す 1個'}],'要作業'),true);
  assert.strictEqual(context.取り置き_作業対象判定_([{現在の状態:'出荷可能',判定:'',要対応:''}],'要作業'),false,'確認済み出荷可能は出ない');
});

if (failures) process.exit(1);
