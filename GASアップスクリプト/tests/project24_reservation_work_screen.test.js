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
  注文数量:1,棚確認:'',追加数量:'',メモ:''},over||{});

test('入力保存マージ: シート入力(棚確認/メモ/数量)をstoreへupsertし生成行へ復元する', () => {
  const sheetRows=[gen({棚確認:'部分在庫',追加数量:1,メモ:'棚A'})];
  const out=context.取り置き_入力保存マージ_([gen()],[],sheetRows,'2026-07-22 12:00:00');
  assert.strictEqual(out.rows[0].棚確認,'部分在庫');
  assert.strictEqual(String(out.rows[0].追加数量),'1');
  assert.strictEqual(out.rows[0].メモ,'棚A');
  const key=context.取り置き_入力キー_(sheetRows[0]);
  assert.strictEqual(String(out.store[key].未反映追加数量),'1');
  assert.strictEqual(out.store[key].棚確認,'部分在庫');
});

test('入力保存マージ: シートが空でも保存済み値から復元する(洗い替え耐性)', () => {
  const key=context.取り置き_入力キー_(gen());
  const saved=[{入力キー:key,受注番号:'700',SKU:'AAAb',商品コード:'AAA',棚確認:'未着',取り置きメモ:'保存済み',未反映追加数量:2}];
  const out=context.取り置き_入力保存マージ_([gen()],saved,[],'2026-07-22 12:00:00');
  assert.strictEqual(out.rows[0].棚確認,'未着');
  assert.strictEqual(out.rows[0].メモ,'保存済み');
  assert.strictEqual(String(out.rows[0].追加数量),'2');
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

test('統合反映: 1入力キーのエラーでも他の有効キーは適用される(全行破棄の廃止)', () => {
  const inputs=[
    {取置ID:'INIT|700|AAA|AAAB',受注番号:'700',商品コード:'AAA',SKU:'AAAb',注文数量:1,追加数量:1,棚確認:'',メモ:''},
    {取置ID:'INIT|701|BBB|BBBB',受注番号:'701',商品コード:'BBB',SKU:'BBBb',注文数量:1,追加数量:5,棚確認:'',メモ:''}
  ];
  const plan=context.取り置き_統合反映計画_(inputs,[],new Date('2026-07-22'));
  assert.ok(plan.errors.some(e=>/701/.test(e)),'701の超過はエラー');
  assert.ok(plan.rows.some(r=>r.取置ID==='INIT|700|AAA|AAAB'&&r.取り置き数量===1),'700は適用');
  assert.ok(!plan.rows.some(r=>r.受注番号==='701'),'701は未適用');
  assert.strictEqual(plan.counts.取り置き行,1,'countsは適用分だけ');
});

test('統合反映: 棚戻しの不正選択はそのIDだけエラーで通常行の適用を妨げない', () => {
  const ledger=[{取置ID:'RET1',状態:'キャンセル戻し',受注番号:'800',商品コード:'CCC',SKU:'CCCb',取り置き数量:1,取置元種別:'EMS',元EMS番号:'EG1',戻し処理結果:'未確認'}];
  const inputs=[
    {取置ID:'RET1',受注番号:'800',商品コード:'CCC',SKU:'CCCb',注文数量:1,追加数量:'',棚確認:'',メモ:'',処理:'変な値'},
    {取置ID:'INIT|700|AAA|AAAB',受注番号:'700',商品コード:'AAA',SKU:'AAAb',注文数量:1,追加数量:1,棚確認:'',メモ:''}
  ];
  const plan=context.取り置き_統合反映計画_(inputs,ledger,new Date('2026-07-22'));
  assert.ok(plan.errors.some(e=>/RET1/.test(e)));
  assert.ok(plan.rows.some(r=>r.取置ID==='INIT|700|AAA|AAAB'),'通常行は適用');
  assert.strictEqual(plan.rows.find(r=>r.取置ID==='RET1').戻し処理結果,'未確認','エラーの戻しは未処理のまま');
});

test('統合反映: エラー行の既存開始前在庫は消えない(誤解除防止)', () => {
  const ledger=[{取置ID:'INIT|701|BBB|BBBB',状態:'取り置き中',受注番号:'701',商品コード:'BBB',SKU:'BBBb',取り置き数量:1,取置元種別:'開始前在庫'}];
  const inputs=[{取置ID:'INIT|701|BBB|BBBB',受注番号:'701',商品コード:'BBB',SKU:'BBBb',注文数量:1,追加数量:'abc',棚確認:'',メモ:''}];
  const plan=context.取り置き_統合反映計画_(inputs,ledger,new Date('2026-07-22'));
  assert.ok(plan.errors.length,'数値でない入力はエラー');
  const kept=plan.rows.find(r=>r.取置ID==='INIT|701|BBB|BBBB');
  assert.ok(kept,'既存行は維持');
  assert.strictEqual(kept.取り置き数量,1);
});

test('統合反映: マイナス数量による解除は適用され、空欄は何もしない(2026-07-23契約変更)', () => {
  const ledger=[{取置ID:'INIT|702|DDD|DDDB',状態:'取り置き中',受注番号:'702',商品コード:'DDD',SKU:'DDDb',取り置き数量:1,取置元種別:'開始前在庫'}];
  const blank=context.取り置き_統合反映計画_(
    [{取置ID:'INIT|702|DDD|DDDB',受注番号:'702',商品コード:'DDD',SKU:'DDDb',注文数量:1,追加数量:'',マイナス数量:'',棚確認:'',メモ:''}],
    ledger,new Date('2026-07-23'));
  assert.ok(blank.rows.some(r=>r.取置ID==='INIT|702|DDD|DDDB'),'空欄=何もしない(誤クリック解除の廃止)');
  const minus=context.取り置き_統合反映計画_(
    [{取置ID:'INIT|702|DDD|DDDB',受注番号:'702',商品コード:'DDD',SKU:'DDDb',注文数量:1,追加数量:'',マイナス数量:1,棚確認:'',メモ:''}],
    ledger,new Date('2026-07-23'));
  assert.deepStrictEqual(json(minus.errors),[]);
  assert.ok(!minus.rows.some(r=>r.取置ID==='INIT|702|DDD|DDDB'),'マイナス1で解除');
  assert.strictEqual(minus.counts.解除,1);
  const over=context.取り置き_統合反映計画_(
    [{取置ID:'INIT|702|DDD|DDDB',受注番号:'702',商品コード:'DDD',SKU:'DDDb',注文数量:1,追加数量:'',マイナス数量:2,棚確認:'',メモ:''}],
    ledger,new Date('2026-07-23'));
  assert.ok(over.errors.some(e=>/702/.test(e)),'登録数を超えるマイナスはエラー');
  assert.ok(over.rows.some(r=>r.取置ID==='INIT|702|DDD|DDDB'),'エラー時は行を消さない');
});

test('反映で作る開始前在庫行は最初から現物確認済み段階を持つ(要移行に戻らない)', () => {
  const inputs=[{取置ID:'INIT|703|EEE|EEEB',受注番号:'703',商品コード:'EEE',SKU:'EEEb',注文数量:1,追加数量:1,棚確認:'',メモ:''}];
  const plan=context.取り置き_統合反映計画_(inputs,[],new Date('2026-07-22'));
  const row=plan.rows.find(r=>r.取置ID==='INIT|703|EEE|EEEB');
  assert.strictEqual(row.引当段階,'現物確認済み','棚で数えた現物として登録される');
  assert.strictEqual(row.供給処理,'供給解放','棚の現物は箱供給を消費しない');
  assert.ok(row.現物確認日時);
  assert.strictEqual(context.取り置き_段階正規化_(row,{}).引当段階,'現物確認済み','要移行にならない');
  // 取り置き登録画面の台帳確保(④確保分)としては数えない=自分の登録が超過エラーにならない
  const key=context.取り置き_入力キー_(row);
  assert.strictEqual(context.取り置き_台帳確保集計_(plan.rows)[key]||0,0);
});

test('別ルート(台湾/中国)行も追加数量で棚登録でき、現物確認済み行になる(2026-07-23)', () => {
  const inputs=[{取置ID:'別ルート|900|TW01|TW01',受注番号:'900',商品コード:'TW01',SKU:'TW01',注文数量:2,追加数量:1,棚確認:'',メモ:'棚C'}];
  const plan=context.取り置き_統合反映計画_(inputs,[],new Date('2026-07-23'));
  assert.deepStrictEqual(json(plan.errors),[]);
  const row=plan.rows.find(r=>r.取置ID==='別ルート|900|TW01|TW01');
  assert.ok(row,'別ルートIDのまま台帳へ登録される');
  assert.strictEqual(row.取り置き数量,1);
  assert.strictEqual(row.引当段階,'現物確認済み');
  assert.strictEqual(row.供給処理,'供給解放');
});

test('別ルート二重控除: 台帳で確保した分だけ入荷日ベースの確保を目減りさせる', () => {
  const l={別ルート:true,ban:'900',sku:'TW01',code:'TW01',別ルート済数量:2};
  const key=context.取り置き_行キー_(l);
  const map={}; map[key]=1;
  context.別ルート二重控除_([l],map);
  assert.strictEqual(l.別ルート済数量,1,'台帳1個分を控除(実効確保=大きい方)');
  const l2={別ルート:true,ban:'901',sku:'TW02',code:'TW02',別ルート済数量:2};
  context.別ルート二重控除_([l2],{});
  assert.strictEqual(l2.別ルート済数量,2,'台帳に無ければそのまま');
  const 韓国={別ルート:false,ban:'902',sku:'KR01',code:'KR01',別ルート済数量:0};
  context.別ルート二重控除_([韓国],map);
  assert.strictEqual(韓国.別ルート済数量,0,'韓国ルートは対象外');
});

test('入力保存マージ: 画面に出ていた行の空欄は入力の取り消しとして保存も消す(2026-07-24)', () => {
  const key='700|AAAB';
  const saved=[{入力キー:key,受注番号:'700',SKU:'AAAb',商品コード:'AAA',未反映マイナス数量:2}];
  const 生成=[{受注番号:'700',SKU:'AAAb',商品コード:'AAA',追加数量:'',マイナス数量:''}];
  // 画面でセルを消した状態(行はあるが空欄)で更新
  const r=context.取り置き_入力保存マージ_(生成,saved,[{受注番号:'700',SKU:'AAAb',商品コード:'AAA',追加数量:'',マイナス数量:''}],'now');
  assert.strictEqual(String(r.rows[0].マイナス数量||''),'','手で消した入力は復活しない');
  assert.strictEqual(String(r.store[key].未反映マイナス数量||''),'','保存側からも消える');
});

test('入力保存マージ: 画面に出ていない行(表示モードで非表示)の保存は消さない', () => {
  const key='701|BBBB';
  const saved=[{入力キー:key,受注番号:'701',SKU:'BBBb',商品コード:'BBB',未反映追加数量:3}];
  const r=context.取り置き_入力保存マージ_([{受注番号:'701',SKU:'BBBb',商品コード:'BBB'}],saved,[],'now');
  assert.strictEqual(String(r.rows[0].追加数量||''),'3','非表示中の入力は保持して復元する');
});

if (failures) process.exit(1);
