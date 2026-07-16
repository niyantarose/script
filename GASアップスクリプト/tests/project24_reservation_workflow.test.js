// 取り置き台帳の初回登録ワークフローに関する純粋ロジックのテスト
// 実行: GASアップスクリプト直下で node tests/project24_reservation_workflow.test.js
const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

const activeSpreadsheet = { getSheetByName: () => null };
const context = {
  console,
  Date,
  Set,
  SpreadsheetApp: {
    openById: () => ({ getSheetByName: () => null }),
    getActive: () => activeSpreadsheet
  },
  Utilities: { formatDate: () => '' },
  Logger: { log: () => {} }
};
vm.createContext(context);
[
  'Project_24/引当.js',
  'Project_24/取り置き計算.js',
  'Project_24/取り置き台帳.js',
  'Project_24/消込台帳.js',
  'Project_24/P列自動記入.js'
].filter(fs.existsSync).forEach(f => vm.runInContext(fs.readFileSync(f, 'utf8'), context));

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

test('部分在庫と希望日待ちの受注だけを初期候補にする', () => {
  const orders=[
    {ban:'101',code:'AAA',sku:'AAAb',qty:2,kbn:'取り寄せ'},
    {ban:'102',code:'BBB',sku:'BBBb',qty:1,kbn:'取り寄せ'},
    {ban:'103',code:'CCC',sku:'CCCb',qty:1,kbn:'取り寄せ'}
  ];
  const rows=context.取り置き_初期候補_(orders,new Set(['101']),new Set(['102']));
  assert.strictEqual(rows.length,2);
  assert.strictEqual(rows[0].取置ID,'INIT|101|AAA|AAAB');
  assert.strictEqual(rows[1].取置ID,'INIT|102|BBB|BBBB');
});

test('初回入力は空欄と0を除外し注文数量超過を止める', () => {
  const result=context.取り置き_初期確定計画_([
    {取置ID:'INIT|1',受注番号:'101',商品コード:'AAA',SKU:'AAAb',注文数量:2,現物取り置き数量:2},
    {取置ID:'INIT|2',受注番号:'102',商品コード:'BBB',SKU:'BBBb',注文数量:1,現物取り置き数量:0},
    {取置ID:'INIT|3',受注番号:'103',商品コード:'CCC',SKU:'CCCb',注文数量:1,現物取り置き数量:2}
  ],[],'2026-07-15 10:00:00');
  assert.ok(result.errors.some(e=>/101|102/.test(e))===false);
  assert.ok(result.errors.some(e=>/103/.test(e)));
  assert.strictEqual(result.rows.length,0, '1件でもエラーなら保存対象を返さない');
});

test('同じINITキーは追加せず目標数量へ洗い替える', () => {
  const existing=[{取置ID:'INIT|101|AAA|AAAB',状態:'取り置き中',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:1,取置元種別:'開始前在庫'}];
  const result=context.取り置き_初期確定計画_([
    {取置ID:'INIT|101|AAA|AAAB',受注番号:'101',商品コード:'AAA',SKU:'AAAb',注文数量:2,現物取り置き数量:2}
  ],existing,'2026-07-15 10:00:00');
  assert.strictEqual(JSON.stringify(result.errors),'[]');
  assert.strictEqual(result.rows.length,1);
  assert.strictEqual(result.rows[0].取り置き数量,2);
});

test('初回入力の非数値は既存INIT行を消さず保存全体を止める', () => {
  const existing=[{取置ID:'INIT|101|AAA|AAAB',状態:'取り置き中',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:1,取置元種別:'開始前在庫'}];
  const result=context.取り置き_初期確定計画_([
    {取置ID:'INIT|101|AAA|AAAB',受注番号:'101',商品コード:'AAA',SKU:'AAAb',注文数量:2,現物取り置き数量:'abc'}
  ],existing,'2026-07-15 10:00:00');
  assert.ok(result.errors.some(e=>/初期登録2行|受注101/.test(e)));
  assert.strictEqual(result.rows.length,0, '非数値が1件でもあれば保存対象を返さない');
});

test('処理済だけ発送済み、キャンセルだけキャンセル戻しへ遷移する', () => {
  const ledger=[
    {取置ID:'A',状態:'取り置き中',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:1,戻し処理結果:''},
    {取置ID:'B',状態:'取り置き中',受注番号:'102',商品コード:'BBB',SKU:'BBBb',取り置き数量:1,戻し処理結果:''},
    {取置ID:'C',状態:'取り置き中',受注番号:'103',商品コード:'CCC',SKU:'CCCb',取り置き数量:1,戻し処理結果:''}
  ];
  const plan=context.取り置き_CSV遷移計画_([
    {受注番号:'101',商品コード:'AAA',SKU:'AAAb',受注ステータス:'処理済'},
    {受注番号:'102',商品コード:'BBB',SKU:'BBBb',受注ステータス:'キャンセル'}
  ],ledger,'2026-07-15 10:00:00');
  assert.strictEqual(plan.rows.find(r=>r.取置ID==='A').状態,'発送済み');
  assert.strictEqual(plan.rows.find(r=>r.取置ID==='B').状態,'キャンセル戻し');
  assert.strictEqual(plan.rows.find(r=>r.取置ID==='B').戻し処理結果,'未確認');
  assert.strictEqual(plan.rows.find(r=>r.取置ID==='C').状態,'取り置き中');
  assert.ok(plan.review.some(x=>x.取置ID==='C'));
});

test('列不足または重複する矛盾ステータスでは全遷移を止める', () => {
  const ledger=[{取置ID:'A',状態:'取り置き中',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:1}];
  const plan=context.取り置き_CSV遷移計画_([
    {受注番号:'101',商品コード:'AAA',SKU:'AAAb',受注ステータス:'処理済'},
    {受注番号:'101',商品コード:'AAA',SKU:'AAAb',受注ステータス:'キャンセル'}
  ],ledger,'2026-07-15 10:00:00');
  assert.ok(plan.errors.some(e=>/競合/.test(e)));
  assert.strictEqual(plan.rows.length,0);
});

test('CSVキャンセル後処理は同じ受注番号の継続商品を巻き込まない', () => {
  const ledger=[
    {取置ID:'A',状態:'取り置き中',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:1},
    {取置ID:'B',状態:'取り置き中',受注番号:'101',商品コード:'BBB',SKU:'BBBb',取り置き数量:1}
  ];
  const plan=context.取り置き_CSV遷移計画_([
    {受注番号:'101',商品コード:'AAA',SKU:'AAAb',受注ステータス:'キャンセル'},
    {受注番号:'101',商品コード:'BBB',SKU:'BBBb',受注ステータス:'新規受付'}
  ],ledger,'2026-07-15 10:00:00');
  let saved=plan.rows, reads=0, writes=0;
  context.個別対応_EMSリスト_=()=>({error:'test'});
  context.取り置き台帳_読む_=()=>{ reads++; return saved; };
  context.取り置き台帳_保存_=rows=>{ writes++; saved=rows; };

  context.キャンセル処理_(['101'],{取り置き台帳を更新:false});

  assert.strictEqual(saved.find(r=>r.取置ID==='A').状態,'キャンセル戻し');
  assert.strictEqual(saved.find(r=>r.取置ID==='B').状態,'取り置き中');
  assert.strictEqual(reads,0,'CSV後処理は取り置き台帳を再読込しない');
  assert.strictEqual(writes,0,'CSV後処理は取り置き台帳を再保存しない');
});

test('手動キャンセルは同じ受注番号の取り置き中を全行戻す', () => {
  let ledger=[
    {取置ID:'A',状態:'取り置き中',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:1},
    {取置ID:'B',状態:'取り置き中',受注番号:'101',商品コード:'BBB',SKU:'BBBb',取り置き数量:1}
  ];
  let writes=0;
  context.個別対応_EMSリスト_=()=>({error:'test'});
  context.取り置き台帳_読む_=()=>ledger;
  context.取り置き台帳_保存_=rows=>{ writes++; ledger=rows; };

  context.キャンセル処理_(['101']);

  assert.strictEqual(ledger.filter(r=>r.状態==='キャンセル戻し').length,2);
  assert.ok(ledger.every(r=>r.戻し処理結果==='未確認'));
  assert.strictEqual(writes,1);
});

test('全ステータスCSVの必須見出し不足はadapterで拒否する', () => {
  assert.throws(
    ()=>context.CSV行を受注行オブジェクトへ_(['受注番号','商品コード','商品SKU','個数'],[['101','AAA','AAAb','1']]),
    /全ステータスCSVの見出し不足: 受注ステータス/
  );
});

test('現物ありと在庫なしだけを戻し結果へ反映する', () => {
  const ledger=[
    {取置ID:'A',状態:'キャンセル戻し',受注番号:'101',商品コード:'AAA',取り置き数量:1,戻し処理結果:'未確認'},
    {取置ID:'B',状態:'キャンセル戻し',受注番号:'102',商品コード:'BBB',取り置き数量:1,戻し処理結果:'未確認'}
  ];
  const result=context.取り置き_戻し確認計画_([{取置ID:'A',現物確認:'現物あり'},{取置ID:'B',現物確認:'在庫なし'}],ledger,'2026-07-15 10:00:00');
  assert.strictEqual(JSON.stringify(result.errors),'[]');
  assert.strictEqual(result.rows.find(r=>r.取置ID==='A').戻し処理結果,'現物あり');
  assert.strictEqual(result.rows.find(r=>r.取置ID==='B').戻し処理結果,'在庫なし');
});

test('未確認のまま、または不明な選択肢では確定しない', () => {
  const ledger=[{取置ID:'A',状態:'キャンセル戻し',受注番号:'101',商品コード:'AAA',取り置き数量:1,戻し処理結果:'未確認'}];
  const result=context.取り置き_戻し確認計画_([{取置ID:'A',現物確認:'たぶんある'}],ledger,'2026-07-15 10:00:00');
  assert.ok(result.errors.length>0);
  assert.strictEqual(result.rows.length,0);
});

test('台帳I/Oと初回登録の公開入口を提供する', () => {
  [
    '取り置き台帳_読む_',
    '取り置き台帳_保存_',
    'EMS在庫移動台帳_読む_',
    'EMS在庫移動台帳_保存_',
    '取り置き初期登録を作成',
    '取り置き初期登録を確定',
    'キャンセル戻し確認を更新',
    'キャンセル戻し確認を確定',
    'Yahoo戻し候補を更新_',
    '選択した取り置きを手動解除'
  ].forEach(name=>assert.strictEqual(typeof context[name],'function',name));
});

test('戻し確認一覧は未確認行だけを保存し現物確認列へ一括入力規則を設定する', () => {
  let saved, rangeCalls=0, validationCalls=0;
  context.取り置き台帳_読む_=()=>[
    {取置ID:'A',状態:'キャンセル戻し',受注番号:'101',商品コード:'AAA',取り置き数量:1,元EMS番号:'EG1',戻し処理結果:'未確認','終了理由・メモ':'確認待ち'},
    {取置ID:'B',状態:'キャンセル戻し',受注番号:'102',商品コード:'BBB',取り置き数量:1,元EMS番号:'EG2',戻し処理結果:'現物あり'}
  ];
  context.取り置き_表を保存_=(sheet,headers,rows)=>{ saved={sheet,headers,rows}; };
  context.SpreadsheetApp.getActive=()=>({
    getSheetByName:()=>({getRange:()=>{ rangeCalls++; return {setDataValidation:()=>{ validationCalls++; }}; }}),
    toast:()=>{}
  });
  context.SpreadsheetApp.newDataValidation=()=>({
    requireValueInList(){ return this; },
    setAllowInvalid(){ return this; },
    build(){ return {}; }
  });

  context.キャンセル戻し確認を更新();

  assert.strictEqual(saved.sheet,'キャンセル戻し確認');
  assert.strictEqual(JSON.stringify(saved.headers),JSON.stringify(['取置ID','受注番号','商品コード','数量','元EMS番号','現物確認','メモ']));
  assert.strictEqual(saved.rows.length,1);
  assert.strictEqual(saved.rows[0].取置ID,'A');
  assert.strictEqual(rangeCalls,1);
  assert.strictEqual(validationCalls,1);
});

test('Yahoo戻し候補は現物ありだけを保存し確認列へ一括チェックボックスを設定する', () => {
  let saved, rangeCalls=0, checkboxCalls=0;
  context.取り置き台帳_読む_=()=>[
    {取置ID:'A',状態:'キャンセル戻し',商品コード:'AAA',取り置き数量:1,元EMS番号:'EG1',戻し処理結果:'現物あり'},
    {取置ID:'B',状態:'キャンセル戻し',商品コード:'BBB',取り置き数量:1,元EMS番号:'EG2',戻し処理結果:'在庫なし'}
  ];
  context.取り置き_表を保存_=(sheet,headers,rows)=>{ saved={sheet,headers,rows}; };
  context.SpreadsheetApp.getActive=()=>({
    getSheetByName:()=>({getRange:()=>{ rangeCalls++; return {insertCheckboxes:()=>{ checkboxCalls++; }}; }})
  });

  context.Yahoo戻し候補を更新_();

  assert.strictEqual(saved.sheet,'Yahoo戻し候補');
  assert.strictEqual(JSON.stringify(saved.headers),JSON.stringify(['取置ID','商品コード','数量','元EMS番号','処理ID','確認']));
  assert.strictEqual(saved.rows.length,1);
  assert.strictEqual(saved.rows[0].処理ID,'YAHOO|RETURN|A');
  assert.strictEqual(rangeCalls,1);
  assert.strictEqual(checkboxCalls,1);
});

test('選択した取り置き中の行は理由付きで手動解除する', () => {
  let saved;
  context.取り置き台帳_読む_=()=>[{取置ID:'A',状態:'取り置き中','終了理由・メモ':''}];
  context.取り置き台帳_保存_=rows=>{ saved=rows; };
  const ui={
    Button:{OK:'OK'}, ButtonSet:{OK_CANCEL:'OK_CANCEL'}, alert:()=>{},
    prompt:()=>({getSelectedButton:()=> 'OK',getResponseText:()=> '登録間違い'})
  };
  context.SpreadsheetApp.getUi=()=>ui;
  context.SpreadsheetApp.getActive=()=>({
    getActiveSheet:()=>({getName:()=> '取り置き台帳',getActiveRange:()=>({getRow:()=>2})})
  });

  context.選択した取り置きを手動解除();

  assert.strictEqual(saved[0].状態,'手動解除');
  assert.strictEqual(saved[0]['終了理由・メモ'],'登録間違い');
});

test('手動解除は空の理由で台帳を保存しない', () => {
  let writes=0, alerts=[];
  context.取り置き台帳_読む_=()=>[{取置ID:'A',状態:'取り置き中','終了理由・メモ':''}];
  context.取り置き台帳_保存_=()=>{ writes++; };
  const ui={
    Button:{OK:'OK'}, ButtonSet:{OK_CANCEL:'OK_CANCEL'}, alert:message=>{ alerts.push(message); },
    prompt:()=>({getSelectedButton:()=> 'OK',getResponseText:()=> '   '})
  };
  context.SpreadsheetApp.getUi=()=>ui;
  context.SpreadsheetApp.getActive=()=>({
    getActiveSheet:()=>({getName:()=> '取り置き台帳',getActiveRange:()=>({getRow:()=>2})})
  });

  context.選択した取り置きを手動解除();

  assert.strictEqual(writes,0);
  assert.ok(alerts.includes('解除理由は必須です'));
});

test('P列計画は同じEMSの取り置き中を固定し、残数だけ新規FIFOへ回す', () => {
  const rows=[{ems:'EG1',code:'AAA',qty:3,pOriginal:'101',arrival:'2026-07-12'}];
  const orders=[
    {ban:'101',code:'AAA',sku:'AAAb',qty:1,need:0,date:new Date('2026-07-01'),keys:['AAA']},
    {ban:'102',code:'AAA',sku:'AAAb',qty:2,need:2,date:new Date('2026-07-02'),keys:['AAA']}
  ];
  const fixed={'EG1|AAA':[{ban:'101',qty:1}]};
  const usage={'EG1|AAA':{取り置き中:1,発送済み:0,戻し未処理:0,在庫なし確定:0,Yahoo移動済み:0}};
  const result=context.P列計画_純計算_(rows,orders,fixed,usage);
  assert.strictEqual(result.rows[0].nextP,'101:1, 102:2');
  assert.strictEqual(result.rows[0].left,0);
});

test('発送済みと戻し未処理は供給を塞ぐがP列の現役注文表示には出さない', () => {
  const rows=[{ems:'EG1',code:'AAA',qty:3,pOriginal:'',arrival:'2026-07-12'}];
  const orders=[{ban:'102',code:'AAA',sku:'AAAb',qty:2,need:2,date:new Date('2026-07-02'),keys:['AAA']}];
  const usage={'EG1|AAA':{取り置き中:0,発送済み:1,戻し未処理:1,在庫なし確定:0,Yahoo移動済み:0}};
  const result=context.P列計画_純計算_(rows,orders,{},usage);
  assert.strictEqual(result.rows[0].nextP,'102:1');
  assert.strictEqual(result.rows[0].entries[0].qty,1);
  assert.strictEqual(orders[0].need,2,'入力注文を変更しない');
});

test('同一供給キーの複数EMS行で固定と使用数を二重計上しない', () => {
  const rows=[
    {ems:'EG1',code:'AAA',qty:1,pOriginal:'',arrival:'2026-07-12'},
    {ems:'EG1',code:'AAA',qty:2,pOriginal:'',arrival:'2026-07-12'}
  ];
  const orders=[{ban:'102',code:'AAA',sku:'AAAb',qty:1,need:1,date:new Date('2026-07-02'),keys:['AAA']}];
  const fixed={'EG1|AAA':[{ban:'101',qty:1}]};
  const usage={'EG1|AAA':{取り置き中:1,発送済み:1,戻し未処理:0,在庫なし確定:0,Yahoo移動済み:0}};
  const result=context.P列計画_純計算_(rows,orders,fixed,usage);
  assert.strictEqual(result.rows[0].nextP,'101');
  assert.strictEqual(result.rows[1].nextP,'102:1');
  assert.strictEqual(result.rows[1].left,0);
});

test('P列direct行は指定ownerだけを表示し一般FIFOへ回さない', () => {
  const rows=[
    {ems:'EG1',code:'POEM65',sourceCode:'POEM65（102）',directBan:'102',qty:1,pOriginal:'',arrival:'2026-07-12'},
    {ems:'EG1',code:'POEM65',sourceCode:'POEM65（101）',directBan:'101',qty:1,pOriginal:'',arrival:'2026-07-12'}
  ];
  const orders=[
    {ban:'101',code:'POEM65',sku:'POEM65b',qty:1,need:1,date:new Date('2026-07-01'),keys:['POEM65']},
    {ban:'102',code:'POEM65',sku:'POEM65b',qty:1,need:1,date:new Date('2026-07-02'),keys:['POEM65']},
    {ban:'103',code:'POEM65',sku:'POEM65b',qty:1,need:1,date:new Date('2026-07-03'),keys:['POEM65']}
  ];
  const result=context.P列計画_純計算_(rows,orders,{},{});
  assert.strictEqual(result.rows[0].nextP,'102');
  assert.strictEqual(result.rows[1].nextP,'101');
  assert.strictEqual(result.orders.find(o=>o.ban==='103').need,1);
});

test('P列direct行は同一owner内でも基底コード一致の商品需要だけを減らす', () => {
  const result=context.P列計画_純計算_([
    {ems:'EG1',code:'POEM65',sourceCode:'POEM65（10116569）',directBan:'10116569',qty:1,pOriginal:'',arrival:'2026-07-12'}
  ],[
    {ban:'10116569',code:'RECIPE42',sku:'RECIPE42b',qty:1,need:1,date:new Date('2026-07-01'),keys:['RECIPE42']},
    {ban:'10116569',code:'POEM65',sku:'POEM65b',qty:1,need:1,date:new Date('2026-07-01'),keys:['POEM65']}
  ],{},{});
  assert.strictEqual(result.orders[0].need,1);
  assert.strictEqual(result.orders[1].need,0);
});

test('P列direct表示は新規explicit割当へ戻さない', () => {
  const plan={rows:[
    {ems:'EG1',code:'POEM65',sourceCode:'POEM65（101）',directBan:'101',entries:[{ban:'101',qty:1}]},
    {ems:'EG1',code:'AAA',sourceCode:'AAA',directBan:'',entries:[{ban:'102',qty:1}]}
  ]};
  assert.strictEqual(JSON.stringify(context.P列計画_確定割当_(plan)),JSON.stringify([
    {ems:'EG1',code:'AAA',sourceCode:'AAA',ban:'102',qty:1}
  ]));
  assert.strictEqual(JSON.stringify(context.P列計画_新規確定割当_(plan,[])),JSON.stringify([
    {ems:'EG1',code:'AAA',sourceCode:'AAA',ban:'102',qty:1}
  ]));
});

test('便の引き直し日付は暦日を検証しうるう日だけ許可する', () => {
  assert.strictEqual(JSON.stringify(context.便の引当をやり直す_日付解析_('2024-02-29')),JSON.stringify({dates:['2024-02-29']}));
  assert.match(context.便の引当をやり直す_日付解析_('2026-02-29').error,/日付/);
  assert.match(context.便の引当をやり直す_日付解析_('2026-02-31').error,/日付/);
});

test('P列計画を作っただけではsetValuesを呼ばない', () => {
  let writes=0;
  const plan={sheet:{getRange:()=>({setValues:()=>{writes++;},setBackground:()=>{}})},startRow:7,colP:16,rowCount:1,values:[['101']],backgrounds:[],writes:[0],summary:{}};
  assert.strictEqual(writes,0);
  context.発注共有P列計画を反映_(plan);
  assert.strictEqual(writes,1);
});

function P列シートモック_(pValue){
  const recvHead=['受注番号','注文日時','商品コード','商品SKU','項目・選択肢','個数','入荷日'];
  const recvRows=[recvHead,['102',new Date('2026-07-02'),'AAA','AAAb','取り寄せ',2,'']];
  const recv={
    getLastRow:()=>recvRows.length,
    getLastColumn:()=>recvHead.length,
    getDataRange:()=>({getValues:()=>recvRows}),
    getRange:(row,col,numRows,numCols)=>({getValues:()=>recvRows.slice(row-1,row-1+numRows).map(r=>r.slice(col-1,col-1+numCols))})
  };
  const emsHead=['ステータス列','購入No.','商品コード','数量','注文番号','EMS到着日','EMS番号'];
  const emsRows=[['到着済','20260701_01_01','AAA',3,pValue,'2026-07-12','EG1']];
  const state={setValues:0,setBackground:0,values:null};
  const ems={
    getLastRow:()=>7,
    getLastColumn:()=>emsHead.length,
    getSheetId:()=>123,
    getParent:()=>({getId:()=> 'BOOK1'}),
    getRange:(row,col,numRows,numCols)=>({
      getValues:()=>row===6?[emsHead.slice(col-1,col-1+numCols)]:emsRows.slice(row-7,row-7+numRows).map(r=>r.slice(col-1,col-1+numCols)),
      getDisplayValues:()=>row===6?[emsHead.slice(col-1,col-1+numCols)]:emsRows.slice(row-7,row-7+numRows).map(r=>r.slice(col-1,col-1+numCols)),
      setValues:values=>{ state.setValues++; state.values=values; },
      setBackground:()=>{ state.setBackground++; }
    })
  };
  return {recv,ems,state};
}

test('シートからP列計画を作る間は書込みミューテータを呼ばない', () => {
  const mock=P列シートモック_('OLD');
  context.P_KAKUTEI_CFG={シート:'EMSリスト',ヘッダー行:6};
  context.SpreadsheetApp.getActive=()=>({getSheetByName:()=>mock.recv});
  context.発注共有を開く_=()=>({getSheetByName:()=>mock.ems});
  context.取り置き台帳_読む_=()=>[{
    取置ID:'A',状態:'取り置き中',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:1,
    取置元種別:'EMS',元EMS番号:'EG1',元EMS商品コード:'AAA'
  }];
  context.EMS在庫移動台帳_読む_=()=>[];

  const plan=context.発注共有P列計画_({currentP:[['']]});

  assert.strictEqual(plan.error,'');
  assert.strictEqual(plan.values[0][0],'101:1, 102:2');
  assert.strictEqual(mock.state.setValues,0);
  assert.strictEqual(mock.state.setBackground,0);
});

test('P列書き直しはメモリ上でクリアした計画をP列へ1回だけ書く', () => {
  const mock=P列シートモック_('OLD');
  context.P_KAKUTEI_CFG={シート:'EMSリスト',ヘッダー行:6};
  context.SpreadsheetApp.getActive=()=>({getSheetByName:()=>mock.recv});
  context.発注共有を開く_=()=>({getSheetByName:()=>mock.ems});
  context.取り置き台帳_読む_=()=>[];
  context.EMS在庫移動台帳_読む_=()=>[];

  const result=context.P列書き直し実行_();

  assert.strictEqual(result.error,undefined);
  assert.strictEqual(mock.state.setValues,1);
  assert.strictEqual(mock.state.values[0][0],'102:2');
});

test('P列書き直しは別wrapperでも同じspreadsheet・sheet IDなら計画を反映する', () => {
  const mock=P列シートモック_('OLD');
  const wrap=sheetId=>Object.assign({},mock.ems,{
    getSheetId:()=>sheetId,
    getParent:()=>({getId:()=> 'BOOK1'})
  });
  const sheets=[wrap(123),wrap(123)]; let opened=0;
  context.P_KAKUTEI_CFG={シート:'EMSリスト',ヘッダー行:6};
  context.SpreadsheetApp.getActive=()=>({getSheetByName:()=>mock.recv});
  context.発注共有を開く_=()=>({getSheetByName:()=>sheets[opened++]});
  context.取り置き台帳_読む_=()=>[];
  context.EMS在庫移動台帳_読む_=()=>[];

  const result=context.P列書き直し実行_();

  assert.strictEqual(result.error,undefined);
  assert.strictEqual(mock.state.setValues,1);
});

test('P列書き直しは別sheet IDの計画を書かずに拒否する', () => {
  const mock=P列シートモック_('OLD');
  const wrap=sheetId=>Object.assign({},mock.ems,{
    getSheetId:()=>sheetId,
    getParent:()=>({getId:()=> 'BOOK1'})
  });
  const sheets=[wrap(123),wrap(456)]; let opened=0;
  context.P_KAKUTEI_CFG={シート:'EMSリスト',ヘッダー行:6};
  context.SpreadsheetApp.getActive=()=>({getSheetByName:()=>mock.recv});
  context.発注共有を開く_=()=>({getSheetByName:()=>sheets[opened++]});
  context.取り置き台帳_読む_=()=>[];
  context.EMS在庫移動台帳_読む_=()=>[];

  const result=context.P列書き直し実行_();

  assert.match(result.error,/範囲がEMSリストと一致しません/);
  assert.strictEqual(mock.state.setValues,0);
});

test('P列計画の表示エントリを確定割当へ展開する', () => {
  const result=context.P列計画_確定割当_({rows:[
    {ems:'EG1',code:'AAA',entries:[{ban:'101',qty:1},{ban:'102',qty:2}]},
    {ems:'EG2',code:'BBB',entries:[]}
  ]});
  assert.strictEqual(JSON.stringify(result),JSON.stringify([
    {ems:'EG1',code:'AAA',sourceCode:'AAA',ban:'101',qty:1},
    {ems:'EG1',code:'AAA',sourceCode:'AAA',ban:'102',qty:2}
  ]));
});

test('受注共通メニューに初回登録の作成と確定を追加する', () => {
  const source=fs.readFileSync('Project_24/引当.js','utf8');
  assert.ok(source.includes(".addItem('📋 取り置き初期登録を作成', '取り置き初期登録を作成')"));
  assert.ok(source.includes(".addItem('✅ 取り置き初期登録を確定', '取り置き初期登録を確定')"));
  assert.ok(source.includes(".addItem('📦 キャンセル戻し確認を更新', 'キャンセル戻し確認を更新')"));
  assert.ok(source.includes(".addItem('✅ キャンセル戻し確認を確定', 'キャンセル戻し確認を確定')"));
  assert.ok(source.includes(".addItem('🔓 選択した取り置きを手動解除', '選択した取り置きを手動解除')"));
});

process.exitCode = failures ? 1 : 0;
console.log(failures ? `FAILURES: ${failures}` : 'ALL TESTS PASSED');
