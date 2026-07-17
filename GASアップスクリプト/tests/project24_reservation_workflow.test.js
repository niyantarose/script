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
  LockService: { getDocumentLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) }, // 直列_用
  Utilities: { formatDate: () => '' },
  Logger: { log: () => {} }
};
vm.createContext(context);
[
  'Project_24/引当.js',
  'Project_24/引当履歴.js',
  'Project_24/取り置き計算.js',
  'Project_24/取り置き台帳.js',
  'Project_24/消込台帳.js',
  'Project_24/P列自動記入.js'
].filter(fs.existsSync).forEach(f => vm.runInContext(fs.readFileSync(f, 'utf8'), context));
const 取り置き表保存実装_=context.取り置き_表を保存_;

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

test('一覧に載る受注だけを初期候補にし、現在の状態ラベルを付ける', () => {
  const orders=[
    {ban:'101',code:'AAA',sku:'AAAb',qty:2,kbn:'取り寄せ'},
    {ban:'102',code:'BBB',sku:'BBBb',qty:1,kbn:'取り寄せ'},
    {ban:'103',code:'CCC',sku:'CCCb',qty:1,kbn:'取り寄せ'}
  ];
  const rows=context.取り置き_初期候補_(orders,[
    {状態:'部分在庫',bans:new Set(['101'])},
    {状態:'希望日待ち',bans:new Set(['102'])}
  ]);
  assert.strictEqual(rows.length,2);
  assert.strictEqual(rows[0].取置ID,'INIT|101|AAA|AAAB');
  assert.strictEqual(rows[0].現在の状態,'部分在庫');
  assert.strictEqual(rows[1].取置ID,'INIT|102|BBB|BBBB');
  assert.strictEqual(rows[1].現在の状態,'希望日待ち');
});

test('初期候補は氏名を受注番号の隣に載せる', () => {
  const rows=context.取り置き_初期候補_(
    [{ban:'101',氏名:'西野 瑠璃',code:'AAA',sku:'AAAb',qty:1,kbn:'取り寄せ'}],
    [{状態:'部分在庫',bans:new Set(['101'])}]);
  assert.strictEqual(rows[0].氏名,'西野 瑠璃');
});

test('どの一覧にも無い着済スタンプ注文も候補に出る(引当待ちに埋もれた着済の取りこぼし防止)', () => {
  const orders=[
    {ban:'10116494',氏名:'金木 あずみ',code:'JMEE-SPKZ-06',sku:'JMEE-SPKZ-06b',qty:2,kbn:'取り寄せ',入荷日:'2026-06-25',EMS:'EG049108127KR'},
    {ban:'10119999',code:'ZZZ',sku:'ZZZb',qty:1,kbn:'取り寄せ'} // スタンプなし・一覧にも無し=候補外
  ];
  const rows=context.取り置き_初期候補_(orders,[
    {状態:'部分在庫',bans:new Set([])},
    {状態:'着済スタンプ(要棚確認)',bans:new Set(['10116494'])}
  ]);
  assert.strictEqual(rows.length,1);
  assert.strictEqual(rows[0].受注番号,'10116494');
  assert.strictEqual(rows[0].現在の状態,'着済スタンプ(要棚確認)');
  assert.strictEqual(rows[0].旧入荷日,'2026-06-25');
});

test('初期候補は旧帳簿の入荷日とEMS番号を目印として載せる', () => {
  const rows=context.取り置き_初期候補_(
    [{ban:'101',氏名:'中村 唯',code:'MRBLUE40',sku:'MRBLUE40-7b',qty:2,kbn:'取り寄せ',入荷日:'2026-06-25',EMS:'EG049618607KR'}],
    [{状態:'希望日待ち',bans:new Set(['101'])}]);
  assert.strictEqual(rows[0].旧入荷日,'2026-06-25');
  assert.strictEqual(rows[0].旧EMS,'EG049618607KR');
});

test('初期候補の再作成は確定済みの開始前在庫数量を入力済みで表示する', () => {
  const candidates=context.取り置き_初期候補_(
    [{ban:'101',code:'AAA',sku:'AAAb',qty:2,kbn:'取り寄せ'},{ban:'102',code:'BBB',sku:'BBBb',qty:1,kbn:'取り寄せ'}],
    [{状態:'部分在庫',bans:new Set(['101','102'])}]);
  const ledger=[{取置ID:candidates[0].取置ID,状態:'取り置き中',取置元種別:'開始前在庫',取り置き数量:2}];
  const rows=context.取り置き_初期候補へ既存数量_(candidates,ledger);
  assert.strictEqual(rows[0].現物取り置き数量,2,'確定済みは数量入り');
  assert.strictEqual(rows[1].現物取り置き数量,'','未確定は空欄のまま');
});

test('登録シートの洗い替えは手入力の数量・棚確認・メモを引き継ぐ', () => {
  const candidates=context.取り置き_初期候補_(
    [{ban:'101',code:'AAA',sku:'AAAb',qty:2,kbn:'取り寄せ'},
     {ban:'102',code:'BBB',sku:'BBBb',qty:1,kbn:'取り寄せ'},
     {ban:'103',code:'CCC',sku:'CCCb',qty:1,kbn:'取り寄せ'}],
    [{状態:'部分在庫',bans:new Set(['101','102','103'])}]);
  const ledger=[{取置ID:candidates[0].取置ID,状態:'取り置き中',取置元種別:'開始前在庫',取り置き数量:2}];
  const sheetRows=[
    {取置ID:candidates[1].取置ID,現物取り置き数量:1,メモ:'棚の右奥'},   // 未確定の手入力
    {取置ID:candidates[2].取置ID,現物取り置き数量:'',メモ:'出荷済み'}   // 旧メモの分類語
  ];
  const rows=context.取り置き_登録シート引き継ぎ_(candidates,sheetRows,ledger);
  assert.strictEqual(rows[0].現物取り置き数量,2,'台帳確定分');
  assert.strictEqual(rows[1].現物取り置き数量,1,'未確定の手入力も残る');
  assert.strictEqual(rows[1].メモ,'棚の右奥');
  assert.strictEqual(rows[2].現物取り置き数量,'','空欄は空欄のまま');
  assert.strictEqual(rows[2].棚確認,'出荷済み','旧メモの分類語はプルダウン列へ移す');
  assert.strictEqual(rows[2].メモ,'','移した分類語はメモから消える');
});

test('棚確認が出荷済み/未着/予約なのに数量入りは確定を止める', () => {
  const result=context.取り置き_初期確定計画_([
    {取置ID:'INIT|1',受注番号:'101',商品コード:'AAA',SKU:'AAAb',注文数量:2,棚確認:'出荷済み',現物取り置き数量:1},
    {取置ID:'INIT|2',受注番号:'102',商品コード:'BBB',SKU:'BBBb',注文数量:1,棚確認:'発送待ち',現物取り置き数量:1},
    {取置ID:'INIT|3',受注番号:'103',商品コード:'CCC',SKU:'CCCb',注文数量:1,棚確認:'予約',現物取り置き数量:1}
  ],[],'2026-07-16 18:00:00');
  assert.ok(result.errors.some(e=>/101.*出荷済み/.test(e)),JSON.stringify(result.errors));
  assert.ok(!result.errors.some(e=>/102/.test(e)),'発送待ち+数量は正常');
  assert.ok(result.errors.some(e=>/103.*予約/.test(e)),'予約+数量は矛盾');
});

test('商品名や選択肢の予約表記では棚確認を自動予約にしない', () => {
  const rows=context.取り置き_初期候補_([
    {ban:'201',code:'AAA',sku:'AAAb',qty:1,予約:true,商品名:'（予約5月発売）商品'},
    {ban:'202',code:'BBB',sku:'BBBb',qty:1,選択肢:'予約9月発売予定'}
  ],[{状態:'着済スタンプ(要棚確認)',bans:new Set(['201','202'])}]);
  assert.strictEqual(JSON.stringify(rows.map(r=>r.棚確認)),JSON.stringify(['','']));
});

test('同じ受注×商品の分割行は注文数量を合算して1候補にする', () => {
  const orders=[
    {ban:'201',code:'ANKI22b',sku:'ANKI22b',qty:2,kbn:'取り寄せ'},
    {ban:'201',code:'ANKI22b',sku:'ANKI22b',qty:1,kbn:'取り寄せ'}
  ];
  const rows=context.取り置き_初期候補_(orders,[{状態:'部分在庫',bans:new Set(['201'])}]);
  assert.strictEqual(rows.length,1);
  assert.strictEqual(rows[0].注文数量,3);
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

test('箱余りのYahoo移動は同じEMS・商品で一度だけ記録する', () => {
  const surplus=[{ems:'EG1',code:'AAA',sourceCode:'AAA',directBan:'',qty:2,arrival:'2026-07-12'}];
  const first=context.EMS在庫移動_箱計画_(surplus,[],'2026-07-15 10:00:00');
  const second=context.EMS在庫移動_箱計画_(surplus,first.rows,'2026-07-15 10:01:00');
  assert.strictEqual(JSON.stringify(first.errors),'[]');
  assert.strictEqual(first.added.length,1);
  assert.strictEqual(first.added[0].処理ID,'YAHOO|EMS|EG1|AAA');
  assert.strictEqual(second.added.length,0);
  assert.strictEqual(second.rows.length,1);
});

test('箱余りのYahoo移動は実surplusのsource identityを商品コードと処理IDへ保持する', () => {
  const surplus=[
    {ems:'EG1',code:'POEM65（10116569）',matchCode:'POEM65',sourceCode:'PoEm65（10116569）',directBan:'10116569',qty:1,arrival:'2026-07-12'},
    {ems:'EG1',code:'RECIPE42/10117126',matchCode:'RECIPE42',sourceCode:'RECIPE42/10117126',directBan:'10117126',qty:1,arrival:'2026-07-12'},
    {ems:'EG1',code:'10117375',matchCode:'10117375',sourceCode:'10117375',directBan:'10117375',qty:1,arrival:'2026-07-12'}
  ];
  const plan=context.EMS在庫移動_箱計画_(surplus,[],'2026-07-15 10:00:00');
  assert.strictEqual(JSON.stringify(plan.errors),'[]');
  assert.strictEqual(JSON.stringify(plan.added.map(r=>[r.処理ID,r.商品コード])),JSON.stringify([
    ['YAHOO|EMS|EG1|POEM65（10116569）','PoEm65（10116569）'],
    ['YAHOO|EMS|EG1|RECIPE42/10117126','RECIPE42/10117126'],
    ['YAHOO|EMS|EG1|10117375','10117375']
  ]));
});

test('実EMS供給parserはdirect ownerをraw sourceCodeへ必ず保持する', () => {
  const rows=[
    ['EG1','PoEm65（10116569）',1],
    ['EG1','RECIPE42/10117126',1],
    ['EG1','10117375',1],
    ['EG1','AAA',1]
  ];
  const supplies=context.EMS供給オブジェクト_(rows,{EMS番号:0,コード:1,数量:2},()=> '2026-07-12');
  assert.strictEqual(JSON.stringify(supplies.map(s=>[s.sourceCode,s.directBan])),JSON.stringify([
    ['PoEm65（10116569）','10116569'],
    ['RECIPE42/10117126','10117126'],
    ['10117375','10117375'],
    ['AAA','']
  ]));
});

test('キャンセル戻しは取置IDごとに一度だけYahoo移動しraw元EMS商品コードを使う', () => {
  const returns=[{取置ID:'OLD1',状態:'キャンセル戻し',戻し処理結果:'現物あり',商品コード:'POEM65',取り置き数量:1,元EMS番号:'EG0',元EMS商品コード:'PoEm65（10116569）'}];
  const first=context.EMS在庫移動_戻し計画_(returns,[],'2026-07-15 10:00:00');
  const second=context.EMS在庫移動_戻し計画_(returns,first.rows,'2026-07-15 10:01:00');
  assert.strictEqual(first.added[0].処理ID,'YAHOO|RETURN|OLD1');
  assert.strictEqual(first.added[0].商品コード,'PoEm65（10116569）');
  assert.strictEqual(second.added.length,0);
});

test('台帳I/Oと初回登録の公開入口を提供する', () => {
  [
    '取り置き台帳_読む_',
    '取り置き台帳_保存_',
    'EMS在庫移動台帳_読む_',
    'EMS在庫移動台帳_保存_',
    'キャンセル戻しをYahoo反映済みにする',
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
    {取置ID:'A',状態:'キャンセル戻し',商品コード:'POEM65',元EMS商品コード:'PoEm65（10116569）',取り置き数量:1,元EMS番号:'EG1',戻し処理結果:'現物あり'},
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
  assert.strictEqual(saved.rows[0].商品コード,'PoEm65（10116569）');
  assert.strictEqual(rangeCalls,1);
  assert.strictEqual(checkboxCalls,1);
});

// 手動解除は行番号ではなく選択行の取置IDで対象を特定する(空行・行操作のズレで別の行を解除しない)
function 手動解除シートモック_(selectedRowId){
  return {
    getName:()=> '取り置き台帳',
    getActiveRange:()=>({getRow:()=>2}),
    getRange:()=>({getDisplayValue:()=>selectedRowId})
  };
}

test('選択した取り置き中の行は理由付きで手動解除する', () => {
  let saved;
  context.取り置き台帳_読む_=()=>[
    {取置ID:'A',状態:'取り置き中','終了理由・メモ':''},
    {取置ID:'B',状態:'取り置き中','終了理由・メモ':''}
  ];
  context.取り置き台帳_保存_=rows=>{ saved=rows; };
  const ui={
    Button:{OK:'OK'}, ButtonSet:{OK_CANCEL:'OK_CANCEL'}, alert:()=>{},
    prompt:()=>({getSelectedButton:()=> 'OK',getResponseText:()=> '登録間違い'})
  };
  context.SpreadsheetApp.getUi=()=>ui;
  context.SpreadsheetApp.getActive=()=>({ getActiveSheet:()=>手動解除シートモック_('A') });

  context.選択した取り置きを手動解除();

  assert.strictEqual(saved[0].状態,'手動解除');
  assert.strictEqual(saved[0]['終了理由・メモ'],'登録間違い');
  assert.strictEqual(saved[1].状態,'取り置き中','選択していない行は触らない');
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
  context.SpreadsheetApp.getActive=()=>({ getActiveSheet:()=>手動解除シートモック_('A') });

  context.選択した取り置きを手動解除();

  assert.strictEqual(writes,0);
  assert.ok(alerts.includes('解除理由は必須です'));
});

test('手動解除は取置IDが台帳と一致しないとき何も書かない', () => {
  let writes=0, alerts=[];
  context.取り置き台帳_読む_=()=>[{取置ID:'A',状態:'取り置き中','終了理由・メモ':''}];
  context.取り置き台帳_保存_=()=>{ writes++; };
  const ui={
    Button:{OK:'OK'}, ButtonSet:{OK_CANCEL:'OK_CANCEL'}, alert:message=>{ alerts.push(message); },
    prompt:()=>({getSelectedButton:()=> 'OK',getResponseText:()=> '理由'})
  };
  context.SpreadsheetApp.getUi=()=>ui;
  context.SpreadsheetApp.getActive=()=>({ getActiveSheet:()=>手動解除シートモック_('ZZZ') });

  context.選択した取り置きを手動解除();

  assert.strictEqual(writes,0);
  assert.ok(alerts.some(m=>String(m).indexOf('一意に特定できません')>=0));
});

function 取り置き表保存モック_(options){
  options=options||{};
  const headers=['取置ID','状態','商品コード'];
  const original=[['OLD1','取り置き中','AAA'],['OLD2','発送済み','BBB'],['','','']];
  const state={data:original.map(r=>r.slice()),dataSetValues:0,clearContent:0,freezeCalls:0};
  const sheet={
    getMaxColumns:()=>headers.length,
    getMaxRows:()=>1+original.length,
    getLastRow:()=>1+original.length,
    insertColumnsAfter:()=>{},
    insertRowsAfter:()=>{},
    getRange:(row,col,numRows,numCols)=>{
      const range={
        setValues(values){
          if(row===1) return range;
          state.dataSetValues++;
          if(options.dataWriteFails) throw new Error('atomic data write failed');
          state.data=values.map(r=>r.slice());
          return range;
        },
        clearContent(){
          state.clearContent++;
          state.data=Array.from({length:numRows},()=>Array(numCols).fill(''));
          return range;
        },
        setFontWeight(){ return range; },setBackground(){ return range; },setFontColor(){ return range; }
      };
      return range;
    },
    setFrozenRows:()=>{
      state.freezeCalls++;
      if(options.freezeFails) throw new Error('freeze failed');
    }
  };
  context.SpreadsheetApp.getActive=()=>({getSheetByName:()=>sheet});
  state.headers=headers;
  state.original=original;
  return state;
}

test('取り置き表保存はclearせず1回のsetValuesで末尾行まで空欄化する', () => {
  const state=取り置き表保存モック_();
  取り置き表保存実装_('取り置き台帳',state.headers,[{取置ID:'NEW1',状態:'取り置き中',商品コード:'CCC'}]);
  const source=fs.readFileSync('Project_24/取り置き台帳.js','utf8');
  const fn=source.slice(source.indexOf('function 取り置き_表を保存_'),source.indexOf('function 取り置き台帳_読む_'));
  assert.strictEqual(fn.includes('.clearContent('),false);
  assert.strictEqual(state.clearContent,0);
  assert.strictEqual(state.dataSetValues,1);
  assert.strictEqual(JSON.stringify(state.data),JSON.stringify([
    ['NEW1','取り置き中','CCC'],['','',''],['','','']
  ]));
});

test('取り置き表のatomic setValues失敗は既存台帳を消さない', () => {
  const state=取り置き表保存モック_({dataWriteFails:true});
  assert.throws(
    ()=>取り置き表保存実装_('取り置き台帳',state.headers,[{取置ID:'NEW1',状態:'取り置き中',商品コード:'CCC'}]),
    /atomic data write failed/
  );
  assert.strictEqual(state.clearContent,0);
  assert.strictEqual(state.dataSetValues,1);
  assert.strictEqual(JSON.stringify(state.data),JSON.stringify(state.original));
});

test('取り置き表はsetFrozenRows失敗時にdata setValuesを試さず既存台帳を保つ', () => {
  const state=取り置き表保存モック_({freezeFails:true});
  assert.throws(
    ()=>取り置き表保存実装_('取り置き台帳',state.headers,[{取置ID:'NEW1',状態:'取り置き中',商品コード:'CCC'}]),
    /freeze failed/
  );
  assert.strictEqual(state.freezeCalls,1);
  assert.strictEqual(state.dataSetValues,0);
  assert.strictEqual(JSON.stringify(state.data),JSON.stringify(state.original));
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
    }),
    getRangeList:()=>({setBackground:()=>{ state.setBackground++; }}) // 背景色は色ごとにRangeListでまとめ書き
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

function 便締めモック_(options){
  options=options||{};
  const externalHeader=['ステータス列','EMS番号','商品コード','数量'];
  const externalRows=[['到着済','EG1','AAA',2],['到着済','EG2','BBB',1]];
  const state={statusWrites:[],externalAttempts:0,movementSaves:0,savedMoves:null,promotions:0,refreshes:0,alerts:[],flushes:0};
  const externalSheet={
    getLastRow:()=>1+externalRows.length,
    getLastColumn:()=>externalHeader.length,
    getRange:(row,col,numRows,numCols)=>{
      const count=numRows||1, width=numCols||1;
      const source=row===1?[externalHeader]:externalRows.slice(row-2,row-2+count);
      const values=source.map(r=>r.slice(col-1,col-1+width));
      return {getValues:()=>values,getDisplayValues:()=>values,getA1Notation:()=> 'A'+row};
    },
    getRangeList:a1s=>({setValue:value=>{
      state.externalAttempts++;
      if(options.externalStatusFails && value==='在庫反映済み') throw new Error('external status failed');
      state.statusWrites.push({a1s:a1s.slice(),value});
    }})
  };
  const jpHeader=['状態','到着日','商品コード','余り数(日本在庫)','EMS番号'];
  const jpRows=options.jpRows||[
    ['到着済','2026-07-12','AAA',2,'EG1'],
    ['到着済','2026-07-12','BBB',1,'EG2']
  ];
  const jpSheet={
    getLastRow:()=>2+jpRows.length,
    getLastColumn:()=>jpHeader.length,
    getRange:(row,col,numRows,numCols)=>{
      const count=numRows||1, width=numCols||1;
      const source=row===2?[jpHeader]:jpRows.slice(row-3,row-3+count);
      const values=source.map(r=>r.slice(col-1,col-1+width));
      return {getValues:()=>values,getDisplayValues:()=>values};
    }
  };
  const ui={
    Button:{OK:'OK'},ButtonSet:{OK:'OK',OK_CANCEL:'OK_CANCEL'},
    alert(...args){ state.alerts.push(args); return 'OK'; },
    prompt:()=>({getSelectedButton:()=> 'OK',getResponseText:()=> 'EG1'})
  };
  const consistency=Object.prototype.hasOwnProperty.call(options,'consistency')
    ? options.consistency : {ts:Date.now(),要確認:0,台帳版:'v1'};
  context.P_KAKUTEI_CFG={シート:'EMSリスト',ヘッダー行:1};
  context.PropertiesService={getDocumentProperties:()=>({getProperty:()=>consistency?JSON.stringify(consistency):''})};
  context.SpreadsheetApp.getUi=()=>ui;
  context.SpreadsheetApp.getActive=()=>({getSheetByName:name=>name==='日本在庫'?jpSheet:null});
  context.SpreadsheetApp.flush=()=>{
    state.flushes++;
    if(options.firstFlushFails && state.flushes===1) throw new Error('external flush failed');
  };
  context.発注共有を開く_=()=>({getSheetByName:()=>externalSheet});
  context.引当履歴_今回到着分を記録_=()=>({追加:0,重複:0,対象:0});
  context.引当履歴_EMSリストから記録_=()=>{ state.promotions++; return {追加:0,重複:0,対象:1}; };
  context.EMS在庫を更新_本体_=()=>{ state.refreshes++; };
  context.EMS在庫移動台帳_読む_=()=>[];
  context.EMS在庫移動台帳_保存_=rows=>{
    state.movementSaves++;
    if(options.movementSaveFails) throw new Error('movement save failed');
    state.savedMoves=rows.map(r=>Object.assign({},r));
  };
  state.execute=()=>context.到着済を在庫反映済みへ本体_();
  return state;
}

test('便締めは選択EMSの日本在庫だけを確認してYahoo移動台帳へ保存する', () => {
  const state=便締めモック_();
  state.execute();
  const alertText=state.alerts.map(args=>args.join('\n')).join('\n');
  assert.ok(alertText.includes('EG1 / AAA / 2'));
  assert.strictEqual(alertText.includes('EG2 / BBB / 1'),false);
  assert.ok(alertText.includes('Yahoo在庫への反映が完了している場合だけOK'));
  assert.strictEqual(state.statusWrites.length,1);
  assert.strictEqual(JSON.stringify(state.statusWrites[0]),JSON.stringify({a1s:['A2'],value:'在庫反映済み'}));
  assert.strictEqual(state.movementSaves,1);
  assert.strictEqual(state.savedMoves.length,1);
  assert.strictEqual(state.savedMoves[0].処理ID,'YAHOO|EMS|EG1|AAA');
  assert.strictEqual(state.promotions,1);
  assert.strictEqual(state.refreshes,1);
});

test('便締めは整合状態が欠落・不正・古い・未来・要確認あり・非v1なら一切変更しない', () => {
  const cases=[
    ['欠落',null],
    ['不正時刻',{ts:'invalid',要確認:0,台帳版:'v1'}],
    ['0時刻',{ts:0,要確認:0,台帳版:'v1'}],
    ['古い',{ts:Date.now()-7*60*60*1000,要確認:0,台帳版:'v1'}],
    ['未来',{ts:Date.now()+60*1000,要確認:0,台帳版:'v1'}],
    ['要確認',{ts:Date.now(),要確認:1,台帳版:'v1'}],
    ['非v1',{ts:Date.now(),要確認:0,台帳版:'legacy'}]
  ];
  cases.forEach(([label,consistency])=>{
    const state=便締めモック_({consistency});
    state.execute();
    assert.strictEqual(state.externalAttempts,0,label+'で外部状態を変更しない');
    assert.strictEqual(state.movementSaves,0,label+'で移動台帳を保存しない');
    assert.strictEqual(state.promotions,0,label+'で履歴昇格しない');
    assert.strictEqual(state.refreshes,0,label+'でEMS更新しない');
  });
});

test('便締めは外部ステータス更新失敗時に移動台帳を保存しない', () => {
  const state=便締めモック_({externalStatusFails:true});
  state.execute();
  assert.strictEqual(state.externalAttempts,1);
  assert.strictEqual(state.movementSaves,0);
  assert.strictEqual(state.promotions,0);
  assert.strictEqual(state.refreshes,0);
});

test('便締めは外部setValue後のflush失敗でも同じセルを到着済へ戻す', () => {
  const state=便締めモック_({firstFlushFails:true});
  state.execute();
  assert.strictEqual(state.movementSaves,0);
  assert.strictEqual(JSON.stringify(state.statusWrites),JSON.stringify([
    {a1s:['A2'],value:'在庫反映済み'},
    {a1s:['A2'],value:'到着済'}
  ]));
  assert.strictEqual(state.flushes,2);
  assert.strictEqual(state.promotions,0);
  assert.strictEqual(state.refreshes,0);
});

test('便締めは移動台帳保存失敗時に同じ外部セルを到着済へ戻して後続を止める', () => {
  const state=便締めモック_({movementSaveFails:true});
  state.execute();
  assert.strictEqual(state.movementSaves,1);
  assert.strictEqual(JSON.stringify(state.statusWrites),JSON.stringify([
    {a1s:['A2'],value:'在庫反映済み'},
    {a1s:['A2'],value:'到着済'}
  ]));
  assert.strictEqual(state.promotions,0);
  assert.strictEqual(state.refreshes,0);
});

function Yahoo確定モック_(options){
  options=options||{};
  const candidate={取置ID:'OLD1',商品コード:'PoEm65（10116569）',数量:1,元EMS番号:'EG0',処理ID:'YAHOO|RETURN|OLD1',確認:true};
  const baseLedger={取置ID:'OLD1',状態:'キャンセル戻し',戻し処理結果:'現物あり',商品コード:'POEM65',元EMS商品コード:'PoEm65（10116569）',取り置き数量:1,元EMS番号:'EG0'};
  const ledger=options.duplicateLedger?[baseLedger,Object.assign({},baseLedger)]:[baseLedger];
  if(options.missingCandidate){ candidate.取置ID='MISSING'; candidate.処理ID='YAHOO|RETURN|MISSING'; }
  if(options.candidatePatch) Object.assign(candidate,options.candidatePatch);
  const candidates=options.duplicateCandidates?[candidate,Object.assign({},candidate)]:[candidate];
  const existing=[{処理ID:'EXISTING',EMS番号:'EG9',商品コード:'ZZZ',数量:1,移動先:'Yahoo即納',処理日時:'old'}];
  const state={moveWrites:[],moveSaveAttempts:0,ledgerWrites:0,savedLedger:null,refreshes:0,alerts:[]};
  const ui={Button:{OK:'OK'},ButtonSet:{OK:'OK',OK_CANCEL:'OK_CANCEL'},alert(...args){ state.alerts.push(args); return 'OK'; }};
  context.SpreadsheetApp.getUi=()=>ui;
  context.取り置き_表を読む_=()=>candidates;
  context.取り置き台帳_読む_=()=>ledger;
  context.EMS在庫移動台帳_読む_=()=>existing;
  context.EMS在庫移動台帳_保存_=rows=>{
    state.moveSaveAttempts++;
    if(options.initialMoveSaveFails && state.moveSaveAttempts===1) throw new Error('initial movement save failed');
    if(options.rollbackFails && state.moveSaveAttempts===2) throw new Error('movement rollback failed');
    state.moveWrites.push(rows.map(r=>Object.assign({},r)));
  };
  context.取り置き台帳_保存_=rows=>{
    state.ledgerWrites++;
    if(options.ledgerSaveFails) throw new Error('ledger save failed');
    state.savedLedger=rows.map(r=>Object.assign({},r));
  };
  context.Yahoo戻し候補を更新_=()=>{ state.refreshes++; };
  state.execute=()=>context.キャンセル戻しをYahoo反映済みにする();
  return state;
}

test('Yahoo戻し確定は同じ取置IDのチェック行が重複していれば保存しない', () => {
  const state=Yahoo確定モック_({duplicateCandidates:true});
  state.execute();
  assert.strictEqual(state.moveSaveAttempts,0);
  assert.strictEqual(state.ledgerWrites,0);
  assert.strictEqual(state.refreshes,0);
});

test('Yahoo戻し確定は候補のraw商品コード・数量・EMS・処理ID改ざんを拒否する', () => {
  [
    ['商品コード',{商品コード:'POEM65'}],
    ['数量',{数量:2}],
    ['EMS',{元EMS番号:'EGX'}],
    ['処理ID',{処理ID:'YAHOO|RETURN|TAMPERED'}]
  ].forEach(([label,candidatePatch])=>{
    const state=Yahoo確定モック_({candidatePatch});
    state.execute();
    assert.strictEqual(state.moveSaveAttempts,0,label+'改ざんで移動を保存しない');
    assert.strictEqual(state.ledgerWrites,0,label+'改ざんで台帳を保存しない');
    assert.strictEqual(state.refreshes,0,label+'改ざんで候補を更新しない');
  });
});

test('Yahoo戻し確定は選択IDが現台帳で欠落または重複なら保存しない', () => {
  [
    ['欠落',{missingCandidate:true}],
    ['重複',{duplicateLedger:true}]
  ].forEach(([label,options])=>{
    const state=Yahoo確定モック_(options);
    state.execute();
    assert.strictEqual(state.moveWrites.length,0,label+'で移動を保存しない');
    assert.strictEqual(state.ledgerWrites,0,label+'で取り置き台帳を保存しない');
    assert.strictEqual(state.refreshes,0,label+'で候補を更新しない');
  });
});

test('Yahoo戻し確定は最初の移動保存失敗を復旧済みと表示しない', () => {
  const state=Yahoo確定モック_({initialMoveSaveFails:true});
  state.execute();
  assert.strictEqual(state.moveSaveAttempts,1);
  assert.strictEqual(state.moveWrites.length,0);
  assert.strictEqual(state.ledgerWrites,0);
  assert.strictEqual(state.refreshes,0);
  assert.ok(state.alerts.some(args=>args[0]==='Yahoo移動台帳の保存に失敗しました'));
  assert.strictEqual(state.alerts.some(args=>String(args[0]).includes('元へ戻しました')),false);
});

test('Yahoo戻し確定は移動保存後に取り置き台帳をYahoo反映済みへ更新する', () => {
  const state=Yahoo確定モック_();
  state.execute();
  assert.strictEqual(state.moveWrites.length,1);
  assert.strictEqual(state.moveWrites[0][1].処理ID,'YAHOO|RETURN|OLD1');
  assert.strictEqual(state.moveWrites[0][1].商品コード,'PoEm65（10116569）');
  assert.strictEqual(state.savedLedger[0].戻し処理結果,'Yahoo反映済み');
  assert.strictEqual(state.refreshes,1);
});

test('Yahoo戻し確定は取り置き台帳保存失敗時に移動台帳を元へ戻す', () => {
  const state=Yahoo確定モック_({ledgerSaveFails:true});
  state.execute();
  assert.strictEqual(state.ledgerWrites,1);
  assert.strictEqual(state.moveWrites.length,2);
  assert.strictEqual(state.moveWrites[0].length,2);
  assert.strictEqual(JSON.stringify(state.moveWrites[1].map(r=>r.処理ID)),JSON.stringify(['EXISTING']));
  assert.strictEqual(state.refreshes,0);
  assert.ok(state.alerts.some(args=>args[0]==='取り置き台帳の保存に失敗したためYahoo移動台帳を元へ戻しました'));
});

test('Yahoo戻し確定は移動ロールバック失敗を復旧済みと表示しない', () => {
  const state=Yahoo確定モック_({ledgerSaveFails:true,rollbackFails:true});
  state.execute();
  assert.strictEqual(state.ledgerWrites,1);
  assert.strictEqual(state.moveSaveAttempts,2);
  assert.strictEqual(state.moveWrites.length,1);
  assert.strictEqual(state.refreshes,0);
  assert.ok(state.alerts.some(args=>args[0]==='取り置き台帳の保存とYahoo移動台帳の復旧に失敗しました'));
  assert.strictEqual(state.alerts.some(args=>String(args[0]).includes('元へ戻しました')),false);
});

test('受注共通メニューに初回登録の作成と確定を追加する', () => {
  const source=fs.readFileSync('Project_24/引当.js','utf8');
  assert.ok(source.includes(".addItem('📋 取り置き登録を作成(候補の洗い替え)', '取り置き初期登録を作成')"));
  assert.ok(source.includes(".addItem('✅ 取り置き登録を確定', '取り置き初期登録を確定')"));
  assert.ok(source.includes(".addItem('📦 キャンセル戻し確認を更新', 'キャンセル戻し確認を更新')"));
  assert.ok(source.includes(".addItem('✅ キャンセル戻し確認を確定', 'キャンセル戻し確認を確定')"));
  assert.ok(source.includes(".addItem('🛒 Yahoo戻しを反映済みにする', 'キャンセル戻しをYahoo反映済みにする')"));
  assert.ok(source.includes(".addItem('🔓 選択した取り置きを手動解除', '選択した取り置きを手動解除')"));
});

// ===== 修正5: 初期候補の範囲拡大・初期確定の再実行ガード・移動台帳の数量不一致ガード =====

test('初期候補は出荷GO未入金・出荷可能の受注も含め、状態のグループ順に並ぶ', () => {
  const orders=[
    {ban:'101',code:'AAA',sku:'AAAb',qty:2,kbn:'取り寄せ'},
    {ban:'102',code:'BBB',sku:'BBBb',qty:1,kbn:'取り寄せ'},
    {ban:'103',code:'CCC',sku:'CCCb',qty:1,kbn:'取り寄せ'},
    {ban:'104',code:'DDD',sku:'DDDb',qty:1,kbn:'取り寄せ'},
    {ban:'105',code:'EEE',sku:'EEEb',qty:1,kbn:'取り寄せ'}
  ];
  const rows=context.取り置き_初期候補_(orders,[
    {状態:'出荷可能',bans:new Set(['104'])},
    {状態:'出荷GO未入金',bans:new Set(['103'])},
    {状態:'部分在庫',bans:new Set(['101'])},
    {状態:'希望日待ち',bans:new Set(['102'])}
  ]);
  assert.strictEqual(rows.length,4);
  assert.ok(rows.every(r=>r.受注番号!=='105'));
  assert.strictEqual(JSON.stringify(rows.map(r=>r.現在の状態)),
    JSON.stringify(['出荷可能','出荷GO未入金','部分在庫','希望日待ち']),'一覧の優先順でグループ化');
});

test('発送済みになった開始前在庫行は初期確定の再実行で復活も消滅もしない', () => {
  const existing=[{取置ID:'INIT|101|AAA|AAAB',状態:'発送済み',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:1,取置元種別:'開始前在庫'}];
  // 数量入りで再確定 → 復活させずエラー
  const withQty=context.取り置き_初期確定計画_([
    {取置ID:'INIT|101|AAA|AAAB',受注番号:'101',商品コード:'AAA',SKU:'AAAb',注文数量:2,現物取り置き数量:1}
  ],existing,'2026-07-16 10:00:00');
  assert.ok(withQty.errors.some(e=>/101/.test(e)),'発送済みの初期行の上書きはエラー: '+JSON.stringify(withQty.errors));
  // 空欄で再確定 → 行を消さない
  const blank=context.取り置き_初期確定計画_([
    {取置ID:'INIT|101|AAA|AAAB',受注番号:'101',商品コード:'AAA',SKU:'AAAb',注文数量:2,現物取り置き数量:''}
  ],existing,'2026-07-16 10:00:00');
  assert.strictEqual(JSON.stringify(blank.errors),'[]');
  assert.strictEqual(blank.rows.length,1,'発送済み行は保持される');
  assert.strictEqual(blank.rows[0].状態,'発送済み');
});

test('同じ処理IDで数量が違うYahoo移動は黙って捨てずエラーにする', () => {
  const existing=[{処理ID:'YAHOO|EMS|EG1|AAA',EMS番号:'EG1',商品コード:'AAA',数量:2,移動先:'Yahoo即納',処理日時:'2026-07-15 10:00:00'}];
  const same=context.EMS在庫移動_追加計画_([{処理ID:'YAHOO|EMS|EG1|AAA',EMS番号:'EG1',商品コード:'AAA',数量:2}],existing,'2026-07-16 10:00:00');
  assert.strictEqual(same.errors.length,0,'同数量の再実行は従来通りスキップ');
  assert.strictEqual(same.added.length,0);
  const diff=context.EMS在庫移動_追加計画_([{処理ID:'YAHOO|EMS|EG1|AAA',EMS番号:'EG1',商品コード:'AAA',数量:3}],existing,'2026-07-16 10:00:00');
  assert.strictEqual(diff.errors.length,1,'数量が変わった同一処理IDは記録漏れの兆候としてエラー');
  assert.ok(diff.errors[0].indexOf('数量')>=0);
});

// ===== P列の手動名指し(コード不一致の救済)を新計画が引き継ぐ =====

test('P手動名指し解析: 単一の受注番号だけを名指しとして読む', () => {
  assert.strictEqual(context.P手動名指し解析_('10117376'),'10117376');
  assert.strictEqual(context.P手動名指し解析_('10117376:1'),'10117376');
  assert.strictEqual(context.P手動名指し解析_('101:1, 102:2'),'','複数名指しは対象外');
  assert.strictEqual(context.P手動名指し解析_('メモ書き'),'');
  assert.strictEqual(context.P手動名指し解析_(''),'');
});

test('発注共有P列計画: 説明文コード+P列手動名指しの行はdirect扱いで名指しを保持する', () => {
  const recvHead=['受注番号','注文日時','商品コード','商品SKU','項目・選択肢','個数','入荷日'];
  const recvRows=[recvHead,['10117376',new Date('2026-07-05'),'HOTOPIK','HOTOPIKb','取り寄せ',1,'']];
  const recv={
    getLastRow:()=>recvRows.length,
    getLastColumn:()=>recvHead.length,
    getDataRange:()=>({getValues:()=>recvRows}),
    getRange:(row,col,numRows,numCols)=>({getValues:()=>recvRows.slice(row-1,row-1+numRows).map(r=>r.slice(col-1,col-1+numCols))})
  };
  const emsHead=['ステータス列','購入No.','商品コード','数量','注文番号','EMS到着日','EMS番号'];
  const emsRows=[['到着済','20260708_01_01','핫 토픽 Hot Topik 2 읽기',1,'10117376','2026-07-12','EG049827401KR']];
  const ems={
    getLastRow:()=>7,
    getLastColumn:()=>emsHead.length,
    getSheetId:()=>123,
    getParent:()=>({getId:()=> 'BOOK1'}),
    getRange:(row,col,numRows,numCols)=>({
      getValues:()=>row===6?[emsHead.slice(col-1,col-1+numCols)]:emsRows.slice(row-7,row-7+numRows).map(r=>r.slice(col-1,col-1+numCols)),
      getDisplayValues:()=>row===6?[emsHead.slice(col-1,col-1+numCols)]:emsRows.slice(row-7,row-7+numRows).map(r=>r.slice(col-1,col-1+numCols)),
      setValues:()=>{ throw new Error('計画段階で書き込みしない'); },
      setBackground:()=>{ throw new Error('計画段階で書き込みしない'); }
    })
  };
  context.P_KAKUTEI_CFG={シート:'EMSリスト',ヘッダー行:6};
  context.SpreadsheetApp.getActive=()=>({getSheetByName:()=>recv});
  context.発注共有を開く_=()=>({getSheetByName:()=>ems});
  context.取り置き台帳_読む_=()=>[];
  context.EMS在庫移動台帳_読む_=()=>[];

  const plan=context.発注共有P列計画_();

  assert.strictEqual(plan.error,'');
  const row=plan.rows[0];
  assert.strictEqual(row.directBan,'10117376','手動名指しをdirectとして扱う');
  assert.strictEqual(row.手動名指し,true);
  assert.strictEqual(plan.values[0][0],'10117376','名指しは消えずP列に残る');
  const map=context.P列救済供給マップ_(plan.rows);
  assert.strictEqual(map[context.取り置き_供給キー_('EG049827401KR','핫 토픽 Hot Topik 2 읽기')],'10117376');
});

test('発注共有P列計画: コードが注文と一致する行の手動Pは従来通り計画で上書きされ得る', () => {
  const recvHead=['受注番号','注文日時','商品コード','商品SKU','項目・選択肢','個数','入荷日'];
  const recvRows=[recvHead,['102',new Date('2026-07-02'),'AAA','AAAb','取り寄せ',2,'']];
  const recv={
    getLastRow:()=>recvRows.length,
    getLastColumn:()=>recvHead.length,
    getDataRange:()=>({getValues:()=>recvRows}),
    getRange:(row,col,numRows,numCols)=>({getValues:()=>recvRows.slice(row-1,row-1+numRows).map(r=>r.slice(col-1,col-1+numCols))})
  };
  const emsHead=['ステータス列','購入No.','商品コード','数量','注文番号','EMS到着日','EMS番号'];
  const emsRows=[['到着済','20260701_01_01','AAA',2,'99999999','2026-07-12','EG1']];
  const ems={
    getLastRow:()=>7,
    getLastColumn:()=>emsHead.length,
    getSheetId:()=>123,
    getParent:()=>({getId:()=> 'BOOK1'}),
    getRange:(row,col,numRows,numCols)=>({
      getValues:()=>row===6?[emsHead.slice(col-1,col-1+numCols)]:emsRows.slice(row-7,row-7+numRows).map(r=>r.slice(col-1,col-1+numCols)),
      getDisplayValues:()=>row===6?[emsHead.slice(col-1,col-1+numCols)]:emsRows.slice(row-7,row-7+numRows).map(r=>r.slice(col-1,col-1+numCols)),
      setValues:()=>{},
      setBackground:()=>{}
    })
  };
  context.P_KAKUTEI_CFG={シート:'EMSリスト',ヘッダー行:6};
  context.SpreadsheetApp.getActive=()=>({getSheetByName:()=>recv});
  context.発注共有を開く_=()=>({getSheetByName:()=>ems});
  context.取り置き台帳_読む_=()=>[];
  context.EMS在庫移動台帳_読む_=()=>[];

  const plan=context.発注共有P列計画_();

  assert.strictEqual(plan.error,'');
  assert.strictEqual(plan.rows[0].directBan,'','コード一致の行は救済対象にしない');
  assert.strictEqual(plan.values[0][0],'102','計画がFIFOで上書き');
});

// ===== 受注番号集合: ④の一覧シートは1行目がタイムスタンプ・見出しは6行目付近(書き出し_のstartRow) =====

test('受注番号集合は見出し行を探して読む(1行目固定にしない)', () => {
  // 実際の部分在庫シートのレイアウト: A1=最終引当タイムスタンプ、6行目=見出し、7行目〜データ
  const values=[
    ['最終引当: 2026/07/16 13:00:00 / 2行','','',''],
    ['','','',''],['','','',''],['','','',''],['','','',''],
    ['受注番号','氏名','商品コード','個数'],
    ['10117249','西野 瑠璃','JMEE-ANS-01-21',2],
    ['10117275','誰か','KRSJCM20-01S',1],
    ['','','','']
  ];
  const sheet={
    getLastRow:()=>values.length,
    getLastColumn:()=>4,
    getDataRange:()=>({getDisplayValues:()=>values.map(r=>r.map(v=>String(v==null?'':v)))})
  };
  context.SpreadsheetApp.getActive=()=>({getSheetByName:()=>sheet});
  const out=context.取り置き_受注番号集合_('部分在庫');
  assert.strictEqual(out.size,2);
  assert.ok(out.has('10117249'));
  assert.ok(out.has('10117275'));
});

test('受注番号集合はシート無し・見出し無し(未生成)なら空集合を返す', () => {
  context.SpreadsheetApp.getActive=()=>({getSheetByName:()=>null});
  assert.strictEqual(context.取り置き_受注番号集合_('部分在庫').size,0);
  const empty={
    getLastRow:()=>1,
    getLastColumn:()=>1,
    getDataRange:()=>({getDisplayValues:()=>[['']]})
  };
  context.SpreadsheetApp.getActive=()=>({getSheetByName:()=>empty});
  assert.strictEqual(context.取り置き_受注番号集合_('出荷可能').size,0);
});

// ===== 登録絞り込み: 自動除外は予約中・出荷GOだけ =====

test('登録絞り込みは予約中・出荷GOだけを数量入力済みでも除外する', () => {
  const rows=[
    {取置ID:'A',受注ステータス:'処理済①',旧入荷日:'2026-06-25',棚確認:'',現物取り置き数量:''},   // 着済のはず→残す
    {取置ID:'B',受注ステータス:'予約中',旧入荷日:'2026-06-25',棚確認:'予約',現物取り置き数量:''}, // 幽霊スタンプ→落とす
    {取置ID:'C',受注ステータス:'出荷待/取寄せ',旧入荷日:'',棚確認:'',現物取り置き数量:''},        // 帳簿上未着でも表示
    {取置ID:'D',受注ステータス:'予約中',旧入荷日:'',棚確認:'予約',現物取り置き数量:2},            // 入力済みでも予約中は落とす
    {取置ID:'E',受注ステータス:'出荷待/希望日',旧入荷日:'',棚確認:'',現物取り置き数量:2},         // 確定済み→残す
    {取置ID:'F',受注ステータス:'■出荷GO',旧入荷日:'2026-07-12',棚確認:'',現物取り置き数量:''},    // 出荷作業に入る注文→落とす
    {取置ID:'G',受注ステータス:'■出荷GO',旧入荷日:'2026-07-12',棚確認:'',現物取り置き数量:1},     // 入力済みでも出荷GOは落とす(台帳側で自動追随)
    {取置ID:'H',受注ステータス:'部分包装',旧入荷日:'',棚確認:'',現物取り置き数量:''},             // 梱包中=現物あり→スタンプ無しでも必ず出す
    {取置ID:'I',受注ステータス:'部分包装',旧入荷日:'',棚確認:'予約',現物取り置き数量:''},         // 手動予約の非表示は棚確認記憶が担当
    {取置ID:'J',受注ステータス:'予約受付終了',旧入荷日:'',棚確認:'',現物取り置き数量:''}          // 「予約中」ではないので表示
  ];
  const out=context.取り置き_登録絞り込み_(rows);
  assert.strictEqual(JSON.stringify(out.map(r=>r.取置ID)),JSON.stringify(['A','C','E','H','I','J']));
});

// ===== 棚確認の判断済み(出荷済み/未着/予約)は記憶して以後表示しない =====

test('出荷済み/未着/予約と判断した行は非表示になり記憶される・注文が消えれば記憶も消える', () => {
  const candidates=[
    {取置ID:'A',棚確認:'出荷済み',現物取り置き数量:'',受注ステータス:'部分包装'},
    {取置ID:'B',棚確認:'',現物取り置き数量:'',受注ステータス:'部分包装'},
    {取置ID:'C',棚確認:'',現物取り置き数量:2,受注ステータス:'部分包装'},
    {取置ID:'D',棚確認:'予約',現物取り置き数量:'',受注ステータス:'部分包装'}
  ];
  const first=context.取り置き_棚確認記憶を適用_(candidates,{});
  assert.strictEqual(JSON.stringify(first.rows.map(r=>r.取置ID)),JSON.stringify(['B','C']),'判断済みA・Dは非表示');
  assert.strictEqual(first.store.A,'出荷済み','判断を記憶');
  assert.strictEqual(first.store.D,'予約','手動予約を記憶');
  // 次回: シート側の棚確認が空でも記憶から復元されて非表示のまま
  const second=context.取り置き_棚確認記憶を適用_(
    [{取置ID:'A',棚確認:'',現物取り置き数量:''},{取置ID:'B',棚確認:'未着',現物取り置き数量:''},
     {取置ID:'D',棚確認:'',現物取り置き数量:''}],first.store);
  assert.strictEqual(second.rows.length,0,'A・Dは記憶で非表示・Bは今回の判断で非表示');
  assert.strictEqual(second.store.A,'出荷済み');
  assert.strictEqual(second.store.B,'未着');
  assert.strictEqual(second.store.D,'予約');
  // 注文が候補から消えたら記憶も自動で消える
  const third=context.取り置き_棚確認記憶を適用_([{取置ID:'B',棚確認:'',現物取り置き数量:''}],second.store);
  assert.strictEqual(third.store.A,undefined,'消えた注文の記憶は掃除');
  // 数量を入れた行は判断済みでも表示される(矛盾は確定時ガードが検知)
  const qty=context.取り置き_棚確認記憶を適用_([{取置ID:'B',棚確認:'',現物取り置き数量:1}],third.store);
  assert.strictEqual(qty.rows.length,1);
});

process.exitCode = failures ? 1 : 0;
console.log(failures ? `FAILURES: ${failures}` : 'ALL TESTS PASSED');
