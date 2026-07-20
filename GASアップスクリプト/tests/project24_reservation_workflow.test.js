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
  PropertiesService: { getDocumentProperties: () => ({ getProperty: () => null, setProperty: () => {}, deleteProperty: () => {} }) }, // 全件再計算のブロックSKU/ガード用（未設定=null）
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
  'Project_24/P列自動記入.js',
  'Project_24/全件再計算.js' // P列計画がブロックSKU判定で参照する(GAS実行時は全ファイル同居)
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

test('予約判定は未来の発売予定だけを自動予約にする', () => {
  const today=new Date(2026,6,17);
  assert.strictEqual(context.取り置き_予約判定_('予約9月韓国発売予定','',today),true);
  assert.strictEqual(context.取り置き_予約判定_('予約7月末韓国発売予定','',today),true);
  assert.strictEqual(context.取り置き_予約判定_('予約5月韓国発売予定','',today),false);
  assert.strictEqual(context.取り置き_予約判定_('','予約早期完売',today),false);
  assert.strictEqual(context.取り置き_予約判定_('予約2026/7/18韓国発売予定','',today),true,'年月日がtodayより後');
  assert.strictEqual(context.取り置き_予約判定_('予約2026/7/17韓国発売予定','',today),false,'当日は未来ではない');
  assert.strictEqual(context.取り置き_予約判定_('韓国発売予定','',today),false,'予約表記なし');

  const rows=context.取り置き_初期候補_([
    {ban:'201',code:'AAA',sku:'AAAb',qty:1,予約:true},
    {ban:'202',code:'BBB',sku:'BBBb',qty:1,予約:false},
    {ban:'203',code:'CCC',sku:'CCCb',qty:1,予約:true,入荷日:'2026-07-11'},
    {ban:'204',code:'DDD',sku:'DDDb',qty:1,予約:true,ステータス:'部分包装'}
  ],[
    {状態:'着済スタンプ(要棚確認)',bans:new Set(['201','202','203'])},
    {状態:'部分包装(要棚確認)',bans:new Set(['204'])}
  ]);
  assert.strictEqual(rows.find(r=>r.受注番号==='201').棚確認,'予約','未来予約を自動選択');
  assert.strictEqual(rows.find(r=>r.受注番号==='202').棚確認,'','過去日・日付不明は自動選択しない');
  assert.strictEqual(rows.find(r=>r.受注番号==='203').棚確認,'','旧入荷日があれば現物形跡を優先');
  assert.strictEqual(rows.find(r=>r.受注番号==='204').棚確認,'','部分包装なら現物形跡を優先');
});

test('自動予約は隠さず「予約」の既定値付きで表示し、発売日が来れば既定値が外れる', () => {
  const sources=[{状態:'予約候補',bans:new Set(['201'])}];
  const future=context.取り置き_初期候補_([
    {ban:'201',code:'AAA',sku:'AAAb',qty:1,予約:true}
  ],sources);
  const first=context.取り置き_棚確認記憶を適用_(future,{});
  assert.strictEqual(first.rows.length,1,'未発売の自動予約も隠さない(棚確認「予約」の紫書式で表示)');
  assert.strictEqual(first.rows[0].棚確認,'予約','既定値として予約を選択');
  assert.strictEqual(JSON.stringify(first.store),'{}','非表示スイッチの記憶は廃止');

  const released=context.取り置き_初期候補_([
    {ban:'201',code:'AAA',sku:'AAAb',qty:1,予約:false}
  ],sources);
  const second=context.取り置き_棚確認記憶を適用_(released,{});
  assert.strictEqual(second.rows.length,1);
  assert.strictEqual(second.rows[0].棚確認,'','発売日が来れば既定値が外れて通常の棚確認へ');

  // シートで選んだ「予約」は棚確認セルごと引き継がれて表示され続ける(記憶はシートの列が担う)
  const manual=context.取り置き_登録シート引き継ぎ_(future,[
    {取置ID:future[0].取置ID,棚確認:'予約',現物取り置き数量:''}
  ],[]);
  const third=context.取り置き_棚確認記憶を適用_(manual,{});
  assert.strictEqual(third.rows.length,1,'手動の予約判断も隠さない');
  assert.strictEqual(third.rows[0].棚確認,'予約');
  assert.strictEqual(JSON.stringify(third.store),'{}');
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

test('分割明細の後続部分包装を現物形跡として優先する', () => {
  const rows=context.取り置き_初期候補_([
    {ban:'301',code:'PART',sku:'PARTb',qty:1,予約:true,ステータス:'出荷待/取寄せ'},
    {ban:'301',code:'PART',sku:'PARTb',qty:1,予約:true,ステータス:'部分包装'}
  ],[{状態:'候補',bans:new Set(['301'])}]);
  assert.strictEqual(rows[0].棚確認,'','後続行が部分包装なら自動予約より現物形跡を優先');
  assert.ok(rows[0].受注ステータス.indexOf('出荷待/取寄せ')>=0);
  assert.ok(rows[0].受注ステータス.indexOf('部分包装')>=0,'全分割行のステータスを既存列へ集約');
});

test('分割明細: 出荷GOは除外するが後続予約中の注文は残す(全行表示化v2)', () => {
  const rows=context.取り置き_初期候補_([
    {ban:'302',code:'RSV',sku:'RSVb',qty:1,ステータス:'出荷待/取寄せ'},
    {ban:'302',code:'RSV',sku:'RSVb',qty:1,ステータス:'予約中'},         // 後続が予約中でも隠さない(早着の可能性)
    {ban:'303',code:'GO',sku:'GOb',qty:1,ステータス:'出荷待/取寄せ'},
    {ban:'303',code:'GO',sku:'GOb',qty:1,ステータス:'■出荷GO'}           // 出荷GOは除外
  ],[{状態:'候補',bans:new Set(['302','303'])}]);
  const out=context.取り置き_登録絞り込み_(rows);
  assert.strictEqual(out.length,1,'出荷GO(303)だけ除外し、予約中を含む302は残す');
  assert.strictEqual(out[0].受注番号,'302');
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

test('受注共通メニューは取り置き登録の更新・一括反映と内部管理サブメニューに再編する', () => {
  const source=fs.readFileSync('Project_24/引当.js','utf8');
  assert.ok(source.includes(".addItem('📋 取り置き登録を更新', '取り置き初期登録を作成')"));
  assert.ok(source.includes(".addItem('✅ 取り置き登録を反映(通常＋棚戻し)', '取り置き登録を反映')"));
  assert.ok(source.includes(".addItem('🧱 取り置き登録の罫線を引き直す', '取り置き登録の罫線を引く')"));
  // キャンセル戻し確認などは通常操作不要のサブメニューへ格下げ
  assert.ok(source.includes("ui.createMenu('🔧 取り置き内部管理(通常操作不要)')"));
  assert.ok(source.includes(".addItem('キャンセル戻し確認を更新', 'キャンセル戻し確認を更新')"));
  assert.ok(source.includes(".addItem('キャンセル戻し確認を確定', 'キャンセル戻し確認を確定')"));
  assert.ok(source.includes(".addItem('Yahoo戻しを反映済みにする', 'キャンセル戻しをYahoo反映済みにする')"));
  assert.ok(source.includes(".addItem('選択した取り置きを手動解除', '選択した取り置きを手動解除')"));
});

test('旧「確定」関数と図形ボタンの割り当ては新しい一括反映へつながる', () => {
  const source=fs.readFileSync('Project_24/取り置き台帳.js','utf8');
  assert.ok(source.includes('function 取り置き初期登録を確定(){ 取り置き登録を反映(); }'));
  assert.strictEqual(typeof context.取り置き登録を反映,'function');
});

test('①CSV取込は取り置き登録を自動更新し、失敗しても取込自体は完了させる', () => {
  const source=fs.readFileSync('Project_24/引当.js','utf8');
  assert.ok(source.includes('取り置き初期登録を作成本体_({silent:true})'));
  assert.ok(source.includes('取り置き登録の自動更新失敗'));
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

// ===== 登録絞り込み: 自動除外は出荷GOだけ(2026-07-20 全行表示化v2で予約中の除外は廃止) =====
// アクティブ注文の中の行だけが候補に来るので、予約中で証拠が無くても隠さず表示する
// (棚に早着している可能性がある＝入れる行が無い、を防ぐ)。純粋な予約だけの注文は
// そもそも候補に入らないのでノイズは増えない。

test('登録絞り込みは出荷GOだけを除外し、予約中は証拠が無くても残す', () => {
  const rows=[
    {取置ID:'A',受注ステータス:'処理済①',旧入荷日:'2026-06-25',棚確認:'',現物取り置き数量:''},   // 着済のはず→残す
    {取置ID:'B',受注ステータス:'予約中',旧入荷日:'2026-06-25',棚確認:'予約',現物取り置き数量:''}, // 早着の証拠→残す
    {取置ID:'C',受注ステータス:'出荷待/取寄せ',旧入荷日:'',棚確認:'',現物取り置き数量:''},        // 帳簿上未着でも表示
    {取置ID:'D',受注ステータス:'予約中',旧入荷日:'',棚確認:'予約',現物取り置き数量:2},            // 数量入力あり→表示
    {取置ID:'E',受注ステータス:'出荷待/希望日',旧入荷日:'',棚確認:'',現物取り置き数量:2},         // 確定済み→残す
    {取置ID:'F',受注ステータス:'■出荷GO',旧入荷日:'2026-07-12',棚確認:'',現物取り置き数量:''},    // 出荷作業に入る注文→落とす
    {取置ID:'G',受注ステータス:'■出荷GO',旧入荷日:'2026-07-12',棚確認:'',現物取り置き数量:1},     // 入力済みでも出荷GOは落とす(台帳側で自動追随)
    {取置ID:'H',受注ステータス:'部分包装',旧入荷日:'',棚確認:'',現物取り置き数量:''},             // 梱包中=現物あり→表示
    {取置ID:'I',受注ステータス:'部分包装',旧入荷日:'',棚確認:'予約',現物取り置き数量:''},         // 手動予約も隠さない
    {取置ID:'J',受注ステータス:'予約受付終了',旧入荷日:'',棚確認:'',現物取り置き数量:''},         // 「予約中」ではないので表示
    {取置ID:'K',受注ステータス:'予約中',旧入荷日:'',旧EMS:'',棚確認:'',現物取り置き数量:''},      // 証拠なしの予約中も今は残す(早着の可能性)
    {取置ID:'L',受注ステータス:'予約中',旧入荷日:'',旧EMS:'EG1',棚確認:'',現物取り置き数量:''},   // 旧EMSあり→残す
    {取置ID:'M',受注ステータス:'予約中',旧入荷日:'',台帳確保数:1,棚確認:'',現物取り置き数量:''}   // ④の台帳確保あり→残す
  ];
  const out=context.取り置き_登録絞り込み_(rows);
  assert.strictEqual(JSON.stringify(out.map(r=>r.取置ID)),JSON.stringify(['A','B','C','D','E','H','I','J','K','L','M']));
});

test('登録絞り込みは出荷GOを無条件で除外し、予約中は証拠が無くても残す', () => {
  const rows=[
    {取置ID:'OLD-RESERVED',受注ステータス:'予約中',旧入荷日:'2026-06-25',棚確認:'',現物取り置き数量:''},    // 早着した予約品→残す
    {取置ID:'PACKED-RESERVED',受注ステータス:'部分包装 / 予約中',旧入荷日:'',棚確認:'',現物取り置き数量:''}, // 証拠なし予約中も残す(早着の可能性)
    {取置ID:'PACKED-GO',受注ステータス:'部分包装 / ■出荷GO',旧入荷日:'',棚確認:'',現物取り置き数量:''},      // 出荷GO→落とす
    {取置ID:'GO-ARRIVED',受注ステータス:'■出荷GO',旧入荷日:'2026-06-25',棚確認:'',現物取り置き数量:2},       // 証拠があっても出荷GOは除外
    {取置ID:'PACKED',受注ステータス:'部分包装',旧入荷日:'',棚確認:'',現物取り置き数量:''}
  ];

  const out=context.取り置き_登録絞り込み_(rows);

  assert.strictEqual(JSON.stringify(out.map(r=>r.取置ID)),JSON.stringify(['OLD-RESERVED','PACKED-RESERVED','PACKED']));
});

// ===== 2026-07-20 全行表示化: 判断済み(出荷済み/未着/予約)も隠さず表示する =====

test('判断済み(出荷済み/未着/予約)の行も隠さず表示し、旧記憶は棚確認の既定値として一度だけ引き継ぐ', () => {
  const candidates=[
    {取置ID:'A',棚確認:'出荷済み',現物取り置き数量:'',受注ステータス:'部分包装'},
    {取置ID:'B',棚確認:'',現物取り置き数量:'',受注ステータス:'部分包装'},
    {取置ID:'C',棚確認:'',現物取り置き数量:2,受注ステータス:'部分包装'},
    {取置ID:'D',棚確認:'予約',現物取り置き数量:'',受注ステータス:'部分包装'}
  ];
  const first=context.取り置き_棚確認記憶を適用_(candidates,{});
  assert.strictEqual(JSON.stringify(first.rows.map(r=>r.取置ID)),JSON.stringify(['A','B','C','D']),'判断済みも全行表示');
  assert.strictEqual(first.rows[0].棚確認,'出荷済み','判断は棚確認列に残す(赤太字の条件付き書式で目立たせる)');
  assert.strictEqual(first.rows[3].棚確認,'予約');
  assert.strictEqual(JSON.stringify(first.store),'{}','非表示スイッチの記憶は廃止');
  // 旧バージョンが残したDocumentPropertiesの判断は、棚確認セルの既定値として一度だけ復元する
  const legacy=context.取り置き_棚確認記憶を適用_(
    [{取置ID:'A',棚確認:'',現物取り置き数量:''}],{A:'出荷済み'});
  assert.strictEqual(legacy.rows.length,1,'復元しても隠さない');
  assert.strictEqual(legacy.rows[0].棚確認,'出荷済み','旧記憶を棚確認セルへ引き継ぐ');
  assert.strictEqual(JSON.stringify(legacy.store),'{}','以後はシートの棚確認列が記憶=セルを空にすれば次回から消える');
  // 到着証拠(旧入荷日など)が付いた「予約」は既定値を外して棚へ出す(早着した予約品を隠れたままにしない)
  const arrived=context.取り置き_棚確認記憶を適用_(
    [{取置ID:'E',棚確認:'予約',現物取り置き数量:'',旧入荷日:'2026-06-25'}],{});
  assert.strictEqual(arrived.rows.length,1);
  assert.strictEqual(arrived.rows[0].棚確認,'','証拠が付いた予約は既定値を外して要棚確認へ回す');
  // 数量を入れた行は従来どおり表示される(矛盾は反映時ガードが検知)
  const qty=context.取り置き_棚確認記憶を適用_([{取置ID:'B',棚確認:'出荷済み',現物取り置き数量:1}],{});
  assert.strictEqual(qty.rows.length,1);
  assert.strictEqual(qty.rows[0].棚確認,'出荷済み');
});

test('要棚確認を優先しても同じ注文の商品行を分断しない', () => {
  const rows=context.取り置き_注文単位で並べる_([
    {受注番号:'200',商品コード:'B1',判定:''},
    {受注番号:'100',商品コード:'A1',判定:'要棚確認'},
    {受注番号:'100',商品コード:'A2',判定:''},
    {受注番号:'300',商品コード:'C1',判定:''}
  ]);
  assert.strictEqual(JSON.stringify(rows.map(r=>r.受注番号)),JSON.stringify(['100','100','200','300']));
});

test('取り置き登録の注文境界は同じ注文の最終行を全17列で返す', () => {
  assert.strictEqual(JSON.stringify(Array.from(context.取り置き_注文境界A1_([
    {受注番号:'100'},{受注番号:'100'},{受注番号:'200'}
  ]))),JSON.stringify(['A3:Q3','A4:Q4']));
});

test('棚確認の条件付き書式定義は5ステータスを決められた順と色で返す', () => {
  const defs=context.取り置き_棚確認書式定義_(11,2);
  assert.strictEqual(JSON.stringify(defs.map(d=>[d.値,d.背景])),JSON.stringify([
    ['発送待ち','#cfe2f3'],
    ['部分在庫','#d9ead3'],
    ['出荷済み','#f4cccc'],
    ['未着','#d9d9d9'],
    ['予約','#d9d2e9']
  ]));
  assert.strictEqual(defs[2].条件,'=$K2="出荷済み"');
  assert.strictEqual(defs[2].文字色,'#990000');
  assert.strictEqual(defs[2].太字,true);
});

test('棚確認の条件付き書式は対象商品行の全列へ固定5ルールを一括設定する', () => {
  const calls={ranges:[],saved:[]};
  context.SpreadsheetApp.newConditionalFormatRule=()=>{
    const rule={};
    const builder={
      whenFormulaSatisfied:value=>{ rule.条件=value; return builder; },
      setBackground:value=>{ rule.背景=value; return builder; },
      setRanges:value=>{ rule.ranges=value; return builder; },
      setFontColor:value=>{ rule.文字色=value; return builder; },
      setBold:value=>{ rule.太字=value; return builder; },
      build:()=>rule
    };
    return builder;
  };
  const sheet={
    getRange:(row,col,rowCount,colCount)=>{
      const range={row,col,rowCount,colCount}; calls.ranges.push(range); return range;
    },
    setConditionalFormatRules:rules=>calls.saved.push(rules)
  };

  context.取り置き_棚確認書式を設定_(sheet,3);

  assert.strictEqual(JSON.stringify(calls.ranges),JSON.stringify([{row:2,col:1,rowCount:3,colCount:17}]));
  assert.strictEqual(calls.saved.length,1,'5ルールを1回で保存');
  assert.strictEqual(calls.saved[0].length,5);
  assert.ok(calls.saved[0].every(rule=>rule.ranges[0]===calls.ranges[0]),'全ルールが商品行だけを対象');
  assert.strictEqual(calls.saved[0][2].文字色,'#990000','出荷済みだけ濃い赤文字');
  assert.strictEqual(calls.saved[0][2].太字,true,'出荷済みだけ太字');
  assert.strictEqual(calls.saved[0].filter(rule=>rule.文字色).length,1);
  assert.strictEqual(calls.saved[0].filter(rule=>rule.太字).length,1);
});

test('棚確認の条件付き書式は0件なら古いルールを空配列で消す', () => {
  const calls={getRange:0,saved:[]};
  const sheet={
    getRange:()=>{ calls.getRange++; throw new Error('0件では範囲を作らない'); },
    setConditionalFormatRules:rules=>calls.saved.push(rules)
  };

  context.取り置き_棚確認書式を設定_(sheet,0);

  assert.strictEqual(calls.getRange,0);
  assert.strictEqual(JSON.stringify(calls.saved),JSON.stringify([[]]));
});

test('取り置き登録の書式更新は候補縮小時に全管理領域を消してから現在行だけへ再設定する', () => {
  const calls=[];
  context.SpreadsheetApp.BorderStyle={SOLID:'SOLID',SOLID_THICK:'SOLID_THICK'};
  context.SpreadsheetApp.newConditionalFormatRule=()=>{
    const rule={};
    const builder={
      whenFormulaSatisfied:value=>{ rule.条件=value; return builder; },
      setBackground:value=>{ rule.背景=value; return builder; },
      setRanges:value=>{ rule.ranges=value; return builder; },
      setFontColor:value=>{ rule.文字色=value; return builder; },
      setBold:value=>{ rule.太字=value; return builder; },
      build:()=>rule
    };
    return builder;
  };
  const sheet={
    getMaxRows:()=>8,
    getLastRow:()=>3,
    getLastColumn:()=>17,
    getRange:(row,col,rowCount,colCount)=>{
      const spec={row,col,rowCount,colCount};
      const range={
        setBackground:value=>{ calls.push({type:'background',spec,value}); return range; },
        setBorder:(...args)=>{ calls.push({type:args.every(value=>value===false)?'clearBorder':'gridBorder',spec,args}); return range; },
        getValues:()=>[['100'],['200']],
        setBackgrounds:values=>{ calls.push({type:'backgrounds',spec,values}); return range; }
      };
      return range;
    },
    getRangeList:a1=>({setBorder:(...args)=>calls.push({type:'thickBorder',a1,args})}),
    setConditionalFormatRules:rules=>calls.push({type:'conditionalRules',rules})
  };
  const candidates=[
    {受注番号:'100',判定:'要棚確認'},
    {受注番号:'200',判定:''},
    {受注番号:'200',判定:'即納'}
  ];

  context.取り置き_登録行書式を更新_(sheet,candidates);

  const clearBackground=calls.find(call=>call.type==='background' && call.value===null);
  const clearBorder=calls.find(call=>call.type==='clearBorder');
  const gridBorder=calls.find(call=>call.type==='gridBorder');
  assert.strictEqual(JSON.stringify(clearBackground.spec),JSON.stringify({row:2,col:1,rowCount:7,colCount:17}),'旧候補数ではなくmaxRowsまで背景を消す');
  assert.strictEqual(JSON.stringify(clearBorder.spec),JSON.stringify({row:2,col:1,rowCount:7,colCount:17}),'全管理領域の格子・太線を一括で消す');
  assert.ok(calls.indexOf(clearBackground)<calls.indexOf(gridBorder),'現在候補の格子より先に旧背景を消す');
  assert.ok(calls.indexOf(clearBorder)<calls.indexOf(gridBorder),'現在候補の格子より先に旧罫線を消す');
  const backgrounds=calls.find(call=>call.type==='backgrounds');
  assert.strictEqual(JSON.stringify(backgrounds.spec),JSON.stringify({row:2,col:1,rowCount:3,colCount:17}));
  assert.strictEqual(backgrounds.values[0][0],'#fff2cc');
  assert.strictEqual(backgrounds.values[1][0],null);
  assert.strictEqual(backgrounds.values[2][0],'#cfe2f3','即納の表示専用行は受注明細と同じ水色');
  assert.strictEqual(calls.find(call=>call.type==='conditionalRules').rules.length,5);
});

test('取り置き登録の書式更新は0件でも全管理領域の背景と罫線を一括で消す', () => {
  const calls=[];
  const sheet={
    getMaxRows:()=>8,
    getLastRow:()=>1,
    getLastColumn:()=>17,
    getRange:(row,col,rowCount,colCount)=>{
      const spec={row,col,rowCount,colCount};
      const range={
        setBackground:value=>{ calls.push({type:'background',spec,value}); return range; },
        setBorder:(...args)=>{ calls.push({type:'border',spec,args}); return range; },
        setBackgrounds:()=>{ throw new Error('0件では現在背景を設定しない'); }
      };
      return range;
    },
    getRangeList:()=>{ throw new Error('0件では太線を設定しない'); },
    setConditionalFormatRules:rules=>calls.push({type:'conditionalRules',rules})
  };

  context.取り置き_登録行書式を更新_(sheet,[]);

  assert.strictEqual(JSON.stringify(calls[0]),JSON.stringify({type:'background',spec:{row:2,col:1,rowCount:7,colCount:17},value:null}));
  assert.strictEqual(JSON.stringify(calls[1]),JSON.stringify({type:'border',spec:{row:2,col:1,rowCount:7,colCount:17},args:[false,false,false,false,false,false]}));
  assert.strictEqual(JSON.stringify(calls[2]),JSON.stringify({type:'conditionalRules',rules:[]}));
  assert.strictEqual(calls.length,3,'背景・罫線・条件付き書式以外は変更しない');
});

// ===== 2026-07-18 オンライン改修分: 棚戻し待ち・台帳確保の重ね合わせ・一括反映 =====

test('メモ追記は手書きメモを消さず、同じ理由は二重に足さない', () => {
  assert.strictEqual(context.取り置き_メモ追記_('','CSV注文キャンセル'),'CSV注文キャンセル');
  assert.strictEqual(context.取り置き_メモ追記_('棚A-3','CSV注文キャンセル'),'棚A-3 / CSV注文キャンセル');
  assert.strictEqual(context.取り置き_メモ追記_('棚A-3 / CSV注文キャンセル','CSV注文キャンセル'),'棚A-3 / CSV注文キャンセル');
  assert.strictEqual(context.取り置き_メモ追記_('棚A-3',''),'棚A-3');
});

test('台帳確保集計は取り置き中のEMS引当等だけを数え、開始前在庫と終了行を除く', () => {
  const summary=context.取り置き_台帳確保集計_([
    {状態:'取り置き中',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:2,取置元種別:'EMS'},
    {状態:'取り置き中',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:1,取置元種別:'キャンセル再引当'},
    {状態:'取り置き中',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:5,取置元種別:'開始前在庫'}, // 登録シート自身の確定分は含めない
    {状態:'発送済み',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:4,取置元種別:'EMS'},
    {状態:'キャンセル戻し',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:3,取置元種別:'EMS'}
  ]);
  const key=context.取り置き_行キー_({受注番号:'101',商品コード:'AAA',SKU:'AAAb'});
  assert.strictEqual(summary[key],3,'EMS2+再引当1のみ');
});

test('台帳確保の全数確保は入力欄を空へ戻し、一部確保は残り個数を出して入力を残す', () => {
  const secured=context.取り置き_台帳確保集計_([
    {状態:'取り置き中',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:2,取置元種別:'EMS'},
    {状態:'取り置き中',受注番号:'102',商品コード:'BBB',SKU:'BBBb',取り置き数量:1,取置元種別:'キャンセル再引当'}
  ]);
  const out=context.取り置き_台帳確保を適用_([
    {取置ID:'A',受注番号:'101',商品コード:'AAA',SKU:'AAAb',注文数量:2,現物取り置き数量:2},  // 全数確保
    {取置ID:'B',受注番号:'102',商品コード:'BBB',SKU:'BBBb',注文数量:3,現物取り置き数量:1},  // 一部確保
    {取置ID:'C',受注番号:'103',商品コード:'CCC',SKU:'CCCb',注文数量:1,現物取り置き数量:1}   // 確保なし
  ],secured);
  assert.strictEqual(out[0].現物取り置き数量,'','EMS確保分を開始前在庫として二重登録させない');
  assert.strictEqual(out[0].台帳確保,'台帳確保済み2個');
  assert.strictEqual(out[0].台帳確保数,2);
  assert.strictEqual(out[1].現物取り置き数量,1,'残り分の入力はそのまま');
  assert.strictEqual(out[1].台帳確保,'台帳確保済み1個／残り2個要確認');
  assert.strictEqual(out[2].台帳確保,'');
  assert.strictEqual(out[2].台帳確保数,0);
});

test('棚確認判定は台帳の全数確保を対象外にし、一部確保は旧情報なしでも要棚確認に残す', () => {
  assert.strictEqual(context.取り置き_棚確認判定_({注文数量:2,台帳確保数:2,旧入荷日:'2026-07-01',現物取り置き数量:'',棚確認:''}),'','全数確保は④の管理下');
  assert.strictEqual(context.取り置き_棚確認判定_({注文数量:3,台帳確保数:1,旧入荷日:'',現物取り置き数量:'',棚確認:''}),'要棚確認','一部確保は残りの現物を確かめる');
  assert.strictEqual(context.取り置き_棚確認判定_({注文数量:2,台帳確保数:0,旧入荷日:'2026-07-01',現物取り置き数量:'',棚確認:''}),'要棚確認');
  assert.strictEqual(context.取り置き_棚確認判定_({注文数量:2,台帳確保数:0,旧入荷日:'',現物取り置き数量:'',棚確認:''}),'');
  assert.strictEqual(context.取り置き_棚確認判定_({注文数量:2,台帳確保数:1,旧入荷日:'',現物取り置き数量:1,棚確認:''}),'','数量入力で解決扱い');
});

test('棚戻し候補は未確認のキャンセル戻しだけを赤い要対応行にし、前回入力を取置IDで引き継ぐ', () => {
  const ledger=[
    {取置ID:'RTN-1',状態:'キャンセル戻し',戻し処理結果:'未確認',受注番号:'201',商品コード:'BBB',SKU:'BBBb',取り置き数量:2,元EMS番号:'EG1','終了理由・メモ':'CSV注文キャンセル（棚戻し待ち）'},
    {取置ID:'RTN-2',状態:'キャンセル戻し',戻し処理結果:'現物あり',受注番号:'202',商品コード:'CCC',SKU:'CCCb',取り置き数量:1},
    {取置ID:'ACT-1',状態:'取り置き中',受注番号:'203',商品コード:'DDD',SKU:'DDDb',取り置き数量:1}
  ];
  const sheetRows=[{取置ID:'RTN-1',氏名:'山田',メモ:'棚B-2',処理:'棚へ戻した'}];
  const out=context.取り置き_棚戻し候補_(ledger,sheetRows);
  assert.strictEqual(out.length,1,'処理済み・取り置き中の行は出さない');
  assert.strictEqual(out[0].取置ID,'RTN-1');
  assert.strictEqual(out[0].判定,'棚戻し待ち');
  assert.strictEqual(out[0].要対応,'棚へ戻す 2個');
  assert.strictEqual(out[0].台帳確保,'確保済み2個');
  assert.strictEqual(out[0].旧EMS,'EG1');
  assert.strictEqual(out[0].処理,'棚へ戻した','洗い替えても選択中の処理は消えない');
  assert.strictEqual(out[0].氏名,'山田');
  assert.strictEqual(out[0].メモ,'棚B-2','手書きメモを引き継ぐ');
  assert.strictEqual(out[0].受注ステータス,'CSV注文キャンセル（棚戻し待ち）','理由を表示する');
});

test('統合反映は通常の数量確定と棚戻し処理を一括で検証して反映する', () => {
  const now='2026-07-19T10:00:00';
  const ledger=[
    {取置ID:'RTN-1',状態:'キャンセル戻し',戻し処理結果:'未確認',受注番号:'201',商品コード:'BBB',SKU:'BBBb',取り置き数量:2,取置元種別:'EMS',元EMS番号:'EG1'}
  ];
  const inputs=[
    {取置ID:'INIT|101|AAA|AAAB',受注番号:'101',商品コード:'AAA',SKU:'AAAb',注文数量:3,現物取り置き数量:2,棚確認:'',メモ:'棚A'},
    {取置ID:'RTN-1',受注番号:'201',商品コード:'BBB',SKU:'BBBb',注文数量:2,現物取り置き数量:'',処理:'棚へ戻した'}
  ];
  const plan=context.取り置き_統合反映計画_(inputs,ledger,now);
  assert.strictEqual(plan.errors.length,0);
  assert.strictEqual(plan.counts.取り置き行,1);
  assert.strictEqual(plan.counts.取り置き数量,2);
  assert.strictEqual(plan.counts.棚へ戻した,1);
  assert.strictEqual(plan.counts.現物なし,0);
  const rtn=plan.rows.find(r=>r.取置ID==='RTN-1');
  assert.strictEqual(rtn.戻し処理結果,'現物あり','棚へ戻した=現物ありとして確定');
  const init=plan.rows.find(r=>r.取置ID==='INIT|101|AAA|AAAB');
  assert.strictEqual(init.状態,'取り置き中');
  assert.strictEqual(init.取り置き数量,2);
});

test('統合反映はどちらか一方のエラーでも台帳へ部分反映しない', () => {
  const now='2026-07-19T10:00:00';
  const ledger=[
    {取置ID:'RTN-1',状態:'キャンセル戻し',戻し処理結果:'未確認',受注番号:'201',商品コード:'BBB',SKU:'BBBb',取り置き数量:2,取置元種別:'EMS',元EMS番号:'EG1'}
  ];
  const inputs=[
    {取置ID:'INIT|101|AAA|AAAB',受注番号:'101',商品コード:'AAA',SKU:'AAAb',注文数量:1,現物取り置き数量:5,棚確認:'',メモ:''}, // 注文超過エラー
    {取置ID:'RTN-1',受注番号:'201',商品コード:'BBB',SKU:'BBBb',注文数量:2,現物取り置き数量:'',処理:'棚へ戻した'}          // こちらは正しい入力
  ];
  const plan=context.取り置き_統合反映計画_(inputs,ledger,now);
  assert.ok(plan.errors.length>0);
  assert.strictEqual(plan.rows.length,0,'正しい棚戻し側も反映しない(全体で止める)');
});

// ===== 2026-07-20 全行表示化: 即納行は表示専用で注文グループに付く =====

test('即納行は取り寄せ候補がある注文にだけ表示専用で付き、メモを引き継ぐ', () => {
  const candidates=[
    {取置ID:'INIT|101|AAA|AAAB',受注番号:'101',氏名:'鶴原',商品コード:'AAA',SKU:'AAAb',注文数量:2,判定:''}
  ];
  const sokuno=[
    {ban:'101',氏名:'鶴原',code:'BBB',sku:'BBBa',qty:1,ステータス:'部分包装'},
    {ban:'101',氏名:'鶴原',code:'BBB',sku:'BBBa',qty:1,ステータス:'部分包装'},
    {ban:'999',氏名:'別注文',code:'ZZZ',sku:'ZZZa',qty:1,ステータス:''}
  ];
  const out=context.取り置き_即納行を付与_(candidates,sokuno,[
    {取置ID:'即納|101|BBB|BBBA',メモ:'レジ横に置いた'}
  ]);
  assert.strictEqual(out.length,2,'取り寄せ候補が無い注文999の即納行は付けない');
  const view=out[1];
  assert.strictEqual(view.取置ID,'即納|101|BBB|BBBA');
  assert.strictEqual(view.判定,'即納');
  assert.strictEqual(view.現在の状態,'即納');
  assert.strictEqual(view.注文数量,2,'分割行は合算');
  assert.strictEqual(view.受注ステータス,'部分包装');
  assert.strictEqual(view.現物取り置き数量,'','数量は入力させない(表示専用)');
  assert.strictEqual(view.棚確認,'');
  assert.strictEqual(view.台帳確保,'');
  assert.strictEqual(view.メモ,'レジ横に置いた','洗い替えてもメモは引き継ぐ');
});

test('反映は即納行の数量入力をエラーで止め、空なら台帳へ登録しない', () => {
  const now='2026-07-20T10:00:00';
  const ok=context.取り置き_統合反映計画_([
    {取置ID:'即納|101|BBB|BBBA',受注番号:'101',商品コード:'BBB',SKU:'BBBa',注文数量:2,現物取り置き数量:'',棚確認:'',メモ:''},
    {取置ID:'INIT|101|AAA|AAAB',受注番号:'101',商品コード:'AAA',SKU:'AAAb',注文数量:3,現物取り置き数量:2,棚確認:'',メモ:''}
  ],[],now);
  assert.strictEqual(ok.errors.length,0,'空の即納行はエラーにしない');
  assert.strictEqual(ok.rows.length,1,'即納行は台帳へ登録しない(全件再計算の開始前在庫を汚さない)');
  assert.strictEqual(ok.rows[0].取置ID,'INIT|101|AAA|AAAB');
  assert.strictEqual(ok.counts.取り置き行,1);

  const bad=context.取り置き_統合反映計画_([
    {取置ID:'即納|101|BBB|BBBA',受注番号:'101',商品コード:'BBB',SKU:'BBBa',注文数量:2,現物取り置き数量:1,棚確認:'',メモ:''}
  ],[],now);
  assert.ok(bad.errors.length>0,'即納行への数量入力は開始前在庫の誤登録なので止める');
  assert.ok(bad.errors.join('\n').indexOf('即納')>=0,'エラー文で即納行だと分かるようにする');
  assert.strictEqual(bad.rows.length,0);
});

// ===== 2026-07-20 全行表示化v2: 明細区分の振り分け・台湾中国(別ルート)の表示専用 =====

test('取り置き_明細区分_は台湾中国を最優先で別ルート、指定なしは即納扱いにする', () => {
  assert.strictEqual(context.取り置き_明細区分_('取り寄せ',false),'取り寄せ');
  assert.strictEqual(context.取り置き_明細区分_('即納',false),'即納');
  assert.strictEqual(context.取り置き_明細区分_('指定なし',false),'即納','指定なしはほぼ即納なので即納扱い');
  assert.strictEqual(context.取り置き_明細区分_('取り寄せ',true),'別ルート','台湾中国の取り寄せは別ルート優先');
  assert.strictEqual(context.取り置き_明細区分_('指定なし',true),'別ルート','区分より別ルートを優先');
});

test('別ルート行はアクティブ注文だけに橙・表示専用で付き、入荷日と要対応を出す', () => {
  const candidates=[
    {取置ID:'INIT|201|AAA|AAAB',受注番号:'201',氏名:'丸山',商品コード:'AAA',SKU:'AAAb',注文数量:1,判定:''}
  ];
  const betsu=[
    {ban:'201',氏名:'丸山',code:'TW01',sku:'TW01',qty:1,ステータス:'部分包装',入荷日:''},                 // 到着だが入荷日未入力→要対応
    {ban:'201',氏名:'丸山',code:'CN02',sku:'CN02',qty:2,ステータス:'出荷待/取寄せ',入荷日:'2026-07-15'},   // 入荷日あり=確保済み
    {ban:'888',氏名:'別注文',code:'ZZZ',sku:'ZZZ',qty:1,ステータス:'',入荷日:''}                          // 取り寄せ候補が無い→付けない
  ];
  const out=context.取り置き_別ルート行を付与_(candidates,betsu,[
    {取置ID:'別ルート|201|TW01|TW01',メモ:'棚C'}
  ]);
  assert.strictEqual(out.length,3,'アクティブ注文201の台湾中国2行だけ付く(888は付けない)');
  const tw=out.find(r=>r.取置ID==='別ルート|201|TW01|TW01');
  const cn=out.find(r=>r.取置ID==='別ルート|201|CN02|CN02');
  assert.ok(tw && cn,'2行とも別ルートIDで付く');
  assert.strictEqual(tw.判定,'別ルート');
  assert.strictEqual(tw.現在の状態,'別ルート');
  assert.strictEqual(tw.現物取り置き数量,'','別ルートは数量入力させない(表示専用)');
  assert.strictEqual(tw.棚確認,'');
  assert.strictEqual(tw.旧入荷日,'','入荷日未入力なら空');
  assert.ok(String(tw.要対応).indexOf('入荷日')>=0,'入荷日未入力は受注明細への入荷日入力を促す');
  assert.strictEqual(tw.メモ,'棚C','洗い替えてもメモは引き継ぐ');
  assert.strictEqual(cn.旧入荷日,'2026-07-15','入荷日ありは確保状態として表示');
  assert.strictEqual(cn.要対応,'','入荷日ありは要対応なし(確保済み)');
  assert.strictEqual(cn.注文数量,2);
});

test('反映は別ルート行の数量入力をエラーで止め、空なら台帳へ登録しない', () => {
  const now='2026-07-20T10:00:00';
  const ok=context.取り置き_統合反映計画_([
    {取置ID:'別ルート|201|TWCODE|TWCODE',受注番号:'201',商品コード:'TWCODE',SKU:'TWCODEa',注文数量:1,現物取り置き数量:'',棚確認:'',メモ:''},
    {取置ID:'INIT|201|AAA|AAAB',受注番号:'201',商品コード:'AAA',SKU:'AAAb',注文数量:1,現物取り置き数量:1,棚確認:'',メモ:''}
  ],[],now);
  assert.strictEqual(ok.errors.length,0,'空の別ルート行はエラーにしない');
  assert.strictEqual(ok.rows.length,1,'別ルート行は台帳へ登録しない(別ルート済数量と二重にしない)');
  assert.strictEqual(ok.rows[0].取置ID,'INIT|201|AAA|AAAB');

  const bad=context.取り置き_統合反映計画_([
    {取置ID:'別ルート|201|TWCODE|TWCODE',受注番号:'201',商品コード:'TWCODE',SKU:'TWCODEa',注文数量:1,現物取り置き数量:1,棚確認:'',メモ:''}
  ],[],now);
  assert.ok(bad.errors.length>0,'別ルート行への数量入力は止める(確保は受注明細の入荷日)');
  assert.ok(bad.errors.join('\n').indexOf('入荷日')>=0,'エラー文で入荷日入力を案内する');
  assert.strictEqual(bad.rows.length,0);
});

test('CSV数量減の超過分は新しい登録から棚戻し待ちへ分離し、行の途中なら分割する', () => {
  const now='2026-07-19T10:00:00';
  const ledger=[
    {取置ID:'OLD',状態:'取り置き中',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:2,取置元種別:'開始前在庫',登録日時:'2026-07-01 10:00:00','終了理由・メモ':''},
    {取置ID:'NEW',状態:'取り置き中',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:3,取置元種別:'EMS',元EMS番号:'EG1',登録日時:'2026-07-10 10:00:00','終了理由・メモ':''}
  ];
  // 注文2に対し確保5 → 超過3。新しいNEW(3個)がまるごと棚戻し待ちへ、古い確定は守る
  const whole=context.取り置き_CSV遷移計画_([
    {受注番号:'101',受注ステータス:'出荷待/取寄せ',商品コード:'AAA',SKU:'AAAb',個数:2}
  ],ledger,now);
  assert.strictEqual(whole.errors.length,0);
  const wholeNew=whole.rows.find(r=>r.取置ID==='NEW');
  assert.strictEqual(wholeNew.状態,'キャンセル戻し');
  assert.strictEqual(wholeNew.戻し処理結果,'未確認');
  assert.strictEqual(whole.rows.find(r=>r.取置ID==='OLD').状態,'取り置き中','古い確定を守る');
  assert.strictEqual(whole.counts.棚戻し待ち,1);
  assert.strictEqual(whole.counts.棚戻し数量,3);
  // 注文3に対し確保5 → 超過2。NEWは1個残して2個だけ分割して棚戻し待ちへ
  const split=context.取り置き_CSV遷移計画_([
    {受注番号:'101',受注ステータス:'出荷待/取寄せ',商品コード:'AAA',SKU:'AAAb',個数:3}
  ],ledger,now);
  const splitNew=split.rows.find(r=>r.取置ID==='NEW');
  assert.strictEqual(splitNew.状態,'取り置き中');
  assert.strictEqual(splitNew.取り置き数量,1);
  const rtn=split.rows.find(r=>r.取置ID==='NEW|RTN');
  assert.ok(rtn,'分割行はNEW|RTNのIDで追加');
  assert.strictEqual(rtn.状態,'キャンセル戻し');
  assert.strictEqual(rtn.取り置き数量,2);
  assert.ok(String(rtn['終了理由・メモ']).indexOf('CSV数量減')>=0);
  assert.strictEqual(split.counts.棚戻し数量,2);
});

test('ymd_は旧バグが書いた「シリアル-01-01」を元の日付へ復元する', () => {
  assert.strictEqual(context.ymd_('46213-01-01'),'2026-07-10');
  assert.strictEqual(context.ymd_('2026-01-01'),'2026-01-01','通常の元日はそのまま');
});

test('引当切替差分は同一キーの複数行を相殺してから更新・追加・削除にする', () => {
  const planned=[
    {ban:'101',code:'AAA',sku:'AAAb',qty:1,state:'引当待ち',ems:''},
    {ban:'101',code:'AAA',sku:'AAAb',qty:2,state:'部分在庫',ems:''}
  ];
  const current=[
    {ban:'101',code:'AAA',sku:'AAAb',qty:1,state:'引当待ち',ems:''}
  ];
  const out=context.引当切替差分_純計算_(planned,current);
  assert.strictEqual(out.length,1,'同一内容の行は相殺され、残る1行だけが差分');
  assert.strictEqual(out[0].change,'追加');
  assert.strictEqual(out[0].after.qty,2);
  // 注文番号在庫などで表示コードが切り替わった行はmatchCodeで同一視して「更新」にする
  const renamed=context.引当切替差分_純計算_(
    [{ban:'102',code:'ORDER102',matchCode:'BBB',sku:'BBBb',qty:1,state:'引当待ち',ems:''}],
    [{ban:'102',code:'BBB',sku:'BBBb',qty:1,state:'引当待ち',ems:''}]);
  assert.strictEqual(renamed.length,1);
  assert.strictEqual(renamed[0].change,'更新','表示コード切替は追加+削除にしない');
});

// ===== 2026-07-20 孤児取り置き一括解除ツール =====

test('孤児取り置き一覧は「取り置き中だが受注明細に無い」行だけ拾い、出荷形跡を付ける', () => {
  const ledger=[
    {取置ID:'INIT|101|AAA|AAAB',状態:'取り置き中',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:2}, // 注文に有る→孤児でない
    {取置ID:'INIT|900|SUB|SUBB',状態:'取り置き中',受注番号:'900',商品コード:'SUB',SKU:'SUBb',取り置き数量:1},  // 注文に無い＋出荷形跡→孤児(チェック候補)
    {取置ID:'INIT|901|OLD|OLDB',状態:'取り置き中',受注番号:'901',商品コード:'OLD',SKU:'OLDb',取り置き数量:1},  // 注文に無い＋出荷形跡なし→孤児(オフ)
    {取置ID:'INIT|902|SHP|SHPB',状態:'発送済み',受注番号:'902',商品コード:'SHP',SKU:'SHPb',取り置き数量:1},    // 取り置き中でない→無視
    {取置ID:'INIT|903|REL|RELB',状態:'手動解除',受注番号:'903',商品コード:'REL',SKU:'RELb',取り置き数量:1}     // 既に解除→無視
  ];
  const 注文キー集合=new Set([context.取り置き_行キー_({受注番号:'101',商品コード:'AAA',SKU:'AAAb'})]);
  const 出荷済み受注番号集合=new Set(['900']);
  const out=context.取り置き_孤児取り置き一覧_(ledger,注文キー集合,出荷済み受注番号集合);
  assert.strictEqual(JSON.stringify(out.map(r=>r.取置ID)),JSON.stringify(['INIT|900|SUB|SUBB','INIT|901|OLD|OLDB']),'孤児2件だけ(発送済み・手動解除・注文有りは除外)');
  assert.strictEqual(out[0].出荷形跡,true,'受注番号が消込台帳にあれば形跡あり(既定チェック)');
  assert.strictEqual(out[1].出荷形跡,false,'消込台帳に無ければ形跡なし');
  assert.strictEqual(out[0].取り置き数量,1);
  assert.strictEqual(out[0].商品コード,'SUB');
});

test('一括解除計画は指定IDの取り置き中だけ手動解除にし、他は触らない(冪等)', () => {
  const now='2026-07-20T16:00:00';
  const ledger=[
    {取置ID:'A',状態:'取り置き中',受注番号:'900',商品コード:'SUB',取り置き数量:1,'終了理由・メモ':''},
    {取置ID:'B',状態:'取り置き中',受注番号:'901',商品コード:'OLD',取り置き数量:1,'終了理由・メモ':'棚C-1'},
    {取置ID:'C',状態:'発送済み',受注番号:'902',商品コード:'SHP',取り置き数量:1,'終了理由・メモ':''}
  ];
  const out=context.取り置き_一括解除計画_(ledger,new Set(['A','C']),'品切れ・代替品で発送',now);
  const a=out.find(r=>r.取置ID==='A'), b=out.find(r=>r.取置ID==='B'), c=out.find(r=>r.取置ID==='C');
  assert.strictEqual(a.状態,'手動解除','指定した取り置き中は解除');
  assert.strictEqual(a['終了理由・メモ'],'品切れ・代替品で発送');
  assert.strictEqual(a.更新日時,now);
  assert.strictEqual(b.状態,'取り置き中','非指定は不変');
  assert.strictEqual(b['終了理由・メモ'],'棚C-1');
  assert.strictEqual(c.状態,'発送済み','指定でも取り置き中でなければ触らない(冪等)');
});

process.exitCode = failures ? 1 : 0;
console.log(failures ? `FAILURES: ${failures}` : 'ALL TESTS PASSED');
