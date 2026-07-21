// 引当全件再構築v3（取り置き台帳統合版）の純粋ロジックテスト
// 実行: GASアップスクリプト直下で node tests/project24_full_allocation_rebuild.test.js
const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

const context = {
  console, Date, Set, Map, JSON, Math,
  normCode_: value => String(value == null ? '' : value).trim().toUpperCase().replace(/_/g, '-'),
  P列指定文字列_: () => ''
};
vm.createContext(context);
vm.runInContext(fs.readFileSync('Project_24/取り置き計算.js', 'utf8'), context);
if (fs.existsSync('Project_24/全件再計算.js')) {
  vm.runInContext(fs.readFileSync('Project_24/全件再計算.js', 'utf8'), context);
}
vm.runInContext(fs.readFileSync('Project_24/P列自動記入.js', 'utf8'), context);

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
const json = value => JSON.parse(JSON.stringify(value));

test('SKU正規化: Yahoo/GoQの在庫枝番だけ外しEMS末尾文字と数値枝番を守る', () => {
  assert.strictEqual(context.全件再計算_SKU正規化_('YMNGD08-1b', 'GoQ'), 'YMNGD08-1');
  assert.strictEqual(context.全件再計算_SKU正規化_('YMNGD08-2a', 'Yahoo'), 'YMNGD08-2');
  assert.strictEqual(context.全件再計算_SKU正規化_('EBS1504B', 'EMS'), 'EBS1504B');
  assert.strictEqual(context.全件再計算_SKU正規化_('YMNGD08-2', 'EMS'), 'YMNGD08-2');
});

test('実EMS行: 棚卸合成行・EMS番号なしを除外し実EMSだけを残す', () => {
  const base = {code: 'YMNGD08-1', qty: 2, arrival: '2026-06-22', row: 10, status: '到着済'};
  assert.strictEqual(context.全件再計算_実EMS行_(Object.assign({ems: '棚卸20260710'}, base)), null);
  assert.strictEqual(context.全件再計算_実EMS行_(Object.assign({ems: ''}, base)), null);
  const real = context.全件再計算_実EMS行_(Object.assign({ems: 'EG049624664KR'}, base));
  assert.strictEqual(real.ems, 'EG049624664KR');
  assert.strictEqual(real.sku, 'YMNGD08-1');
  assert.strictEqual(real.qty, 2);
});

test('GoQ最新化: 受注番号+商品IDの最新だけを採用し最新数量0は消費しない', () => {
  const reduced = context.全件再計算_発送最新化_([
    {ban: '101', itemId: 'A', importedAt: '2026-07-01 10:00:00', qty: 1, status: '処理済み', shipDate: '2026-07-01'},
    {ban: '101', itemId: 'A', importedAt: '2026-07-02 10:00:00', qty: 0, status: '処理済み', shipDate: '2026-07-01'},
    {ban: '101', itemId: 'B', importedAt: '2026-07-01 10:00:00', qty: 2, status: '処理済み', shipDate: '2026-07-01'}
  ]);
  assert.deepStrictEqual(json(reduced.rows.map(r => [r.ban, r.itemId, r.qty])), [['101', 'B', 2]]);
  assert.strictEqual(reduced.issues.length, 0);
});

test('GoQ最新化: 同じ最新時刻で内容が競合したキーは要確認にする', () => {
  const reduced = context.全件再計算_発送最新化_([
    {ban: '101', itemId: 'A', importedAt: '2026-07-02 10:00:00', qty: 1, status: '処理済み', shipDate: '2026-07-01'},
    {ban: '101', itemId: 'A', importedAt: '2026-07-02 10:00:00', qty: 2, status: '処理済み', shipDate: '2026-07-01'}
  ]);
  assert.strictEqual(reduced.rows.length, 0);
  assert.strictEqual(reduced.issues[0].type, '発送最新競合');
});

test('再構築: 発送日より後のEMSを発送済みに使わずSKUをブロックする', () => {
  const result = context.全件再計算_再構築_({
    supplies: [{ems: 'EMS-LATE', code: 'TEST-01', qty: 1, arrival: '2026-07-10', row: 2, status: '到着済'}],
    shipped: [{ban: '101', itemId: 'A', sku: 'TEST-01b', code: 'TEST-01', qty: 1, shipDate: '2026-07-01', orderDate: '2026-06-20'}],
    currentOrders: [], yahooA: {'TEST-01': 0}
  });
  assert.ok(result.blockedSkus.indexOf('TEST-01') >= 0);
  assert.strictEqual(result.ledgerRows.filter(r => r.状態 === '発送済み').length, 0);
  assert.ok(result.issues.some(r => r.type === '発送供給不足'));
});

test('再構築: GoQ発送済みの即納aはEMSを消費せず取寄せbの供給を守る', () => {
  const result = context.全件再計算_再構築_({
    supplies: [{ems:'EMS-A',code:'TEST-A',qty:1,arrival:'2026-07-01',row:2,status:'到着済'}],
    shipped: [{ban:'A01',itemId:'A',sku:'TEST-Aa',code:'TEST-A',qty:1,shipDate:'2026-07-02',orderDate:'2026-07-01'}],
    currentOrders: [{ban:'B01',sku:'TEST-Ab',code:'TEST-A',qty:1,orderDate:'2026-07-03',row:3,route:'韓国取寄せ'}],
    yahooA: {'TEST-A':0}
  });
  assert.strictEqual(result.blockedSkus.length, 0);
  assert.strictEqual(result.ledgerRows.filter(r => r.状態 === '発送済み').length, 0);
  assert.strictEqual(result.ledgerRows.filter(r => r.状態 === '取り置き中' && r.受注番号 === 'B01').reduce((n,r)=>n+r.取り置き数量,0), 1);
});

test('再構築: EMS商品コード末尾の受注番号タグをFIFOより優先する', () => {
  const result = context.全件再計算_再構築_({
    supplies: [{ems:'EMS-D',code:'TEST-D（1010002）',qty:1,arrival:'2026-07-01',row:2,status:'到着済'}],
    shipped: [],
    currentOrders: [
      {ban:'1010001',sku:'TEST-Db',code:'TEST-D',qty:1,orderDate:'2026-06-01',row:3,route:'韓国取寄せ'},
      {ban:'1010002',sku:'TEST-Db',code:'TEST-D',qty:1,orderDate:'2026-06-02',row:4,route:'韓国取寄せ'}
    ],
    yahooA: {'TEST-D':0}
  });
  const active = result.ledgerRows.filter(r => r.状態 === '取り置き中');
  assert.strictEqual(active.length, 1);
  assert.strictEqual(active[0].受注番号, '1010002');
  assert.strictEqual(active[0].元EMS商品コード, 'TEST-D（1010002）');
});

test('再構築: Yahoo自由在庫を先に確保して残りだけを取寄せへ割り当てる', () => {
  const result = context.全件再計算_再構築_({
    supplies: [{ems: 'EMS-1', code: 'TEST-02', qty: 3, arrival: '2026-07-01', row: 2, status: '到着済'}],
    shipped: [],
    currentOrders: [{ban: '200', sku: 'TEST-02b', code: 'TEST-02', qty: 1, orderDate: '2026-07-02', row: 10, route: '韓国取寄せ'}],
    yahooA: {'TEST-02': 2}
  });
  assert.strictEqual(result.blockedSkus.length, 0);
  assert.strictEqual(result.movementRows.filter(r => r.移動先 === 'Yahoo自由在庫').reduce((n, r) => n + r.数量, 0), 2);
  assert.strictEqual(result.ledgerRows.filter(r => r.状態 === '取り置き中' && r.受注番号 === '200').reduce((n, r) => n + r.取り置き数量, 0), 1);
});

test('再構築: 現在取寄せは注文日時・受注番号・行の古い順にする', () => {
  const result = context.全件再計算_再構築_({
    supplies: [{ems: 'EMS-1', code: 'TEST-03', qty: 1, arrival: '2026-07-01', row: 2, status: '到着済'}],
    shipped: [],
    currentOrders: [
      {ban: '301', sku: 'TEST-03b', code: 'TEST-03', qty: 1, orderDate: '2026-07-03', row: 20, route: '韓国取寄せ'},
      {ban: '300', sku: 'TEST-03b', code: 'TEST-03', qty: 1, orderDate: '2026-07-02', row: 30, route: '韓国取寄せ'}
    ],
    yahooA: {'TEST-03': 0}
  });
  const active = result.ledgerRows.filter(r => r.状態 === '取り置き中');
  assert.strictEqual(active.length, 1);
  assert.strictEqual(active[0].受注番号, '300');
});

test('再構築: 同じ日の現在取寄せも注文時刻の古い順にする', () => {
  const result = context.全件再計算_再構築_({
    supplies: [{ems: 'EMS-1', code: 'TEST-03T', qty: 1, arrival: '2026-07-01', row: 2, status: '到着済'}],
    shipped: [],
    currentOrders: [
      {ban: '100', sku: 'TEST-03Tb', code: 'TEST-03T', qty: 1, orderDate: '2026-07-02 10:00:00', row: 20, route: '韓国取寄せ'},
      {ban: '200', sku: 'TEST-03Tb', code: 'TEST-03T', qty: 1, orderDate: '2026-07-02 09:00:00', row: 30, route: '韓国取寄せ'}
    ],
    yahooA: {'TEST-03T': 0}
  });
  const active = result.ledgerRows.filter(r => r.状態 === '取り置き中');
  assert.strictEqual(active[0].受注番号, '200');
});

test('再構築: 現在取寄せの注文日時が無効なら古い順を推測せずブロックする', () => {
  const result = context.全件再計算_再構築_({
    supplies: [{ems:'EMS-T',code:'TEST-TIME',qty:1,arrival:'2026-07-01',row:2,status:'到着済'}],
    shipped: [],
    currentOrders: [{ban:'T01',sku:'TEST-TIMEb',code:'TEST-TIME',qty:1,orderDate:'',row:3,route:'韓国取寄せ'}],
    yahooA: {'TEST-TIME':0}
  });
  assert.ok(result.blockedSkus.indexOf('TEST-TIME') >= 0);
  assert.ok(result.issues.some(r => r.type === '受注日時不正'));
  assert.strictEqual(result.ledgerRows.filter(r => r.状態 === '取り置き中').length, 0);
});

test('再構築: 現在受注の正でない整数数量は黙って無視せずブロックする', () => {
  const result = context.全件再計算_再構築_({
    supplies: [{ems:'EMS-Q',code:'TEST-QTY',qty:1,arrival:'2026-07-01',row:2,status:'到着済'}],
    shipped: [],
    currentOrders: [{ban:'Q01',sku:'TEST-QTYb',code:'TEST-QTY',qty:1.5,orderDate:'2026-07-01 10:00:00',row:3,route:'韓国取寄せ'}],
    yahooA: {'TEST-QTY':0}
  });
  assert.ok(result.blockedSkus.indexOf('TEST-QTY') >= 0);
  assert.ok(result.issues.some(r => r.type === '受注数量不正'));
});

test('再構築: 説明できない到着済みEMS余りはブロックする', () => {
  const result = context.全件再計算_再構築_({
    supplies: [{ems: 'EMS-1', code: 'TEST-04', qty: 2, arrival: '2026-07-01', row: 2, status: '到着済'}],
    shipped: [], currentOrders: [], yahooA: {'TEST-04': 0}
  });
  assert.ok(result.blockedSkus.indexOf('TEST-04') >= 0);
  assert.ok(result.issues.some(r => r.type === 'EMS説明不能余り' && r.qty === 2));
});

test('再構築: 現行の取り置き台帳・移動台帳スキーマで行を作る', () => {
  const result = context.全件再計算_再構築_({
    supplies: [{ems: 'EMS-1', code: 'TEST-05', qty: 2, arrival: '2026-07-01', row: 2, status: '到着済'}],
    shipped: [{ban: '500', itemId: 'A', sku: 'TEST-05b', code: 'TEST-05', qty: 1, shipDate: '2026-07-02', orderDate: '2026-07-01'}],
    currentOrders: [], yahooA: {'TEST-05': 1}
  });
  const ledgerHeaders = ['取置ID','状態','受注番号','商品コード','SKU','取り置き数量','取置元種別','元EMS番号','元EMS商品コード','元取置ID','登録日時','更新日時','戻し処理結果','終了理由・メモ'];
  const movementHeaders = ['処理ID','EMS番号','商品コード','数量','移動先','処理日時'];
  ledgerHeaders.forEach(h => assert.ok(Object.prototype.hasOwnProperty.call(result.ledgerRows[0], h), `台帳列不足: ${h}`));
  movementHeaders.forEach(h => assert.ok(Object.prototype.hasOwnProperty.call(result.movementRows[0], h), `移動列不足: ${h}`));
});

test('Yahoo厳密集計: aだけを自由在庫にしbは物理在庫へ入れない', () => {
  const csv = [
    'code,name,sub-code,quantity,allow-overdraft,stock-close',
    'TEST-06,商品,TEST-06a,3,,',
    'TEST-06,商品,TEST-06b,99,,'
  ].join('\n');
  const result = context.全件再計算_YahooCSV厳密集計_(csv);
  assert.strictEqual(result.error, '');
  assert.strictEqual(result.a在庫['TEST-06'], 3);
  assert.strictEqual(Object.keys(result.a在庫).length, 1);
});

test('Yahoo厳密集計: code+sub-code重複と非整数数量は入力エラーにする', () => {
  const duplicate = [
    'code,name,sub-code,quantity,allow-overdraft,stock-close',
    'TEST-07,商品,TEST-07a,1,,',
    'TEST-07,商品,TEST-07a,2,,'
  ].join('\n');
  assert.ok(context.全件再計算_YahooCSV厳密集計_(duplicate).error.indexOf('重複') >= 0);
  const invalid = [
    'code,name,sub-code,quantity,allow-overdraft,stock-close',
    'TEST-07,商品,TEST-07a,1.5,,'
  ].join('\n');
  assert.ok(context.全件再計算_YahooCSV厳密集計_(invalid).error.indexOf('整数') >= 0);
});

test('Yahoo厳密集計: sub-code空の基本商品行は在庫>0でも中止せず対象外として数える', () => {
  // 実CSV(yahoo全在庫260720.csv)の3行目と同型。a/b運用外の直在庫商品が605行あり、
  // 全件検算(棚卸.jsのYahooCSV集計_)は同じ行を要確認リスト行きで通している
  const csv = [
    'code,name,sub-code,quantity,allow-overdraft,stock-close',
    '021224nurie,親商品(在庫0),,0,,',
    '021224photo01,訳アリ現品(直在庫),,1,,',
    'TEST-08,商品,TEST-08a,2,,'
  ].join('\n');
  const result = context.全件再計算_YahooCSV厳密集計_(csv);
  assert.strictEqual(result.error, '');
  assert.strictEqual(result.a在庫['TEST-08'], 2);
  assert.strictEqual(Object.keys(result.a在庫).length, 1);
  assert.strictEqual(result.subなし件数, 1); // 在庫>0だけ数える(親行の在庫0は含めない)
});

test('Yahoo厳密集計: codeが空で在庫>0の行は従来どおり入力エラーにする', () => {
  const csv = [
    'code,name,sub-code,quantity,allow-overdraft,stock-close',
    ',商品,TEST-09a,1,,'
  ].join('\n');
  assert.ok(context.全件再計算_YahooCSV厳密集計_(csv).error.indexOf('code') >= 0);
});

test('Yahoo厳密集計: 負数quantityは中止しない(在庫切れ承諾で正常発生。数値妥当性はSKU単位で再構築側が判定)', () => {
  // 実CSV(yahoo全在庫260720.csv)の実例: JMEE128b=-1。b枠はa在庫に入らないので影響ゼロで通す
  const csv = [
    'code,name,sub-code,quantity,allow-overdraft,stock-close',
    'JMEE128,商品,JMEE128b,-1,,',
    'JMEE128,商品,JMEE128a,2,,',
    'TEST-10,商品,TEST-10a,-3,,'
  ].join('\n');
  const result = context.全件再計算_YahooCSV厳密集計_(csv);
  assert.strictEqual(result.error, '');
  assert.strictEqual(result.a在庫['JMEE128'], 2);
  assert.strictEqual(result.a在庫['TEST-10'], -3); // a行の負数はそのまま返し全件再計算_再構築_の「Yahoo数量不正」ガードがSKU単位でブロックする
});

test('再構築: Yahoo a在庫が負のSKUはYahoo数量不正としてSKU単位でブロックし全体は止めない', () => {
  const result = context.全件再計算_再構築_({
    supplies: [{ems:'EMS-N',code:'TEST-11',qty:1,arrival:'2026-07-01',row:2,status:'到着済'}],
    shipped: [],
    currentOrders: [{ban:'N01',sku:'TEST-11b',code:'TEST-11',qty:1,orderDate:'2026-07-02',row:3,route:'韓国取寄せ'}],
    yahooA: {'TEST-10': -3, 'TEST-11': 0}
  });
  assert.ok(result.blockedSkus.indexOf('TEST-10') >= 0);
  assert.ok(result.issues.some(r => r.type === 'Yahoo数量不正' && r.sku === 'TEST-10'));
  assert.ok(result.blockedSkus.indexOf('TEST-11') < 0); // 他SKUの引当は通常どおり
  assert.strictEqual(result.ledgerRows.filter(r => r.状態 === '取り置き中' && r.受注番号 === 'N01').length, 1);
});

test('反映前検査: 未確認・現物ありのキャンセル戻しは全件置換を止める', () => {
  const rows = [
    {取置ID:'A',状態:'キャンセル戻し',戻し処理結果:'未確認'},
    {取置ID:'B',状態:'キャンセル戻し',戻し処理結果:'現物あり'},
    {取置ID:'C',状態:'発送済み',戻し処理結果:''}
  ];
  const result = context.全件再計算_未解決台帳作業_(rows);
  assert.deepStrictEqual(json(result.map(r => r.取置ID)), ['A','B']);
});

test('通常④連携: ブロックSKUのEMS供給だけを除外する', () => {
  const supplies = [
    {ems:'EMS-1',code:'TEST-08',sourceCode:'TEST-08',qty:2},
    {ems:'EMS-2',code:'SAFE-01',sourceCode:'SAFE-01',qty:1}
  ];
  const filtered = context.全件再計算_ブロック供給_(supplies, new Set(['TEST-08']));
  assert.deepStrictEqual(json(filtered.map(r => r.code)), ['SAFE-01']);
});

test('通常④連携: 再構築台帳に元EMS到着日を保持して入荷日を復元できる', () => {
  const result = context.全件再計算_再構築_({
    supplies: [{ems:'EMS-9',code:'TEST-09',qty:1,arrival:'2026-06-22',row:9,status:'在庫反映済み'}],
    shipped: [],
    currentOrders: [{ban:'900',sku:'TEST-09b',code:'TEST-09',qty:1,orderDate:'2026-06-20 09:00:00',row:90,route:'韓国取寄せ'}],
    yahooA: {'TEST-09':0}
  });
  const active = result.ledgerRows.find(r => r.状態 === '取り置き中');
  assert.strictEqual(context.全件再計算_台帳到着日_(active), '2026-06-22');
});

test('反映: 韓国取寄せだけ旧入荷日・EMSのクリア対象にし台湾中国と即納を守る', () => {
  assert.strictEqual(context.全件再計算_韓国派生クリア対象_('★在庫の設定-お取り寄せ（韓国から）','商品'), true);
  assert.strictEqual(context.全件再計算_韓国派生クリア対象_('台湾取り寄せ','商品'), false);
  assert.strictEqual(context.全件再計算_韓国派生クリア対象_('お取り寄せ','中国限定商品'), false);
  assert.strictEqual(context.全件再計算_韓国派生クリア対象_('即納','商品'), false);
});

test('反映: 再計算時は新規P割当0件でもforceWriteで旧P列を全消去する', () => {
  let written = null;
  const sheet = {getRange: () => ({setValues: values => { written = values; }})};
  const plan = {error:'',writes:[],forceWrite:true,sheet,startRow:7,colP:16,rowCount:2,values:[[''],['']],backgrounds:[],summary:{記入:0}};
  context.発注共有P列計画を反映_(plan);
  assert.deepStrictEqual(json(written), [[''],['']]);
});

// ===== 開始前在庫(棚確認済みの現物)の引き継ぎ =====

test('再構築: 開始前在庫は需要から差し引き・台帳へ引き継ぎ・供給も消費して誤ブロックしない', () => {
  const result = context.全件再計算_再構築_({
    supplies: [{ems:'EMS-H1',code:'HOLD-01',qty:2,arrival:'2026-07-01',row:1,status:'在庫反映済み'}],
    shipped: [],
    currentOrders: [{ban:'500',sku:'HOLD-01b',code:'HOLD-01',qty:2,orderDate:'2026-07-02 10:00:00',row:10,route:'韓国取寄せ'}],
    yahooA: {},
    initialHolds: [{取置ID:'INIT|500|HOLD-01|HOLD-01B',状態:'取り置き中',受注番号:'500',商品コード:'HOLD-01',SKU:'HOLD-01b',取り置き数量:2,取置元種別:'開始前在庫',登録日時:'2026-07-17 12:00:00'}]
  });
  // 需要2は開始前在庫2で充足済み→新たなEMS取り置き行を作らない
  assert.ok(!result.ledgerRows.some(r => String(r.取置ID).indexOf('REBUILD|ACTIVE') === 0), '二重確保しない');
  // 開始前在庫行はそのまま置換後台帳へ引き継ぐ
  const carried = result.ledgerRows.find(r => r.取置ID === 'INIT|500|HOLD-01|HOLD-01B');
  assert.ok(carried, '棚確認済みの現物を消さない');
  assert.strictEqual(carried.状態, '取り置き中');
  // 棚の現物ぶんEMS供給を消費→説明不能余りにならずブロックもしない
  assert.ok(result.movementRows.some(r => r.移動先 === '開始前在庫充当' && r.数量 === 2));
  assert.ok(!result.issues.some(i => i.type === 'EMS説明不能余り'));
  assert.strictEqual(result.blockedSkus.length, 0);
  assert.strictEqual(result.summary[0].開始前在庫, 2);
  assert.strictEqual(result.summary[0].判定, 'OK');
});

test('再構築: 開始前在庫がEMSで説明できなくても情報のみでブロックしない', () => {
  const result = context.全件再計算_再構築_({
    supplies: [], shipped: [],
    currentOrders: [{ban:'501',sku:'HOLD-02b',code:'HOLD-02',qty:1,orderDate:'2026-07-02 10:00:00',row:11,route:'韓国取寄せ'}],
    yahooA: {},
    initialHolds: [{取置ID:'INIT|501|HOLD-02|HOLD-02B',状態:'取り置き中',受注番号:'501',商品コード:'HOLD-02',SKU:'HOLD-02b',取り置き数量:1,取置元種別:'開始前在庫'}]
  });
  assert.strictEqual(result.blockedSkus.length, 0, '棚の現物はEMS履歴の説明を必須にしない');
  assert.ok(result.issues.some(i => i.type === '開始前在庫EMS外'));
  assert.ok(result.ledgerRows.some(r => r.取置ID === 'INIT|501|HOLD-02|HOLD-02B'));
  assert.ok(!result.issues.some(i => i.type === '現在取寄せ未引当'), '需要は開始前在庫で充足済み');
});

test('再構築: 一部だけ開始前在庫の注文は残り数量だけEMSへ割り当てる', () => {
  const result = context.全件再計算_再構築_({
    supplies: [{ems:'EMS-H3',code:'HOLD-03',qty:3,arrival:'2026-07-01',row:1,status:'在庫反映済み'}],
    shipped: [],
    currentOrders: [{ban:'502',sku:'HOLD-03b',code:'HOLD-03',qty:3,orderDate:'2026-07-02 10:00:00',row:12,route:'韓国取寄せ'}],
    yahooA: {},
    initialHolds: [{取置ID:'INIT|502|HOLD-03|HOLD-03B',状態:'取り置き中',受注番号:'502',商品コード:'HOLD-03',SKU:'HOLD-03b',取り置き数量:1,取置元種別:'開始前在庫'}]
  });
  const rebuilt = result.ledgerRows.filter(r => String(r.取置ID).indexOf('REBUILD|ACTIVE') === 0);
  assert.strictEqual(rebuilt.length, 1);
  assert.strictEqual(rebuilt[0].取り置き数量, 2, '3個中1個は棚確保済み→EMSからは2個だけ');
  assert.ok(result.movementRows.some(r => r.移動先 === '開始前在庫充当' && r.数量 === 1));
  assert.ok(!result.issues.some(i => i.type === 'EMS説明不能余り'), '3個の箱=棚1+割当2で全て説明できる');
  assert.strictEqual(result.blockedSkus.length, 0);
});

test('再構築: ブロックSKUでも開始前在庫の引き継ぎ行は消さない', () => {
  const result = context.全件再計算_再構築_({
    supplies: [],
    shipped: [{ban:'600',itemId:'IT-600',sku:'BLK-01b',code:'BLK-01',qty:1,shipDate:'2026-07-10',orderDate:'2026-07-01 09:00:00'}],
    currentOrders: [], yahooA: {},
    initialHolds: [{取置ID:'INIT|601|BLK-01|BLK-01B',状態:'取り置き中',受注番号:'601',商品コード:'BLK-01',SKU:'BLK-01b',取り置き数量:1,取置元種別:'開始前在庫'}]
  });
  assert.ok(result.blockedSkus.indexOf('BLK-01') >= 0, '発送供給不足でブロック');
  assert.ok(result.ledgerRows.some(r => r.取置ID === 'INIT|601|BLK-01|BLK-01B'), '棚確認済みの現物は残す');
});

test('再構築: 発送済み・手動解除の台帳行は開始前在庫として引き継がない', () => {
  const result = context.全件再計算_再構築_({
    supplies: [], shipped: [], currentOrders: [], yahooA: {},
    initialHolds: [
      {取置ID:'INIT|700|OLD-01|OLD-01B',状態:'発送済み',受注番号:'700',商品コード:'OLD-01',SKU:'OLD-01b',取り置き数量:1,取置元種別:'開始前在庫'},
      {取置ID:'INIT|701|OLD-02|OLD-02B',状態:'手動解除',受注番号:'701',商品コード:'OLD-02',SKU:'OLD-02b',取り置き数量:1,取置元種別:'開始前在庫'},
      {取置ID:'EMS|EG1|702',状態:'取り置き中',受注番号:'702',商品コード:'OLD-03',SKU:'OLD-03b',取り置き数量:1,取置元種別:'EMS',元EMS番号:'EG1'}
    ]
  });
  assert.strictEqual(result.ledgerRows.length, 0, '履歴とEMS由来は事実データから再構築する側');
});

test('反映日時: 引き継いだ開始前在庫の登録日時は上書きしない', () => {
  const dated = context.全件再計算_反映日時を設定_({
    ledgerRows: [
      {取置ID:'REBUILD|ACTIVE|1',登録日時:''},
      {取置ID:'INIT|1',登録日時:'2026-07-17 12:00:00'}
    ],
    movementRows: []
  }, '2026-07-19 19:00:00');
  assert.strictEqual(dated.ledger[0].登録日時, '2026-07-19 19:00:00');
  assert.strictEqual(dated.ledger[1].登録日時, '2026-07-17 12:00:00', '棚確認した時点の記録を守る');
});

// ===== 別便(ダニエル等)EMS番号付き注文の除外 =====

test('再構築: EMSリストに無い別便EMS番号付きの取寄せ注文は韓国供給を横取りしない', () => {
  const result = context.全件再計算_再構築_({
    supplies: [{ems:'EMS-K1',code:'DAN-01',qty:1,arrival:'2026-07-01',row:1,status:'在庫反映済み'}],
    shipped: [],
    currentOrders: [
      {ban:'700',sku:'DAN-01b',code:'DAN-01',qty:1,orderDate:'2026-07-01 08:00:00',row:20,route:'韓国取寄せ',boxEms:'EJ111111111KR'},
      {ban:'701',sku:'DAN-01b',code:'DAN-01',qty:1,orderDate:'2026-07-02 08:00:00',row:21,route:'韓国取寄せ',boxEms:'EMS-K1'}
    ],
    yahooA: {}
  });
  assert.ok(result.issues.some(i => i.type === '別便供給の注文(ダニエル等)' && i.ban === '700'));
  const active = result.ledgerRows.filter(r => r.状態 === '取り置き中');
  assert.strictEqual(active.length, 1);
  assert.strictEqual(active[0].受注番号, '701', '古い別便注文ではなくリスト内EMSの注文へ');
  assert.strictEqual(result.blockedSkus.length, 0);
});

if (failures) process.exit(1);
