// 取り置き（開始前の手元在庫）まわりの純粋ロジックのテスト
// 実行: GASアップスクリプト直下で node tests/project24_torioki.test.js
// 背景: 取り置きで足りている注文が今回便から再引当され、取り置きから出た出荷済みが
//       今回の箱を推測消費して「帳尻は合うが引当先が違う」状態になる問題への対策。
const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

const context = {
  console,
  Date,
  Set,
  SpreadsheetApp: { openById: () => ({ getSheetByName: () => null }) },
  Utilities: { formatDate: () => '' },
  Logger: { log: () => {} }
};
vm.createContext(context);
[
  'Project_24/引当.js',
  'Project_24/取り置き台帳.js',
  'Project_24/P列自動記入.js',
  'Project_24/消込台帳.js',
  'Project_24/ダニエル余り.js'
].forEach(f => vm.runInContext(fs.readFileSync(f, 'utf8'), context));

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

// ===== 残必要計算_: 現在の取り置き中数量だけを差し引く =====

test('取り置き中数量が注文数を満たす行は今回便からの必要数0', () => {
  assert.strictEqual(context.残必要計算_({qty:1, 取り置き中数量:1}), 0);
});

test('取り置き中数量が一部だけなら残りだけ必要', () => {
  assert.strictEqual(context.残必要計算_({qty:3, 取り置き中数量:1}), 2);
});

test('現在台帳の取り置きなしでは旧履歴を必要数から差し引かない', () => {
  assert.strictEqual(context.残必要計算_({qty:2}), 2);
  assert.strictEqual(context.残必要計算_({qty:2, alloc:1}), 1);
  assert.strictEqual(context.残必要計算_({qty:2, alloc:1, 履歴Alloc:1}), 1);
});

test('引当・取り置きの合計が注文数を超えても負にならない', () => {
  assert.strictEqual(context.残必要計算_({qty:1, alloc:1, 取り置き中数量:1}), 0);
});

// ===== P列需要: 行単位で発送済みの分割を除外する =====

test('P列需要は発送済みの分割行を除外し未発送行だけ残す', () => {
  const M = { 出荷日: 0, 出荷日毎: 1 };
  assert.strictEqual(context.P列需要対象行_(['', ''], M), true);
  assert.strictEqual(context.P列需要対象行_(['', '2026/07/14'], M), false);
  assert.strictEqual(context.P列需要対象行_(['2026-07-14', ''], M), false);
});

// ===== 受注明細表示: 集計確保数を未発送の分割行だけへ配分する =====

test('発送済み2個行は空欄で未発送1個行だけ確保1を表示する', () => {
  const result = context.受注明細_現在確保を行配分_([
    {qty:2, 発送済み:true, rowNumber:10},
    {qty:1, 発送済み:false, rowNumber:11}
  ], 1);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.rows)), [
    {確保済み:'', 不足:'', 現在表示:false, 発送済み:true},
    {確保済み:1, 不足:0, 現在表示:true, 発送済み:false}
  ]);
  assert.strictEqual(result.未配分, 0);
});

test('未発送分割行は上から注文数量を上限に確保数を配る', () => {
  const one = context.受注明細_現在確保を行配分_([
    {qty:2, 発送済み:false, rowNumber:20},
    {qty:1, 発送済み:false, rowNumber:21}
  ], 1);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(one.rows.map(r => [r.確保済み, r.不足]))), [[1,1],[0,1]]);

  const full = context.受注明細_現在確保を行配分_([
    {qty:2, 発送済み:false, rowNumber:20},
    {qty:1, 発送済み:false, rowNumber:21}
  ], 3);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(full.rows.map(r => [r.確保済み, r.不足]))), [[2,0],[1,0]]);
});

test('現在確保が未発送注文数を超えても表示は注文数まで', () => {
  const result = context.受注明細_現在確保を行配分_([
    {qty:2, 発送済み:false, rowNumber:30},
    {qty:1, 発送済み:false, rowNumber:31}
  ], 4);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(result.rows.map(r => r.確保済み))), [2,1]);
  assert.strictEqual(result.未配分, 1);
});

test('注文数量が数値でなければ受注明細行番号を含めて停止する', () => {
  assert.throws(
    () => context.受注明細_現在確保を行配分_([{qty:'不明', 発送済み:false, rowNumber:44}], 1),
    /44行目.*注文数量/
  );
});

test('表示列を再作成した後の列マップで発送済み分割行を判定する', () => {
  const originals = {
    getActive: context.SpreadsheetApp.getActive,
    map: context.列マップ_,
    columns: context.受注明細_確保列を用意_,
    emsColumn: context.EMS番号列を用意_,
    shipped: context.引当_行出荷済み_,
    key: context.取り置き_行キー_
  };
  const sourceRows = [
    ['10117477','MRBLUE41','MRBLUE41',2,'','','','', '2026/07/14',''],
    ['10117477','MRBLUE41','MRBLUE41',1,'','','','','','']
  ];
  const writes = {};
  let mapCalls = 0, emsColumnCalls = 0, emsExists = false;
  const recv = {
    getLastRow: () => 3,
    getLastColumn: () => 10,
    getRange: (row, col, rowCount) => ({
      getValues: () => col === 1 ? sourceRows : [['旧EMS1'],['旧EMS2']],
      setValues: values => { writes[col] = JSON.parse(JSON.stringify(values)); }
    })
  };
  try {
    context.SpreadsheetApp.getActive = () => ({getSheetByName: () => recv});
    context.列マップ_ = () => {
      mapCalls++;
      return mapCalls === 1
        ? {hr:1,番号:0,SKU:1,コード:2,個数:3,出荷日:4,出荷日毎:5}
        : {hr:1,番号:0,SKU:1,コード:2,個数:3,出荷日:7,出荷日毎:8};
    };
    context.受注明細_確保列を用意_ = () => emsExists
      ? ({確保済み:6,不足:7,確保内訳:8})
      : ({確保済み:5,不足:6,確保内訳:7});
    context.EMS番号列を用意_ = () => { emsColumnCalls++; emsExists = true; return 5; };
    context.引当_行出荷済み_ = (row, M) => !!row[M.出荷日] || !!row[M.出荷日毎];
    context.取り置き_行キー_ = row => [String(row.ban||row.受注番号||''),String(row.sku||row.SKU||''),String(row.code||row.商品コード||'')].join('|');

    context.受注明細_確保列を書く_([{
      受注番号:'10117477',SKU:'MRBLUE41',商品コード:'MRBLUE41',確保済み:1,
      確保内訳:'自動1(EG050152967KR)',現在EMS:'EG050152967KR',判定:''
    }]);

    assert.strictEqual(mapCalls, 2, '表示列作成後に列マップを再取得する');
    assert.strictEqual(emsColumnCalls, 2, '全列作成後のEMS列位置を再取得する');
    assert.deepStrictEqual(writes[6], [[''],[1]]);
    assert.deepStrictEqual(writes[7], [[''],[0]]);
    assert.deepStrictEqual(writes[5], [[''],['EG050152967KR']]);
  } finally {
    context.SpreadsheetApp.getActive = originals.getActive;
    context.列マップ_ = originals.map;
    context.受注明細_確保列を用意_ = originals.columns;
    context.EMS番号列を用意_ = originals.emsColumn;
    context.引当_行出荷済み_ = originals.shipped;
    context.取り置き_行キー_ = originals.key;
  }
});

// ===== 取り置き登録: 入力規則を二次元配列で一括計画する =====

test('通常・棚戻し・即納の入力規則を列別配列へ振り分ける', () => {
  const plan = context.取り置き_入力規則計画_([
    {判定:''},
    {判定:'棚戻し待ち'},
    {判定:'即納'}
  ], 5, '棚規則', '戻し規則');
  assert.deepStrictEqual(JSON.parse(JSON.stringify(plan)), {
    棚確認:[['棚規則'],[null],[null],[null],[null]],
    処理:[[null],['戻し規則'],[null],[null],[null]]
  });
});

test('候補0行でも旧管理行数をnullで消せる', () => {
  const plan = context.取り置き_入力規則計画_([], 2, '棚規則', '戻し規則');
  assert.deepStrictEqual(JSON.parse(JSON.stringify(plan)), {
    棚確認:[[null],[null]],
    処理:[[null],[null]]
  });
});

// ===== 取り置き出荷_: 台帳メモによる人為オーバーライド =====

test('台帳メモに「取り置き」がある出荷済み行を判定できる', () => {
  assert.strictEqual(context.取り置き出荷_({メモ:'取り置き'}), true);
  assert.strictEqual(context.取り置き出荷_({メモ:'取置分から出荷'}), true);
  assert.strictEqual(context.取り置き出荷_({メモ:'取り置きから'}), true);
  assert.strictEqual(context.取り置き出荷_({メモ:''}), false);
  assert.strictEqual(context.取り置き出荷_({メモ:'通常出荷'}), false);
  assert.strictEqual(context.取り置き出荷_({}), false);
});

// ===== ダニエル余り: 取り置き出荷はダニエル便も消費しない =====

test('取り置きメモ付きの出荷はダニエル余りからも差し引かない', () => {
  const rows = context.ダニエル余り集計_({
    記録: [{ems:'EE111KR', code:'AAA1', qty:5}],
    大邱ems: [],
    出荷済: [
      {ban:'1001', code:'AAA1', sku:'', qty:2, 入荷日:'', メモ:'取り置き'},
      {ban:'1002', code:'AAA1', sku:'', qty:1, 入荷日:'', メモ:''}
    ],
    受注: [], 反映: null
  });
  const r = rows.find(x => x.code === 'AAA1');
  assert.strictEqual(r.出荷済, 1, '取り置き出荷2個は除外され通常出荷1個だけ');
  assert.strictEqual(r.余り, 4);
});

process.exitCode = failures ? 1 : 0;
console.log(failures ? `FAILURES: ${failures}` : 'ALL TESTS PASSED');
