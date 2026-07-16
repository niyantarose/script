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

// ===== 残必要計算_: 取り置き台帳と今回計画だけを差し引く =====

test('必要数は取り置き台帳数量と今回計画数量だけを差し引く', () => {
  assert.strictEqual(context.残必要計算_({qty:3,取り置き中数量:1,alloc:0}),2);
  assert.strictEqual(context.残必要計算_({qty:3,取り置き中数量:1,alloc:2}),0);
});

test('旧取り置き数・入荷日・履歴数量は必要数へ加えない', () => {
  assert.strictEqual(context.残必要計算_({qty:3,取り置き数:3,入荷:true,履歴Alloc:3,取り置き中数量:0,alloc:0}),3);
});

test('取り置き中で全数確保された取り寄せ行は出荷準備OK', () => {
  assert.strictEqual(context.注文出荷準備OK_([{kbn:'取り寄せ',qty:2,取り置き中数量:2,alloc:0,キャンセル:false}]),true);
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

// ===== 台湾・中国ルート: 韓国EMSに供給が無いため手入力入荷日を確保として扱う =====
// 設計書は「台湾・中国ルートの在庫管理変更」を対象外と明記(=旧運用の維持が必須)。

test('別ルート判定: 選択肢/商品名の台湾・中国を検出する', () => {
  assert.strictEqual(context.引当_別ルート判定_('★在庫の設定=お取り寄せ（中国から） ★種類の選択=2.ティル', ''), true);
  assert.strictEqual(context.引当_別ルート判定_('★在庫の設定=お取り寄せ（台湾から８月末発売予定）', ''), true);
  assert.strictEqual(context.引当_別ルート判定_('', '台湾版 まんが (初版限定版)'), true);
  assert.strictEqual(context.引当_別ルート判定_('★在庫の選択=お取り寄せ（韓国から）', '韓国 グッズ'), false);
  assert.strictEqual(context.引当_別ルート判定_('', ''), false);
});

test('残必要計算_: 別ルートの手入力入荷分を差し引く', () => {
  assert.strictEqual(context.残必要計算_({qty:2,取り置き中数量:0,alloc:0,別ルート済数量:2}), 0);
  assert.strictEqual(context.残必要計算_({qty:2,取り置き中数量:0,alloc:0,別ルート済数量:0}), 2);
});

test('注文出荷準備OK_: 台湾・中国の入荷済み行は出荷可能に数える', () => {
  assert.strictEqual(context.注文出荷準備OK_([
    {kbn:'取り寄せ',qty:1,取り置き中数量:1,alloc:0,キャンセル:false},
    {kbn:'取り寄せ',qty:1,取り置き中数量:0,alloc:0,別ルート済数量:1,キャンセル:false}
  ]), true);
  assert.strictEqual(context.注文出荷準備OK_([
    {kbn:'取り寄せ',qty:1,取り置き中数量:0,alloc:0,別ルート済数量:0,キャンセル:false}
  ]), false);
});

test('引当行状態_: 別ルート入荷済みは着済(ラベンダー)表示', () => {
  const cfg={色_グレー:'g',色_水:'m',色_橙:'o',色_黄:'y',色_着:'l'};
  assert.strictEqual(context.引当行状態_({kbn:'取り寄せ',qty:1,別ルート:true,入荷:true,別ルート済数量:1},cfg).st,'着済');
  assert.strictEqual(context.引当行状態_({kbn:'取り寄せ',qty:1,別ルート:true,入荷:false,別ルート済数量:0},cfg).st,'在庫待ち');
});

test('今回行判定_: 別ルートは入荷日が今日のときだけ今回扱い(当日②で出荷可能へ)', () => {
  const today=new Date(); today.setHours(0,0,0,0);
  assert.strictEqual(context.今回行判定_({alloc:0,別ルート:true,入荷:true,入荷日値:today}), true);
  const past=new Date(today.getTime()-86400000);
  assert.strictEqual(context.今回行判定_({alloc:0,別ルート:true,入荷:true,入荷日値:past}), false);
  assert.strictEqual(context.今回行判定_({alloc:1}), true);
});

process.exitCode = failures ? 1 : 0;
console.log(failures ? `FAILURES: ${failures}` : 'ALL TESTS PASSED');
