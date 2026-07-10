// 全件検算レポート(改訂版v2)の純粋ロジックのテスト
// 実行: GASアップスクリプト直下で node tests/project24_zenken_kensan.test.js
// 設計: docs/superpowers/specs/2026-07-10-full-allocation-rebuild-v2-design.md
const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

const context = {
  console,
  Date,
  Set,
  SpreadsheetApp: { openById: () => ({ getSheetByName: () => null }) }
};
vm.createContext(context);
[
  'Project_24/引当.js',
  'Project_24/P列確定.js',
  'Project_24/消込台帳.js',
  'Project_24/棚卸.js',
  'Project_24/全件検算.js',
  'Project_24/入荷日チェック.js' // 参照はしないが構文チェックを兼ねて読み込む
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

// ===== Task2: Yahoo CSV読み込みの抽出(棚卸.js) =====

test('CSV行分解_: 引用符内カンマと""エスケープを扱える', () => {
  // vmレルムの配列はdeepStrictEqualでプロトタイプ不一致になるためJSONで比較
  assert.strictEqual(JSON.stringify(context.CSV行分解_('"a,b","c""d",e')), JSON.stringify(['a,b', 'c"d', 'e']));
});

test('YahooCSV集計_: 末尾aだけ加算しbは無視する', () => {
  const csv = [
    'code,name,sub-code,quantity,allow-overdraft,stock-close',
    'MOFUN-IV-09,"モフサンド,ぬい",MOFUN-IV-09a,3,,',
    'MOFUN-IV-09,"モフサンド,ぬい",MOFUN-IV-09b,91,,',
    'TAROT10,"タロット ""特装版""",TAROT10a,2,,',
    'NOSUB,サブ無し,,5,,',
    'ONLYCODE'
  ].join('\n');
  const r = context.YahooCSV集計_(csv);
  assert.strictEqual(r.a在庫['MOFUN-IV-09'], 3); // bの91は数えない
  assert.strictEqual(r.a在庫['TAROT10'], 2);
  assert.strictEqual(r.商品名['MOFUN-IV-09'], 'モフサンド,ぬい');
  assert.strictEqual(r.商品名['TAROT10'], 'タロット "特装版"');
  assert.strictEqual(r.subなし.length, 1);
  assert.ok(r.subなし[0].indexOf('NOSUB') === 0);
  assert.strictEqual(r.解析スキップ, 1); // ONLYCODE(4列未満)
});

test('YahooCSV集計_: 空のCSVはエラー', () => {
  assert.ok(context.YahooCSV集計_('').error);
});

// ===== Task3: 出荷済みの重複排除(消込台帳.js) =====

test('受注基底コード_: SKU優先・タグ除去・末尾A/B落とし', () => {
  assert.strictEqual(context.受注基底コード_('MOFUN-IV-09a', 'MOFUN-IV-09'), 'MOFUN-IV-09');
  assert.strictEqual(context.受注基底コード_('', 'POEM65（10116569）'), 'POEM65');
  assert.strictEqual(context.受注基底コード_('JPSJCM39-03S', ''), 'JPSJCM39-03S'); // SはA/Bではないので落とさない
});

test('出荷済み重複排除_: 表記ゆれの同一発送を1件(数量最大)にまとめる', () => {
  const rows = [
    { ban: '10117001', code: 'POEM65（10117001）', sku: '', qty: 1 },
    { ban: '10117001', code: 'POEM65', sku: 'POEM65a', qty: 3 },
    { ban: '10117002', code: 'POEM65', sku: '', qty: 2 },  // 受注番号違いは残す
    { ban: '10117001', code: 'RECIPE42', sku: '', qty: 1 }, // コード違いは残す
    { ban: '', code: 'POEM65', sku: '', qty: 1 },           // 受注番号なしは素通し
    { ban: '', code: 'POEM65', sku: '', qty: 1 }
  ];
  const out = context.出荷済み重複排除_(rows);
  assert.strictEqual(out.length, 5);
  const merged = out.filter(r => r.ban === '10117001' && context.受注基底コード_(r.sku, r.code) === 'POEM65');
  assert.strictEqual(merged.length, 1);
  assert.strictEqual(merged[0].qty, 3); // 数量最大の行を採用
});

// ===== Task4: 全件検算_集計_(全件検算.js) =====

const row = (r, c) => r.rows.find(x => x.code === c);

test('全件検算_集計_: 箱の個数以上の消費は⚠️超過消費', () => {
  const r = context.全件検算_集計_({
    ems: [{ code: 'TEST-01', st: '到着済', qty: 10, arrival: '2026-07-01' }],
    出荷済: [{ ban: '10117001', code: 'TEST-01', sku: '', qty: 8, 入荷日: '2026-07-01' }],
    受注: [{ code: 'TEST-01', sku: '', qty: 4, 選択肢: '取り寄せ', 商品名: '', 入荷日: '2026-07-01' }],
    a在庫: {}
  });
  const a = row(r, 'TEST-01');
  assert.strictEqual(a.出荷到着, 8);
  assert.strictEqual(a.確保到着, 4);
  assert.strictEqual(a.残, -2);
  assert.strictEqual(a.判定, '⚠️超過消費');
  assert.strictEqual(r.counts['⚠️超過消費'], 1);
});

test('全件検算_集計_: どの箱とも一致しない入荷日は⚠️入荷日ズレ', () => {
  const r = context.全件検算_集計_({
    ems: [{ code: 'TEST-01', st: '到着済', qty: 5, arrival: '2026-07-01' }],
    出荷済: [],
    受注: [{ code: 'TEST-01', sku: '', qty: 2, 選択肢: '取り寄せ', 商品名: '', 入荷日: '2026-07-05' }],
    a在庫: { 'TEST-01': 5 }
  });
  const a = row(r, 'TEST-01');
  assert.strictEqual(a.確保ズレ, 2);
  assert.strictEqual(a.判定, '⚠️入荷日ズレ');
});

test('全件検算_集計_: 待ちが見えている供給を超えると📦供給不足', () => {
  const r = context.全件検算_集計_({
    ems: [{ code: 'TEST-02', st: '発注済', qty: 2, arrival: '' }],
    出荷済: [],
    受注: [{ code: 'TEST-02', sku: '', qty: 5, 選択肢: '取り寄せ', 商品名: '', 入荷日: '' }],
    a在庫: {}
  });
  const b = row(r, 'TEST-02');
  assert.strictEqual(b.未着, 2);
  assert.strictEqual(b.待ち, 5);
  assert.strictEqual(b.判定, '📦供給不足');
});

test('全件検算_集計_: 到着済残がYahoo aより多いとℹ️箱残>Yahoo', () => {
  const r = context.全件検算_集計_({
    ems: [{ code: 'CCC', st: '到着済', qty: 5, arrival: '2026-07-01' }],
    出荷済: [],
    受注: [],
    a在庫: { CCC: 2 }
  });
  assert.strictEqual(row(r, 'CCC').判定, 'ℹ️箱残>Yahoo');
});

test('全件検算_集計_: Yahoo aが到着済残より多いとℹ️EMS外在庫', () => {
  const r = context.全件検算_集計_({
    ems: [],
    出荷済: [{ ban: '10117003', code: 'DDD', sku: '', qty: 1, 入荷日: '' }],
    受注: [],
    a在庫: { DDD: 3 }
  });
  const d = row(r, 'DDD');
  assert.strictEqual(d.出荷不明, 1);
  assert.strictEqual(d.判定, 'ℹ️EMS外在庫');
});

test('全件検算_集計_: 説明がつく行はOK(入荷日はDateでも一致する)', () => {
  const r = context.全件検算_集計_({
    ems: [{ code: 'EEE', st: '到着済', qty: 3, arrival: '2026-07-01' }],
    出荷済: [],
    受注: [{ code: 'EEE', sku: '', qty: 3, 選択肢: '取り寄せ', 商品名: '推しぬい', 入荷日: new Date('2026-07-01T00:00:00') }],
    a在庫: {}
  });
  const e = row(r, 'EEE');
  assert.strictEqual(e.確保到着, 3);
  assert.strictEqual(e.残, 0);
  assert.strictEqual(e.判定, 'OK');
  assert.strictEqual(e.名, '推しぬい');
});

test('全件検算_集計_: 台湾/中国・即納は対象外、b枝番SKUは基底コードへ丸まる', () => {
  const r = context.全件検算_集計_({
    ems: [],
    出荷済: [],
    受注: [
      { code: 'FFF', sku: 'FFFb', qty: 2, 選択肢: '取り寄せ', 商品名: '', 入荷日: '' },
      { code: 'GGG', sku: '', qty: 1, 選択肢: '台湾取り寄せ', 商品名: '', 入荷日: '' },
      { code: 'HHH', sku: '', qty: 1, 選択肢: '即納', 商品名: '', 入荷日: '' }
    ],
    a在庫: {}
  });
  assert.strictEqual(r.rows.length, 1);
  assert.strictEqual(r.rows[0].code, 'FFF');
  assert.strictEqual(r.rows[0].待ち, 2);
});

test('全件検算_集計_: 在庫反映済み箱に紐付く分は過去便へ分類し到着済残に影響しない', () => {
  const r = context.全件検算_集計_({
    ems: [{ code: 'III', st: '在庫反映済み', qty: 4, arrival: '2026-06-22' }],
    出荷済: [{ ban: '10117004', code: 'III', sku: '', qty: 1, 入荷日: '2026-06-22' }],
    受注: [{ code: 'III', sku: '', qty: 2, 選択肢: '取り寄せ', 商品名: '', 入荷日: '2026-06-22' }],
    a在庫: null // Yahoo照合なし
  });
  const i = row(r, 'III');
  assert.strictEqual(i.反映済, 4);
  assert.strictEqual(i.出荷過去, 1);
  assert.strictEqual(i.確保過去, 2);
  assert.strictEqual(i.残, 0);
  assert.strictEqual(i.判定, 'OK'); // Yahoo無しなのでℹ️系は出さない
  assert.strictEqual(i.yahooA, null);
});

test('全件検算_集計_: 並びは判定の重い順→コード順', () => {
  const r = context.全件検算_集計_({
    ems: [
      { code: 'AA-01', st: '到着済', qty: 1, arrival: '2026-07-01' },
      { code: 'ZZ-09', st: '到着済', qty: 1, arrival: '2026-07-01' }
    ],
    出荷済: [{ ban: '10117005', code: 'ZZ-09', sku: '', qty: 2, 入荷日: '2026-07-01' }],
    受注: [{ code: 'AA-01', sku: '', qty: 1, 選択肢: '取り寄せ', 商品名: '', 入荷日: '2026-07-01' }],
    a在庫: {}
  });
  assert.strictEqual(r.rows[0].code, 'ZZ-09'); // ⚠️超過消費が先頭
  assert.strictEqual(r.rows[0].判定, '⚠️超過消費');
  assert.strictEqual(r.rows[1].code, 'AA-01');
});

if (failures) process.exit(1);
