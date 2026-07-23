// 日本在庫の余り→★Yahoo在庫変更(code/sub-code/quantity/mode)出力の純粋ロジックテスト
// 実行: GASアップスクリプト直下で node tests/project24_yahoo_stock_export.test.js
const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

const context = {
  console, Date, Set, Map, JSON, Math,
  normCode_: value => String(value == null ? '' : value).trim().toUpperCase().replace(/_/g, '-'),
  P列指定文字列_: () => ''
};
vm.createContext(context);
[
  'Project_24/取り置き計算.js',
  'Project_24/全件再計算.js',
  'Project_24/Yahoo在庫変更出力.js'
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
const json = value => JSON.parse(JSON.stringify(value)); // vm跨ぎのdeepStrictEqual対策(既存テストと同じ)

test('対象行: EMS番号で絞り、PromotionalItemと受注番号形式を除外し、余り0以下は落とす', () => {
  const rows = [
    {商品コード: 'MOFUN-AS-24-2', 余り数: '2', EMS番号: 'EG1'},
    {商品コード: '★コピペ', 余り数: 5, EMS番号: 'EG1'},
    {商品コード: 'PromotionalItem', 余り数: 1, EMS番号: 'EG1'},
    {商品コード: '10117508', 余り数: 3, EMS番号: 'EG1'},
    {商品コード: 'KW58-1', 余り数: 1, EMS番号: 'EG2'}, // 便が違う
    {商品コード: 'ZERO01', 余り数: 0, EMS番号: 'EG1'}
  ];
  const r = context.Yahoo変更_対象行_(rows, new Set(['EG1']));
  assert.deepStrictEqual(json(r.対象.map(x => x.商品コード)), ['MOFUN-AS-24-2']);
  assert.strictEqual(r.対象[0].余り数, 2);
  assert.deepStrictEqual(json(r.除外.map(x => x.理由)), ['付属ポスター印(★コピペ)', 'PromotionalItem(贈呈品)', '受注番号形式コード']);
});

test('対象行: emsSetが空なら全便を対象にする', () => {
  const rows = [
    {商品コード: 'A1', 余り数: 1, EMS番号: 'EG1'},
    {商品コード: 'B2', 余り数: 1, EMS番号: 'EG2'}
  ];
  const r = context.Yahoo変更_対象行_(rows, new Set());
  assert.strictEqual(r.対象.length, 2);
});

test('対象行: マスタ除外コード(人名等)は除外一覧へ', () => {
  const rows = [
    {商品コード: '吉田富貴子', 余り数: 1, EMS番号: 'EG1'},
    {商品コード: 'A1', 余り数: 1, EMS番号: 'EG1'}
  ];
  const r = context.Yahoo変更_対象行_(rows, new Set(['EG1']), new Set([context.normCode_('吉田富貴子')]));
  assert.deepStrictEqual(json(r.対象.map(x => x.商品コード)), ['A1']);
  assert.strictEqual(r.除外[0].理由, 'マスタ除外コード');
});

test('サブコード逆引き: Yahoo CSVからsub-code→(code,sub)を作る(引用符・BOM・sub空スキップ)', () => {
  const csv = [
    '﻿"code","name","sub-code","quantity","allow-overdraft","stock-close"',
    '"MOFUN-AS-24","商品, カンマ入り","MOFUN-AS-24-2a","1","0","0"',
    '"MOFUN-AS-24","同上","MOFUN-AS-24-2b","9","0","0"',
    '"021224nurie","親行(sub空)","","0","0","0"'
  ].join('\n');
  const map = context.Yahoo変更_サブコード逆引き_(csv);
  assert.deepStrictEqual(json(map[context.normCode_('MOFUN-AS-24-2a')]), {code: 'MOFUN-AS-24', sub: 'MOFUN-AS-24-2a'});
  assert.strictEqual(map[context.normCode_('MOFUN-AS-24-2b')].sub, 'MOFUN-AS-24-2b');
  assert.strictEqual(Object.keys(map).length, 2);
});

test('行変換: 同一コードを合算し、逆引きヒットは親code+実sub-code+mode「+」、ミスは要確認へ', () => {
  const 対象 = [
    {商品コード: 'MOFUN-AS-24-2', 余り数: 2, EMS番号: 'EG1'},
    {商品コード: 'MOFUN-AS-24-2', 余り数: 1, EMS番号: 'EG1'}, // 同じコードの別行→合算
    {商品コード: 'SHINPIN01', 余り数: 1, EMS番号: 'EG1'}      // Yahooに無い
  ];
  const 逆引き = {};
  逆引き[context.normCode_('MOFUN-AS-24-2a')] = {code: 'MOFUN-AS-24', sub: 'MOFUN-AS-24-2a'};
  const r = context.Yahoo変更_行変換_(対象, 逆引き);
  assert.strictEqual(r.行.length, 1);
  assert.deepStrictEqual(json(r.行[0]), {code: 'MOFUN-AS-24', 'sub-code': 'MOFUN-AS-24-2a', quantity: 3, mode: '+'});
  assert.strictEqual(r.要確認.length, 1);
  assert.strictEqual(r.要確認[0].商品コード, 'SHINPIN01');
});

test('未着(先行)の余り行はYahoo出力の対象にせず除外理由を付ける', () => {
  const r = context.Yahoo変更_対象行_([
    {状態:'到着済',商品コード:'ARR01',余り数:1,EMS番号:'EMS-A'},
    {状態:'未着',商品コード:'FUT01',余り数:2,EMS番号:'EMS-F'}
  ], new Set(['EMS-A','EMS-F']), new Set());
  assert.strictEqual(r.対象.length, 1);
  assert.strictEqual(r.対象[0].商品コード, 'ARR01');
  assert.strictEqual(r.除外.length, 1);
  assert.strictEqual(r.除外[0].商品コード, 'FUT01');
  assert.ok(/未着|先行/.test(r.除外[0].理由));
});

// ===== 2026-07-23 棚へ戻ったキャンセル現物(戻り)を日本在庫へ合流し、CSV出力に含める =====

test('日本在庫_戻り行: 現物ありだけを商品コード+確定日でまとめ、EMS番号は空にする', () => {
  const rows = context.日本在庫_戻り行_([
    {状態:'キャンセル戻し',戻し処理結果:'現物あり',商品コード:'AAA',元EMS商品コード:'AAA',取り置き数量:1,更新日時:'2026-07-23T20:00:00'},
    {状態:'キャンセル戻し',戻し処理結果:'現物あり',商品コード:'AAA',元EMS商品コード:'AAA',取り置き数量:2,更新日時:'2026/07/23 21:00:00'},
    {状態:'キャンセル戻し',戻し処理結果:'未確認',商品コード:'BBB',取り置き数量:1,更新日時:'2026-07-23T20:00:00'},
    {状態:'キャンセル戻し',戻し処理結果:'再引当済み',商品コード:'CCC',取り置き数量:1,更新日時:'2026-07-23T20:00:00'},
    {状態:'キャンセル戻し',戻し処理結果:'Yahoo反映済み',商品コード:'DDD',取り置き数量:1,更新日時:'2026-07-23T20:00:00'},
    {状態:'取り置き中',戻し処理結果:'',商品コード:'EEE',取り置き数量:1,更新日時:'2026-07-23T20:00:00'}
  ]);
  assert.strictEqual(rows.length, 1, '現物ありだけ。未確認/再引当済み/Yahoo反映済み/取り置き中は出さない');
  assert.deepStrictEqual(json(rows[0]), ['戻り', '2026-07-23', 'AAA', 3, ''],
    '同じコード+同じ日はまとめ、EMS番号は空(⑤便締めに巻き込まれない)');
});

test('日本在庫_戻り行: 元EMS商品コードを優先し、数量0や無コードは落とす', () => {
  const rows = context.日本在庫_戻り行_([
    {状態:'キャンセル戻し',戻し処理結果:'現物あり',商品コード:'PARENT',元EMS商品コード:'BOX-CODE',取り置き数量:1,更新日時:'2026-07-23'},
    {状態:'キャンセル戻し',戻し処理結果:'現物あり',商品コード:'',取り置き数量:1,更新日時:'2026-07-23'},
    {状態:'キャンセル戻し',戻し処理結果:'現物あり',商品コード:'ZERO',取り置き数量:0,更新日時:'2026-07-23'}
  ]);
  assert.strictEqual(rows.length, 1);
  assert.strictEqual(rows[0][2], 'BOX-CODE', '箱のコードでYahoo逆引きするため元EMS商品コード優先');
});

test('対象行: 戻り行は便の指定に関係なく常に対象(未着チェックにも掛からない)', () => {
  const r = context.Yahoo変更_対象行_([
    {状態:'到着済',商品コード:'ARR01',余り数:1,EMS番号:'EG1'},
    {状態:'到着済',商品コード:'OTHER',余り数:1,EMS番号:'EG9'},
    {状態:'戻り',商品コード:'RET01',余り数:2,EMS番号:''}
  ], new Set(['EG1']), new Set());
  assert.deepStrictEqual(json(r.対象.map(x => x.商品コード)), ['ARR01', 'RET01'],
    '別便EG9は落ち、戻りは便指定に関わらず出る');
  assert.strictEqual(r.除外.length, 0, '戻りを未着扱いで除外しない');
});

test('対象行: 戻りのみモードは箱の余りを出さず戻り行だけをCSVにする', () => {
  const r = context.Yahoo変更_対象行_([
    {状態:'到着済',商品コード:'ARR01',余り数:1,EMS番号:'EG1'},
    {状態:'戻り',商品コード:'RET01',余り数:2,EMS番号:''},
    {状態:'戻り',商品コード:'★コピペ',余り数:1,EMS番号:''}
  ], new Set(), new Set(), {戻りのみ:true});
  assert.deepStrictEqual(json(r.対象.map(x => x.商品コード)), ['RET01']);
  assert.strictEqual(r.除外.length, 1, '戻りでも★コピペ等の在庫対象外は従来どおり除外');
  assert.strictEqual(r.除外[0].理由, '付属ポスター印(★コピペ)');
});

test('対象行: 戻りのみモードでない従来呼び出しは挙動が変わらない(emsSet空=全便)', () => {
  const r = context.Yahoo変更_対象行_([
    {商品コード:'A1',余り数:1,EMS番号:'EG1'},
    {商品コード:'B2',余り数:1,EMS番号:'EG2'}
  ], new Set());
  assert.strictEqual(r.対象.length, 2);
});

process.exitCode = failures ? 1 : 0;
console.log(failures ? `${failures} FAILED` : 'ALL PASS');
