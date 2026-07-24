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

function functionBody(source, name) {
  const marker = `function ${name}(`;
  const start = source.indexOf(marker);
  assert.notStrictEqual(start, -1, `${name} が見つかりません`);
  const next = source.indexOf('\nfunction ', start + marker.length);
  return source.slice(start, next < 0 ? source.length : next);
}

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

test('日本在庫_戻り行: 確定済みで未出力のものだけを商品コード+確定日でまとめ、EMS番号は空にする', () => {
  const rows = context.日本在庫_戻り行_([
    {処理ID:'YAHOO|RETURN|A',商品コード:'AAA',数量:1,確定日時:'2026-07-24 09:00:00',出力日時:''},
    {処理ID:'YAHOO|RETURN|B',商品コード:'AAA',数量:2,確定日時:'2026/07/24 10:00:00',出力日時:''},
    {処理ID:'YAHOO|RETURN|C',商品コード:'BBB',数量:1,確定日時:'2026-07-24 09:00:00',出力日時:'2026-07-24 11:00:00'},
    {処理ID:'YAHOO|RETURN|D',商品コード:'',数量:1,確定日時:'2026-07-24 09:00:00',出力日時:''},
    {処理ID:'YAHOO|RETURN|E',商品コード:'ZERO',数量:0,確定日時:'2026-07-24 09:00:00',出力日時:''}
  ]);
  assert.strictEqual(rows.length, 1, '出力済み・無コード・数量0は出さない');
  assert.deepStrictEqual(json(rows[0]), ['戻り', '2026-07-24', 'AAA', 3, ''],
    '同じコード+同じ日はまとめ、EMS番号は空(⑤便締めに巻き込まれない)');
});

test('日本在庫_戻り行: CSVを作った(出力日時あり)分はリストから消える', () => {
  const 確定済み=[{処理ID:'YAHOO|RETURN|A',商品コード:'AAA',数量:2,確定日時:'2026-07-24 09:00:00',出力日時:''}];
  assert.strictEqual(context.日本在庫_戻り行_(確定済み).length, 1);
  const 出力後=確定済み.map(r=>Object.assign({},r,{出力日時:'2026-07-24 12:00:00'}));
  assert.strictEqual(context.日本在庫_戻り行_(出力後).length, 0, 'Yahooへ渡した分は日本在庫から外れる');
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

test('対象行: 戻り行は便を指定した出力(⑤便締め・従来📤)には混ざらない', () => {
  const rows=[
    {状態:'到着済',商品コード:'ARR01',余り数:1,EMS番号:'EG1'},
    {状態:'戻り',商品コード:'RET01',余り数:2,EMS番号:''}
  ];
  const 便指定 = context.Yahoo変更_対象行_(rows, new Set(['EG1']), new Set());
  assert.deepStrictEqual(json(便指定.対象.map(x=>x.商品コード)), ['ARR01'],
    '⑤に混ざると出力日時が付かず二重加算になるため戻りは出さない');
  const 全便 = context.Yahoo変更_対象行_(rows, new Set(), new Set());
  assert.deepStrictEqual(json(全便.対象.map(x=>x.商品コード)), ['ARR01','RET01'],
    '便指定なし(日本在庫のCSVボタン)なら箱の余りも戻りも出す');
});

test('対象行: 戻りのみモードは箱の余りを出さず戻り行だけ(旧ボタン互換)', () => {
  const r = context.Yahoo変更_対象行_([
    {状態:'到着済',商品コード:'ARR01',余り数:1,EMS番号:'EG1'},
    {状態:'戻り',商品コード:'RET01',余り数:2,EMS番号:''}
  ], new Set(), new Set(), {戻りのみ:true});
  assert.deepStrictEqual(json(r.対象.map(x=>x.商品コード)), ['RET01']);
});

// ===== 2026-07-24 日本在庫の「在庫対象外」表示とCSV出力の除外を同じ判定に統一 =====

test('対象外理由: 足せないコードは理由付きで返し、通常コードは空を返す', () => {
  const ex = new Set([context.normCode_('吉田富貴子')]);
  assert.strictEqual(context.Yahoo変更_対象外理由_('★コピペ', ex), '付属ポスター印(★コピペ)');
  assert.strictEqual(context.Yahoo変更_対象外理由_('PromotionalItem', ex), 'PromotionalItem(贈呈品)');
  assert.strictEqual(context.Yahoo変更_対象外理由_('10117508', ex), '受注番号形式コード');
  assert.strictEqual(context.Yahoo変更_対象外理由_('吉田富貴子', ex), 'マスタ除外コード');
  assert.strictEqual(context.Yahoo変更_対象外理由_('MRBLUE41', ex), '', '通常コードは足せる');
  assert.strictEqual(context.Yahoo変更_対象外理由_('', ex), '');
});

test('対象行: 日本在庫が「在庫対象外(理由)」と表示した行は、未着ではなく正しい理由で除外される', () => {
  const r = context.Yahoo変更_対象行_([
    {状態:'在庫対象外(付属ポスター印(★コピペ))',商品コード:'★コピペ',余り数:5,EMS番号:'EG1'},
    {状態:'在庫対象外(PromotionalItem(贈呈品))',商品コード:'PromotionalItem',余り数:5,EMS番号:'EG1'},
    {状態:'未着',商品コード:'FUT01',余り数:1,EMS番号:'EG1'},
    {状態:'到着済',商品コード:'OK01',余り数:1,EMS番号:'EG1'}
  ], new Set(['EG1']), new Set());
  assert.deepStrictEqual(json(r.対象.map(x=>x.商品コード)), ['OK01']);
  assert.deepStrictEqual(json(r.除外.map(x=>x.理由)),
    ['付属ポスター印(★コピペ)','PromotionalItem(贈呈品)','未着(先行)の余りはYahooへ足せない'],
    'コード起因の除外が未着理由に化けない');
});

test('手動出力記録: 全便CSVの中に対象便の同じEMS・コード・数量があれば締められる', () => {
  const record = {対象:[
    {EMS番号:'EG1',商品コード:'AAA',余り数:2},
    {EMS番号:'EG2',商品コード:'BBB',余り数:1},
    {EMS番号:'',商品コード:'RETURN',余り数:3}
  ]};
  assert.strictEqual(context.Yahoo変更_出力記録が対象を含む_([
    {EMS番号:'EG1',商品コード:'AAA',余り数:2}
  ], record), true);
});

test('手動出力記録: 数量変更・未記録・別便は締め済み扱いにしない', () => {
  const record = {対象:[{EMS番号:'EG1',商品コード:'AAA',余り数:2}]};
  assert.strictEqual(context.Yahoo変更_出力記録が対象を含む_([
    {EMS番号:'EG1',商品コード:'AAA',余り数:1}
  ], record), false);
  assert.strictEqual(context.Yahoo変更_出力記録が対象を含む_([
    {EMS番号:'EG2',商品コード:'AAA',余り数:2}
  ], record), false);
  assert.strictEqual(context.Yahoo変更_出力記録が対象を含む_([
    {EMS番号:'EG1',商品コード:'AAA',余り数:2}
  ], null), false);
  assert.strictEqual(context.Yahoo変更_出力記録が対象を含む_([
    {EMS番号:'EG1',商品コード:'AAA',余り数:2}
  ], {sig:'EG1|AAA|2'}), false, '旧⑤の自動出力記録は手動確認済みとみなさない');
});

test('⑤便締めはYahoo CSVを自動作成せず手動出力記録を検証してから全同期する', () => {
  const history = fs.readFileSync('Project_24/引当履歴.js', 'utf8');
  const body = functionBody(history, '到着済を在庫反映済みへ本体_');
  assert.doesNotMatch(body, /Yahoo在庫変更を出力本体_\s*\(/);
  assert.match(body, /Yahoo変更_出力記録が対象を含む_\s*\(/);
  assert.match(body, /引当_数値変更後全同期_\s*\(\s*\{[^}]*EMS更新\s*:\s*true/s);
});

process.exitCode = failures ? 1 : 0;
console.log(failures ? `${failures} FAILED` : 'ALL PASS');
