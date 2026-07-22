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

process.exitCode = failures ? 1 : 0;
console.log(failures ? `${failures} FAILED` : 'ALL PASS');
