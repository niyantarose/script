// 三段階(現物確認済み/到着済/先行)を含む注文分類の純粋ロジックテスト
// 実行: GASアップスクリプト直下で node tests/project24_three_stage_classification.test.js
const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

const context = {
  console, Date, Set, Map, JSON, Math,
  SpreadsheetApp: { openById: () => ({ getSheetByName: () => null }) }
};
vm.createContext(context);
vm.runInContext(fs.readFileSync('Project_24/引当.js', 'utf8'), context);
vm.runInContext(fs.readFileSync('Project_24/取り置き計算.js', 'utf8'), context);

let failures = 0;
function test(name, fn) {
  try { fn(); console.log(`PASS ${name}`); }
  catch (error) { failures++; console.error(`FAIL ${name}: ${error.message}`); }
}

const CFG = {色_緑:'#b7e1cd',色_黄:'#fce8b2',色_グレー:'#efefef',色_橙:'#fcd5b4',色_赤:'#f4cccc',色_水:'#cfe2f3',色_着:'#d9d2e9'};
// 段階付与済みの取り寄せ行を作る。省略時は数量1・確保なし。
function line(over){
  return Object.assign({kbn:'取り寄せ',qty:1,届:'',キャンセル:false,段階付与:true,
    現物確認済み数量:0,到着済引当数量:0,先行引当数量:0,別ルート済数量:0}, over||{});
}
const FUTURE='2999/12/31';

test('全数先行+入金済は出荷可能(行表示は先行)', () => {
  const arr=[line({先行引当数量:1})];
  assert.strictEqual(context.注文区分判定_(arr,true,false), 'ship');
  assert.strictEqual(context.引当行状態_(arr[0],CFG).st, '先行');
  assert.strictEqual(context.引当行状態_(arr[0],CFG).color, CFG.色_水);
});

test('一部だけ先行の注文は部分在庫', () => {
  const arr=[line({qty:2,先行引当数量:1})];
  assert.strictEqual(context.注文区分判定_(arr,true,false), 'part');
});

test('全数先行+未来希望日は希望日待ち', () => {
  const arr=[line({先行引当数量:1,届:FUTURE})];
  assert.strictEqual(context.注文区分判定_(arr,true,false), 'hold');
});

test('全数先行+未入金は出荷GO未入金', () => {
  const arr=[line({先行引当数量:1})];
  assert.strictEqual(context.注文区分判定_(arr,false,false), 'keep');
});

test('全数先行+代引きは未入金でも出荷可能', () => {
  const arr=[line({先行引当数量:1})];
  assert.strictEqual(context.注文区分判定_(arr,false,true), 'ship');
});

test('何も確保できない注文は引当待ち', () => {
  const arr=[line({}),line({})];
  assert.strictEqual(context.注文区分判定_(arr,true,false), 'wait');
});

test('混在注文(現物+到着済+先行)も1分類だけへ出る', () => {
  const arr=[line({現物確認済み数量:1}),line({到着済引当数量:1}),line({先行引当数量:1})];
  assert.strictEqual(context.注文区分判定_(arr,true,false), 'ship');
  const s=context.注文充足集計_(arr);
  assert.strictEqual(s.注文数量,3);
  assert.strictEqual(s.現物確認済み,1);
  assert.strictEqual(s.到着済引当,1);
  assert.strictEqual(s.先行引当,1);
  assert.strictEqual(s.確保総数,3);
  assert.strictEqual(s.不足,0);
});

test('箱到着で分類は変わらず行の段階表示だけ先行→到着済へ変わる', () => {
  const before=[line({先行引当数量:1})], after=[line({到着済引当数量:1})];
  assert.strictEqual(context.注文区分判定_(before,true,false), context.注文区分判定_(after,true,false));
  assert.strictEqual(context.引当行状態_(before[0],CFG).st, '先行');
  assert.strictEqual(context.引当行状態_(after[0],CFG).st, '到着済');
  assert.strictEqual(context.引当行状態_(after[0],CFG).color, CFG.色_着);
});

test('物理出荷可: 全数先行はfalse・全数到着済/現物はtrue', () => {
  assert.strictEqual(context.注文物理出荷可_([line({先行引当数量:1})]), false);
  assert.strictEqual(context.注文物理出荷可_([line({到着済引当数量:1})]), true);
  assert.strictEqual(context.注文物理出荷可_([line({現物確認済み数量:1})]), true);
  assert.strictEqual(context.注文物理出荷可_([line({qty:2,到着済引当数量:1})]), false, '不足ありはfalse');
});

test('行状態: 現物全数=緑・不足=色なし・キャンセル/即納/指定なしは従来どおり', () => {
  assert.strictEqual(context.引当行状態_(line({現物確認済み数量:1}),CFG).color, CFG.色_緑);
  const lack=context.引当行状態_(line({qty:3,到着済引当数量:1}),CFG);
  assert.ok(/不足/.test(lack.st));
  assert.strictEqual(lack.color, null);
  assert.strictEqual(context.引当行状態_(Object.assign(line({}),{キャンセル:true}),CFG).st, 'キャンセル');
  assert.strictEqual(context.引当行状態_(Object.assign(line({}),{kbn:'即納'}),CFG).st, '即納');
  assert.strictEqual(context.引当行状態_(Object.assign(line({}),{kbn:'指定なし'}),CFG).st, '要確認');
});

test('段階未付与の旧経路は台帳確保+今回割当で従来同等に分類する', () => {
  const legacy={kbn:'取り寄せ',qty:1,届:'',キャンセル:false,取り置き中数量:1,alloc:0,別ルート済数量:0};
  assert.strictEqual(context.注文区分判定_([legacy],true,false), 'ship');
  const wait={kbn:'取り寄せ',qty:1,届:'',キャンセル:false,取り置き中数量:0,alloc:0,別ルート済数量:0};
  assert.strictEqual(context.注文区分判定_([wait],true,false), 'wait');
});

test('行へ反映: 計画後集計の段階数量を行へ付与し要移行(旧開始前在庫)は到着済相当に数える', () => {
  const l={ban:'700',code:'AAA',sku:'AAAb',qty:3,kbn:'取り寄せ',届:'',キャンセル:false,別ルート済数量:0};
  const key=context.取り置き_行キー_(l);
  const summary=()=>{
    const stageByKey={}; stageByKey[key]={現物確認済み数量:1,到着済引当数量:0,先行引当数量:1,合計確保数量:2,要移行数量:1,行内訳:[]};
    const activeByKey={}; activeByKey[key]=3;
    return {activeByKey,activeRowsByKey:{},stageByKey};
  };
  context.引当計画_行へ反映_([l],summary(),summary(),[]);
  assert.strictEqual(l.段階付与,true);
  assert.strictEqual(l.現物確認済み数量,1);
  assert.strictEqual(l.到着済引当数量,1,'要移行1を到着済相当に合算');
  assert.strictEqual(l.先行引当数量,1);
  assert.strictEqual(l.未引当数量,0);
  assert.strictEqual(l.主段階,'先行');
  assert.strictEqual(context.注文区分判定_([l],true,false),'ship');
  assert.strictEqual(context.注文物理出荷可_([l]),false,'先行を含むため物理出荷は不可');
});

if (failures) process.exit(1);
