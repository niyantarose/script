// 計算機シートの数式が「意図した行」を参照しているかを検証する。
// 行を1つ足すと全参照がずれるため、ラベル単位で意味的に突き合わせる。
// 実行: node tests/keisan-refs.test.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const src = fs.readFileSync(path.join(__dirname, '..', '計算機.js'), 'utf8');

const ctx = { console };
ctx.globalThis = ctx;
ctx.SpreadsheetApp = {};
ctx.HtmlService = {};
ctx.UrlFetchApp = {};
ctx.Logger = { log: () => {} };
vm.createContext(ctx);
vm.runInContext(src, ctx, { filename: '計算機.js' });

let failed = 0;
function eq(name, actual, expected) {
  const a = JSON.stringify(actual), e = JSON.stringify(expected);
  if (a !== e) { failed++; console.error(`[NG] ${name}: expected=${e} actual=${a}`); }
  else { console.log(`[OK] ${name}`); }
}

/** ソースから国別の行配列を取り出して評価する（帯設定行も連結） */
function buildRows(name) {
  const re = new RegExp('const ' + name + ' = (\\[[\\s\\S]*?\\])\\.concat\\(計算機_原価帯設定行_\\(\\)\\);');
  const m = src.match(re);
  if (!m) throw new Error('配列が見つかりません: ' + name);
  // 配列内で使う初期値変数をダミー定義（数値であれば検証に影響しない）
  ctx.twGoal = 0.25; ctx.twUwa = 4425;
  ctx.cnGoal = 0.25; ctx.cnUwa = 3000;
  ctx.krGoal = 0.25; ctx.krUwa = 6000;
  return vm.runInContext('(' + m[1] + ').concat(計算機_原価帯設定行_())', ctx);
}

function rowOf(rows, label) {
  const idx = rows.findIndex(r => r[0] === label);
  if (idx < 0) throw new Error('行が見つかりません: ' + label);
  return { index: idx + 1, value: rows[idx][1] };
}

/** ラベルの数式が参照している行のラベル一覧（重複なし・出現順） */
function refLabels(rows, label) {
  const cell = rowOf(rows, label);
  const refs = String(cell.value).match(/B\d+/g) || [];
  const out = [];
  refs.forEach(ref => {
    const n = parseInt(ref.slice(1), 10);
    const target = rows[n - 1];
    const name = target ? target[0] : '(範囲外 ' + ref + ')';
    if (out.indexOf(name) < 0) out.push(name);
  });
  return out;
}

const 帯ラベル = ['〜 2,000円まで', '〜 6,000円まで', '〜 12,000円まで', '〜 20,000円まで', '20,000円超'];

function check(国, 配列名, 価格ラベル, レートラベル, 原価内訳) {
  console.log('--- ' + 国 + ' ---');
  const rows = buildRows(配列名);

  // ① 目標粗利率は「数値」であること（原価帯と連動する数式にしない＝手動試算用）
  const goal = rowOf(rows, '目標粗利率');
  eq(国 + ': 目標粗利率は数値（原価帯と非連動）', typeof goal.value, 'number');

  // ② うわのせ欄の真下に売値があること
  const uwa = rowOf(rows, '実際のうわのせ');
  const live = rowOf(rows, '→ この上乗せでの売値');
  eq(国 + ': 売値表示はうわのせの真下', live.index, uwa.index + 1);
  eq(国 + ': 売値表示の参照', refLabels(rows, '→ この上乗せでの売値'), ['原価(円)', '実際のうわのせ']);

  // ③ 微調整セクションの連鎖
  eq(国 + ': 実際売値の参照', refLabels(rows, '実際売値'), ['原価(円)', '実際のうわのせ']);
  eq(国 + ': 実質粗利額の参照', refLabels(rows, '実質粗利額'), ['実際売値', '手数料合計', '原価(円)']);
  eq(国 + ': 実質粗利率の参照', refLabels(rows, '実質粗利率'), ['実質粗利額', '実際売値']);
  eq(国 + ': 判定の参照', refLabels(rows, '判定'), ['実質粗利率']);

  // ④ おすすめ売値＝入力の目標粗利率を使う
  ['おすすめ売値(10円丸め)', 'おすすめ売値(100円丸め)'].forEach(label => {
    eq(国 + ': ' + label + 'の参照', refLabels(rows, label), ['原価(円)', '目標粗利率', '手数料合計']);
  });

  // ⑤ 原価帯売値＝原価帯の適用率だけで全自動（入力の目標粗利率を参照しない）
  ['原価帯売値(10円丸め)', '原価帯売値(100円丸め)'].forEach(label => {
    const refs = refLabels(rows, label);
    eq(国 + ': ' + label + 'の参照', refs, ['原価(円)', '適用目標粗利率(原価帯から自動)', '手数料合計']);
    eq(国 + ': ' + label + 'は目標粗利率を参照しない', refs.indexOf('目標粗利率') < 0, true);
  });

  // ⑥ 自動計算まわり
  eq(国 + ': 原価(円)の参照', refLabels(rows, '原価(円)'), 原価内訳);
  const 適用 = refLabels(rows, '適用目標粗利率(原価帯から自動)');
  eq(国 + ': 適用率の参照', 適用, ['原価(円)'].concat(帯ラベル));
}

check('台湾', '台湾', '本の価格(元)', '使用レート(円/元)',
  ['本の価格(元)', '送料(元)', '使用レート(円/元)']);

check('中国', '中国', '商品価格(元)', '使用レート(円/元)',
  ['商品価格(元)', '送料(元)', '使用レート(円/元)']);

check('韓国', '韓国', '本の価格(ウォン)', '使用レート(円/ウォン)',
  ['本の価格(ウォン)', '使用レート(円/ウォン)', '重さ(g)', '重量係数(送料)']);

// 手数料合計は国ごとに構成が違うので個別に確認
eq('台湾: 手数料合計の参照', refLabels(buildRows('台湾'), '手数料合計'), ['モール手数料', '消費税']);
eq('中国: 手数料合計の参照', refLabels(buildRows('中国'), '手数料合計'), ['モール手数料', 'カード手数料', '消費税']);
eq('韓国: 手数料合計の参照', refLabels(buildRows('韓国'), '手数料合計'), ['モール手数料', '消費税']);

process.exitCode = failed ? 1 : 0;
console.log(failed ? `\n${failed} failed` : '\nall passed');
