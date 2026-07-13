// EMS大邱作業データ → EMSリスト転送の照合計画と転送ズレ修復の純粋ロジックのテスト
// 実行: GASアップスクリプト直下で node tests/project19_daegu_transfer.test.js
// 背景: 転送済み行を大邱側で後編集（EMS番号/商品コード/数量/発注NO振り直し）すると
//       既存行を認識できず最下部に重複追記され、古い行が残る問題の対策。
const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

const context = {
  console,
  Date,
  Set,
  Utilities: {
    formatDate: (d) => {
      const p = n => ('0' + n).slice(-2);
      return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
    }
  },
  Logger: { log: () => {} },
  CalendarApp: { EventColor: { BLUE: 1, ORANGE: 2, GREEN: 3 } },
  SpreadsheetApp: {
    getActive: () => ({ getSheetByName: () => null, toast: () => {} }),
    getActiveSpreadsheet: () => ({ getSheetByName: () => null }),
    getUi: () => ({ alert: () => {} })
  },
  ScriptApp: { getProjectTriggers: () => [] },
  LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock: () => {} }) }
};
vm.createContext(context);
[
  'Project_19/エクセルからデータ取得.js', // normCode_ / codeKeys_ / _emsDateOnly_ / isBlank_
  'Project_19/【大邱】データ転送.js'
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

// ---- テストデータ作成ヘルパー ----

// EMS大邱作業データの行（A入荷日,B発送日,D EMS番号,H商品コード,I数量,K品目,T購入No）
function srcRow(pno, track, code, qty, arr, ship) {
  const r = new Array(20).fill('');
  r[0] = arr || '';
  r[1] = ship || '';
  r[3] = track;
  r[7] = code;
  r[8] = qty;
  r[10] = 'Comic Book';
  r[19] = pno;
  return r;
}

// EMSリストの行（A No,B入荷日,C発送日,F購入No,I商品コード,J数量,M EMS番号）
function dstRow(no, arr, ship, pno, code, qty, track) {
  const r = new Array(15).fill('');
  r[0] = no;
  r[1] = arr;
  r[2] = ship;
  r[3] = '⇒';
  r[5] = pno;
  r[6] = '未着';
  r[8] = code;
  r[9] = qty;
  r[10] = 'Comic Book';
  r[12] = track;
  return r;
}

function dstHeader() {
  const r = new Array(15).fill('');
  r[0] = 'No.';
  return r;
}

// 転送側の計画を素の行データから作る（本番と同じ構築関数を使う）
function makePlan(srcRows, dstVals) {
  const candidates = [];
  const sourceExactKeys = new Set();
  srcRows.forEach((r, i) => {
    const o = context.EMS_転送候補_(r, i + 1);
    if (!o) return;
    candidates.push(o);
    if (o.purchaseNo) o.exactKeys.forEach(k => sourceExactKeys.add(k));
  });
  const existing = context.EMS_既存キー情報構築_(dstVals, sourceExactKeys);
  const stats = context.EMS_転送統計_(candidates);
  return context.EMS_転送照合計画_(candidates, existing, stats);
}

// ===== EMS_転送照合計画_: 転送済み行の後編集を追記でなく既存行更新にする =====

test('完全一致（日付も同じ）は追記・修正・日付更新なし', () => {
  const plan = makePlan(
    [srcRow('20260704_08_3', 'EG049827401KR', 'KAORUHANA21S', 1, '26/07/09', '26/07/09')],
    [dstHeader(), dstRow(1, '26/07/09', '26/07/09', '20260704_08_3', 'KAORUHANA21S', 1, 'EG049827401KR')]
  );
  assert.strictEqual(plan.appends.length, 0);
  assert.strictEqual(plan.fixes.length, 0);
  assert.strictEqual(plan.dateUpdates.length, 0);
});

test('EMS番号の後変更: 追記せず既存行のM列を更新（★コピペ実例）', () => {
  const plan = makePlan(
    [srcRow('20260627_05_3', 'EG049579302KR', '★コピペ', 3, '26/07/03', '26/07/03')],
    [dstHeader(), dstRow(577, '26/07/03', '26/07/06', '20260627_05_3', '★コピペ', 3, 'EG049624465KR')]
  );
  assert.strictEqual(plan.appends.length, 0, '重複追記しない');
  const fix = plan.fixes.find(f => f.col === 13);
  assert.ok(fix, 'M列の修正がある');
  assert.strictEqual(fix.row, 2);
  assert.strictEqual(fix.value, 'EG049579302KR');
  // 発送日も 07/06 -> 07/03 に追随する
  assert.ok(plan.dateUpdates.some(u => u.col === 3), '発送日更新がある');
});

test('商品コードの後修正: 追記せず既存行のI列を更新（DPHOTO09→DPHOTO09P実例）', () => {
  const plan = makePlan(
    [srcRow('20260702_02_1', 'EG049624664KR', 'DPHOTO09P', 1, '26/07/05', '26/07/06')],
    [dstHeader(), dstRow(624, '26/07/05', '26/07/06', '20260702_02_1', 'DPHOTO09', 1, 'EG049624664KR')]
  );
  assert.strictEqual(plan.appends.length, 0, '重複追記しない');
  const fix = plan.fixes.find(f => f.col === 9);
  assert.ok(fix, 'I列の修正がある');
  assert.strictEqual(fix.value, 'DPHOTO09P');
});

test('数量の後変更: 追記せず既存行のJ列を更新', () => {
  const plan = makePlan(
    [srcRow('20260706_11_1', 'EG049827401KR', 'JPSJCM42-02-EX', 2, '26/07/09', '26/07/09')],
    [dstHeader(), dstRow(772, '26/07/09', '26/07/09', '20260706_11_1', 'JPSJCM42-02-EX', 1, 'EG049827401KR')]
  );
  assert.strictEqual(plan.appends.length, 0);
  const fix = plan.fixes.find(f => f.col === 10);
  assert.ok(fix, 'J列の修正がある');
  assert.strictEqual(String(fix.value), '2');
});

test('発注NO振り直し: looseで既存行に一致し、F列を現行番号へ更新（20260614_07→08実例）', () => {
  const plan = makePlan(
    [srcRow('20260614_08_1', 'EG048959176KR', '10116039', 1, '26/06/16', '26/06/17')],
    [dstHeader(), dstRow(60, '26/06/16', '26/06/17', '20260614_07_1', '10116039', 1, 'EG048959176KR')]
  );
  assert.strictEqual(plan.appends.length, 0);
  const fix = plan.fixes.find(f => f.col === 6);
  assert.ok(fix, 'F列の修正がある');
  assert.strictEqual(fix.value, '20260614_08_1');
});

test('looseで一致してもdst購入Noが現大邱に実在する番号ならF列は触らない', () => {
  const plan = makePlan(
    [
      srcRow('20260701_01_1', 'EG000000001KR', 'CODEA', 1, '26/07/01', '26/07/02'),
      srcRow('20260630_09_9', 'EG000000009KR', 'CODEZ', 5, '26/07/01', '26/07/02')
    ],
    // dst行の購入No 20260630_09_9 は大邱の別行として現存する → 振り直しではない
    [dstHeader(), dstRow(1, '26/07/01', '26/07/02', '20260630_09_9', 'CODEA', 1, 'EG000000001KR')]
  );
  assert.strictEqual(plan.fixes.filter(f => f.col === 6).length, 0, 'F列は更新しない');
});

test('新規行は従来どおり追記', () => {
  const plan = makePlan(
    [srcRow('20260708_05_1', 'EG049827401KR', 'SPYANI08S', 1, '26/07/09', '26/07/09')],
    [dstHeader()]
  );
  assert.strictEqual(plan.appends.length, 1);
  assert.strictEqual(plan.fixes.length, 0);
});

test('分割発送（同一 購入No+コード+数量 が別EMS番号で2行）は両方とも完全一致で維持', () => {
  const plan = makePlan(
    [
      srcRow('20260402_01_1', 'ES396936615KR', 'KRSJCM03-0506S', 16, '26/05/20', '26/05/28'),
      srcRow('20260402_01_1', 'ES396936624KR', 'KRSJCM03-0506S', 16, '26/05/20', '26/05/28')
    ],
    [
      dstHeader(),
      dstRow(1, '26/05/20', '26/05/28', '20260402_01_1', 'KRSJCM03-0506S', 16, 'ES396936615KR'),
      dstRow(3, '26/05/20', '26/05/28', '20260402_01_1', 'KRSJCM03-0506S', 16, 'ES396936624KR')
    ]
  );
  assert.strictEqual(plan.appends.length, 0);
  assert.strictEqual(plan.fixes.length, 0);
});

test('曖昧なケース（同一 購入No+コード+数量 の候補が複数）は従来どおり追記に退避', () => {
  const plan = makePlan(
    [
      srcRow('20260402_01_1', 'ES396936615KR', 'KRSJCM03-0506S', 16, '26/05/20', '26/05/28'),
      srcRow('20260402_01_1', 'ES396936624KR', 'KRSJCM03-0506S', 16, '26/05/20', '26/05/28')
    ],
    // dstには第3のEMS番号の行が1行だけ → どちらの行か確定できない
    [dstHeader(), dstRow(1, '26/05/20', '26/05/28', '20260402_01_1', 'KRSJCM03-0506S', 16, 'ES396936638KR')]
  );
  assert.strictEqual(plan.fixes.filter(f => f.col === 13).length, 0, 'EMS番号は書き換えない');
  assert.strictEqual(plan.appends.length, 2, '両方追記される（従来動作に退避）');
});

// ===== EMS_転送ズレ修復計画_: 既に発生した重複/旧番号の修復 =====

function makeRepairPlan(srcRows, dstVals) {
  const srcList = context.EMS_修復元一覧_(srcRows);
  const dstList = context.EMS_修復先一覧_(dstVals);
  return context.EMS_転送ズレ修復計画_(srcList, dstList);
}

test('修復: 重複ペア（孤児が上・EMS番号違い）は上の行に統合して下の重複行を削除', () => {
  const plan = makeRepairPlan(
    [srcRow('20260627_05_3', 'EG049579302KR', '★コピペ', 3, '26/07/03', '26/07/03')],
    [
      dstHeader(),                                                                          // 行1
      dstRow(577, '26/07/03', '26/07/06', '20260627_05_3', '★コピペ', 3, 'EG049624465KR'), // 行2 = 孤児
      dstRow(774, '26/07/03', '26/07/03', '20260627_05_3', '★コピペ', 3, 'EG049579302KR')  // 行3 = 再追記
    ]
  );
  assert.strictEqual(plan.collapses.length, 1);
  assert.strictEqual(plan.collapses[0].mergeIntoRow, 2, '孤児行(上)に統合');
  assert.strictEqual(plan.collapses[0].deleteRow, 3, '再追記行(下)を削除');
  assert.strictEqual(plan.pnoFixes.length, 0);
  assert.strictEqual(plan.unresolved.length, 0);
});

test('修復: 重複ペア（孤児が下・コード違い）は孤児行だけ削除', () => {
  const plan = makeRepairPlan(
    [srcRow('20260629_02_1', 'EG049624465KR', '10117126', 1, '26/07/03', '26/07/06')],
    [
      dstHeader(),                                                                                  // 行1
      dstRow(538, '26/07/03', '26/07/06', '20260629_02_1', '10117126', 1, 'EG049624465KR'),         // 行2 = 正
      dstRow(559, '26/07/03', '26/07/06', '20260629_02_1', 'RECIPE42/10117126', 1, 'EG049624465KR') // 行3 = 孤児
    ]
  );
  assert.strictEqual(plan.collapses.length, 1);
  assert.strictEqual(plan.collapses[0].mergeIntoRow, undefined, '統合は不要');
  assert.strictEqual(plan.collapses[0].deleteRow, 3, '孤児行(下)を削除');
});

test('修復: 発注NO旧番号はF列を現行番号へ更新（行削除しない）', () => {
  const plan = makeRepairPlan(
    [srcRow('20260614_08_1', 'EG048959176KR', '10116039', 1, '26/06/16', '26/06/17')],
    [dstHeader(), dstRow(60, '26/06/16', '26/06/17', '20260614_07_1', '10116039', 1, 'EG048959176KR')]
  );
  assert.strictEqual(plan.collapses.length, 0);
  assert.strictEqual(plan.pnoFixes.length, 1);
  assert.strictEqual(plan.pnoFixes[0].row, 2);
  assert.strictEqual(plan.pnoFixes[0].value, '20260614_08_1');
});

test('修復: 対応が見つからない孤児は unresolved として報告のみ', () => {
  const plan = makeRepairPlan(
    [srcRow('20260701_01_1', 'EG000000001KR', 'CODEA', 1, '26/07/01', '26/07/02')],
    [
      dstHeader(),
      dstRow(1, '26/07/01', '26/07/02', '20260701_01_1', 'CODEA', 1, 'EG000000001KR'),
      dstRow(2, '26/06/01', '26/06/02', '20260601_01_1', 'GONE', 9, 'EG000000002KR')
    ]
  );
  assert.strictEqual(plan.collapses.length, 0);
  assert.strictEqual(plan.pnoFixes.length, 0);
  assert.strictEqual(plan.unresolved.length, 1);
});

test('修復: 全行一致なら何も計画しない', () => {
  const plan = makeRepairPlan(
    [
      srcRow('20260704_08_3', 'EG049827401KR', 'KAORUHANA21S', 1, '26/07/09', '26/07/09'),
      srcRow('20260704_08_4', 'EG049827401KR', 'HNTBLUE06-ALD', 1, '26/07/09', '26/07/09')
    ],
    [
      dstHeader(),
      dstRow(763, '26/07/09', '26/07/09', '20260704_08_3', 'KAORUHANA21S', 1, 'EG049827401KR'),
      dstRow(764, '26/07/09', '26/07/09', '20260704_08_4', 'HNTBLUE06-ALD', 1, 'EG049827401KR')
    ]
  );
  assert.strictEqual(plan.collapses.length, 0);
  assert.strictEqual(plan.pnoFixes.length, 0);
  assert.strictEqual(plan.unresolved.length, 0);
});

test('修復: 同一ペアの twin は二重に使わない（孤児2つ・twin1つなら片方は unresolved）', () => {
  const plan = makeRepairPlan(
    [srcRow('20260627_05_3', 'EG049579302KR', '★コピペ', 3, '26/07/03', '26/07/03')],
    [
      dstHeader(),
      dstRow(1, '26/07/03', '26/07/06', '20260627_05_3', '★コピペ', 3, 'EG049624465KR'), // 孤児1
      dstRow(2, '26/07/03', '26/07/07', '20260627_05_3', '★コピペ', 3, 'EG049999999KR'), // 孤児2
      dstRow(3, '26/07/03', '26/07/03', '20260627_05_3', '★コピペ', 3, 'EG049579302KR')  // twin
    ]
  );
  assert.strictEqual(plan.collapses.length, 1, 'twinは1回だけ使う');
  assert.strictEqual(plan.unresolved.length, 1, '残りは報告のみ');
});

process.exitCode = failures ? 1 : 0;
console.log(failures ? `FAILURES: ${failures}` : 'ALL TESTS PASSED');
