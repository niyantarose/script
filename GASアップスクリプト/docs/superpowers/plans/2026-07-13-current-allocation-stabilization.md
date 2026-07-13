# 現行引当方式の安定化 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 棚卸行とEMS番号空欄行を引当供給から完全に除外し、OFW304/305の数値枝番を別商品として固定し、旧棚卸由来だけを安全に解除して実EMSで再引当できる状態にする。

**Architecture:** 現行の②引当エンジンを維持し、`実EMS番号_` を供給判定の共通入口として使う。まず現在オンラインにある部分修正を回帰テストで固定し、次に旧棚卸解除の行変換を純粋関数へ分離し、最後に解除直後の自動②実行を廃止して「P列書き直し → EMS在庫更新 → ② → 全件検算」を明示的な手順にする。

**Tech Stack:** Google Apps Script V8、SpreadsheetApp/DriveApp、Node.js `assert` + `vm`、PowerShell、clasp安全同期スクリプト。

## Global Constraints

- 対象は `Project_24`、対応テスト、承認済み設計書、本計画だけとする。
- `OFW304-1`、`OFW304-2`、`OFW305-1`、`OFW305-2` は相互に別商品として扱う。
- EMS番号が空欄、または `棚卸` で始まる行は供給・P列・履歴・検算へ入れない。
- Yahoo末尾 `a` は比較材料だけに使い、EMS引当供給へ加えない。
- 台湾・中国ルートの手入力入荷日は保持する。
- 通常は全リセットを使わず、旧棚卸由来だけを解除する。
- 作業開始時に `git pull` と `tools/gas_pull_sync.ps1 Project_24` を実行する。
- Apps Scriptへの反映は `tools/gas_safe_push.ps1 Project_24` だけを使い、素の `clasp push` は使わない。
- 未追跡のv3設計・計画ファイルは本作業のコミットへ含めない。

---

### Task 1: 実EMS供給とOFWコードを回帰テストで固定する

**Files:**
- Create: `tests/project24_real_ems_only.test.js`
- Reference: `Project_24/引当.js:293`
- Reference: `Project_24/P列確定.js:23`
- Reference: `Project_24/P列自動記入.js:136`
- Reference: `Project_24/個別対応.js`
- Reference: `Project_24/引当履歴.js`
- Reference: `Project_24/入荷日チェック.js`
- Reference: `Project_24/全件検算.js:16`

**Interfaces:**
- Consumes: `normCode_(value)`, `codeKeys_(code)`, `実EMS番号_(ems)`, `EMS明細_(sheet)`, `EMS番号書戻し値_(current, line)`, `P列確定マップ_()`, `全件検算_集計_(src)`。
- Produces: 実EMS供給ルールとOFW4コード独立性を固定するテストファイル。

- [ ] **Step 1: 作業開始時の同期と差分確認を行う**

Run:

```powershell
git pull --ff-only
powershell -ExecutionPolicy Bypass -File tools\gas_pull_sync.ps1 Project_24
git status --short
```

Expected:

```text
Project_24 はオンラインとローカルが一致しています。
?? docs/superpowers/plans/2026-07-10-full-allocation-rebuild-v3.md
?? docs/superpowers/specs/2026-07-10-full-allocation-rebuild-v3-design.md
```

上記以外の変更があれば作業を止め、その変更の所有者と内容を確認する。

- [ ] **Step 2: 実EMS限定の特徴テストを作成する**

Create `tests/project24_real_ems_only.test.js` with:

```javascript
const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

const headers = ['ステータス列', 'EMS到着日', '商品コード', '数量', 'EMS番号', '注文番号'];
const pRows = [
  ['到着済', '2026-07-10', 'OFW304-1', '1', 'EG000000001KR', '10117001'],
  ['到着済', '2026-07-10', 'OFW304-2', '1', '棚卸20260710', '10117002'],
  ['到着済', '2026-07-10', 'OFW305-1', '2', '', '10117003'],
  ['在庫反映済み', '2026-07-09', 'OFW305-2', '2', 'EG000000002KR', '10117004']
];

const range = (values, displayValues = values) => ({
  getValues: () => values,
  getDisplayValues: () => displayValues
});

const externalSheet = {
  getLastRow: () => 6 + pRows.length,
  getLastColumn: () => headers.length,
  getRange: row => row === 6 ? range([headers]) : range(pRows)
};

const context = {
  console,
  Date,
  Set,
  SpreadsheetApp: {
    openById: () => ({ getSheetByName: () => externalSheet })
  }
};

vm.createContext(context);
[
  'Project_24/引当.js',
  'Project_24/P列確定.js',
  'Project_24/全件検算.js'
].forEach(file => vm.runInContext(fs.readFileSync(file, 'utf8'), context));

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

test('OFW304/305の数値枝番は4つの別コードとして残る', () => {
  const codes = ['OFW304-1', 'OFW304-2', 'OFW305-1', 'OFW305-2'];
  const normalized = codes.map(code => context.normCode_(code));
  assert.strictEqual(new Set(normalized).size, 4);
  codes.forEach(code => {
    assert.strictEqual(JSON.stringify(context.codeKeys_(code)), JSON.stringify([code]));
  });
  assert.strictEqual(context.normCode_('OFW305-1OFW304-2'), 'OFW305-1OFW304-2');
});

test('実EMS番号だけを有効にする', () => {
  assert.strictEqual(context.実EMS番号_('EG000000001KR'), true);
  assert.strictEqual(context.実EMS番号_('棚卸20260710'), false);
  assert.strictEqual(context.実EMS番号_(''), false);
  assert.strictEqual(context.実EMS番号_('   '), false);
});

test('EMS明細は行位置を保ったまま棚卸と番号空欄を数量0にする', () => {
  const values = [
    ['状態', '商品コード', '数量', 'EMS番号', 'EMS到着日'],
    ['到着済', 'OFW304-1', 1, 'EG000000001KR', '2026-07-10'],
    ['到着済', 'OFW304-2', 1, '棚卸20260710', '2026-07-10'],
    ['到着済', 'OFW305-1', 2, '', '2026-07-10']
  ];
  const sheet = { getDataRange: () => range(values) };
  const result = context.EMS明細_(sheet);
  assert.strictEqual(result.rows.length, 3);
  assert.strictEqual(result.rows[0][result.cols.コード], 'OFW304-1');
  assert.strictEqual(result.rows[0][result.cols.数量], 1);
  assert.strictEqual(result.rows[1][result.cols.コード], '');
  assert.strictEqual(result.rows[1][result.cols.数量], 0);
  assert.strictEqual(result.rows[2][result.cols.コード], '');
  assert.strictEqual(result.rows[2][result.cols.数量], 0);
  assert.strictEqual(result.除外, 2);
});

test('P列確定は到着済み実EMSだけを採用する', () => {
  const map = context.P列確定マップ_();
  assert.strictEqual(map['10117001'][0].key, 'OFW304-1');
  assert.strictEqual(map['10117002'], undefined);
  assert.strictEqual(map['10117003'], undefined);
  assert.strictEqual(map['10117004'], undefined);
});

test('全件検算は棚卸と番号空欄を供給へ加えない', () => {
  const result = context.全件検算_集計_({
    ems: [
      { code: 'OFW304-1', st: '到着済', qty: 1, arrival: '2026-07-10', ems: 'EG000000001KR' },
      { code: 'OFW304-1', st: '到着済', qty: 1, arrival: '2026-07-10', ems: '棚卸20260710' },
      { code: 'OFW305-1', st: '到着済', qty: 2, arrival: '2026-07-10', ems: '' }
    ],
    出荷済: [],
    受注: [],
    a在庫: null
  });
  assert.strictEqual(result.rows.length, 1);
  assert.strictEqual(result.rows[0].code, 'OFW304-1');
  assert.strictEqual(result.rows[0].到着済, 1);
});

test('EMS番号書戻しは実EMSを保持し棚卸番号を消す', () => {
  assert.strictEqual(context.EMS番号書戻し値_('EG000000001KR', { 入荷: true }), 'EG000000001KR');
  assert.strictEqual(context.EMS番号書戻し値_('棚卸20260710', { 入荷: true }), '');
  assert.strictEqual(context.EMS番号書戻し値_('', { 箱EMS: 'EG000000002KR' }), 'EG000000002KR');
});

test('P列・個別対応・履歴・入荷日チェックも実EMS判定へ接続されている', () => {
  const checks = [
    ['Project_24/P列自動記入.js', /実EMS番号_\(ev\[i\]\[0\]\)/],
    ['Project_24/P列確定.js', /実EMS番号_\(r\[cEms\]\)/],
    ['Project_24/個別対応.js', /実EMS番号_\(r\.EMS番号\)/],
    ['Project_24/引当履歴.js', /実EMS番号_\(r\[c\.EMS番号\]\)/],
    ['Project_24/入荷日チェック.js', /実EMS番号_\(r\[cE\]\)/],
    ['Project_24/全件検算.js', /実EMS番号_\(r\.ems\)/]
  ];
  checks.forEach(([file, pattern]) => {
    const source = fs.readFileSync(file, 'utf8');
    assert.ok(pattern.test(source), `${file} が実EMS番号_を通っていない`);
  });
});

if (failures) process.exit(1);
```

- [ ] **Step 3: 特徴テストを実行する**

Run:

```powershell
node tests\project24_real_ems_only.test.js
```

Expected: 7件すべて `PASS`。現在オンラインから同期済みの実EMS除外処理を、変更前の基準として固定する。

- [ ] **Step 4: テストファイルだけをコミットする**

Run:

```powershell
git add tests\project24_real_ems_only.test.js
git commit -m "test(hikiate): 実EMS限定とOFW枝番を固定"
```

Expected: `tests/project24_real_ems_only.test.js` だけを含むコミットが作成される。

---

### Task 2: 旧棚卸解除の行変換を純粋関数へ分離する

**Files:**
- Modify: `Project_24/入荷日チェック.js:1-44`
- Modify: `tests/project24_real_ems_only.test.js`

**Interfaces:**
- Consumes: `ems番号`, `入荷日`, `別ルート`。
- Produces: `旧棚卸割当解除値_(ems番号, 入荷日, 別ルート) -> {対象:boolean, EMS番号:any, 入荷日:any}`。

- [ ] **Step 1: 失敗する旧棚卸解除テストを追加する**

`tests/project24_real_ems_only.test.js` のVM読込一覧を次へ置き換える。

```javascript
[
  'Project_24/引当.js',
  'Project_24/P列確定.js',
  'Project_24/全件検算.js',
  'Project_24/入荷日チェック.js'
].forEach(file => vm.runInContext(fs.readFileSync(file, 'utf8'), context));
```

既存テストの末尾へ次を追加する。

```javascript
test('旧棚卸解除は韓国ルートだけ入荷日を消し実EMSを保持する', () => {
  assert.strictEqual(
    JSON.stringify(context.旧棚卸割当解除値_('棚卸20260710', '2026-07-10', false)),
    JSON.stringify({ 対象: true, EMS番号: '', 入荷日: '' })
  );
  assert.strictEqual(
    JSON.stringify(context.旧棚卸割当解除値_('棚卸20260710', '2026-07-10', true)),
    JSON.stringify({ 対象: true, EMS番号: '', 入荷日: '2026-07-10' })
  );
  assert.strictEqual(
    JSON.stringify(context.旧棚卸割当解除値_('EG000000001KR', '2026-07-10', false)),
    JSON.stringify({ 対象: false, EMS番号: 'EG000000001KR', 入荷日: '2026-07-10' })
  );
});
```

- [ ] **Step 2: テストが未定義で失敗することを確認する**

Run:

```powershell
node tests\project24_real_ems_only.test.js
```

Expected: `context.旧棚卸割当解除値_ is not a function` で1件失敗する。

- [ ] **Step 3: 純粋関数を実装する**

`Project_24/入荷日チェック.js` の先頭、公開関数より前へ追加する。

```javascript
function 旧棚卸割当解除値_(ems番号, 入荷日, 別ルート){
  const ems=String(ems番号==null?'':ems番号).trim();
  const 対象=/^棚卸/i.test(ems);
  return {
    対象:対象,
    EMS番号:対象?'':ems,
    入荷日:対象 && !別ルート ? '' : 入荷日
  };
}
```

- [ ] **Step 4: 既存の解除ループを純粋関数へ置き換える**

`Project_24/入荷日チェック.js` の対象行ループを次へ置き換える。

```javascript
対象.forEach(idx=>{
  const row=R[M.hr+idx];
  const 別ルート=/台湾|中国/.test(String(row[M.選択肢]||'')) ||
    (M.商品名>=0 && /台湾|中国/.test(String(row[M.商品名]||'')));
  const next=旧棚卸割当解除値_(ev[idx][0], iv[idx][0], 別ルート);
  ev[idx][0]=next.EMS番号;
  if(next.入荷日!==iv[idx][0]){
    iv[idx][0]=next.入荷日;
    入荷クリア++;
  }
});
```

- [ ] **Step 5: 集中テストと既存テストを実行する**

Run:

```powershell
node tests\project24_real_ems_only.test.js
node tests\project24_zenken_kensan.test.js
node tests\project24_arrived_box_color.test.js
```

Expected: 全テストが `PASS`。

- [ ] **Step 6: 純粋関数化をコミットする**

Run:

```powershell
git add Project_24\入荷日チェック.js tests\project24_real_ems_only.test.js
git commit -m "refactor(hikiate): 旧棚卸解除の行変換をテスト可能にする"
```

---

### Task 3: 旧棚卸解除直後の自動②実行を停止する

**Files:**
- Modify: `Project_24/入荷日チェック.js:1-63`
- Modify: `Project_24/引当.js:60-74`
- Modify: `tests/project24_real_ems_only.test.js`

**Interfaces:**
- Consumes: 旧棚卸解除処理の件数結果。
- Produces: `旧棚卸解除後手順_() -> string[]`。解除後は自動引当せず、P列書き直し、EMS在庫更新、②、全件検算を順に案内する。

- [ ] **Step 1: 解除後手順と自動②禁止の失敗テストを追加する**

`tests/project24_real_ems_only.test.js` の末尾へ追加する。

```javascript
test('旧棚卸解除後はP列書き直しから手動で進める', () => {
  assert.strictEqual(
    JSON.stringify(context.旧棚卸解除後手順_()),
    JSON.stringify(['P列を書き直す', 'EMS在庫を更新', '②引き当て実行', '全件検算レポート'])
  );
  const source=fs.readFileSync('Project_24/入荷日チェック.js','utf8');
  const start=source.indexOf('function 旧棚卸割当だけを解除して再引当本体_');
  const end=source.indexOf('// ===== 引当データの全リセット', start);
  const body=source.slice(start, end);
  assert.strictEqual(body.includes('引当実行_本体_();'), false);
});
```

- [ ] **Step 2: 現行コードでは失敗することを確認する**

Run:

```powershell
node tests\project24_real_ems_only.test.js
```

Expected: `旧棚卸解除後手順_ is not a function`、または自動 `引当実行_本体_()` 検出で失敗する。

- [ ] **Step 3: 次手順を返す純粋関数を追加する**

`Project_24/入荷日チェック.js` の `旧棚卸割当解除値_` の後へ追加する。

```javascript
function 旧棚卸解除後手順_(){
  return ['P列を書き直す', 'EMS在庫を更新', '②引き当て実行', '全件検算レポート'];
}
```

- [ ] **Step 4: 確認ダイアログから自動実行表現を除く**

確認文の末尾を次へ変更する。

```javascript
'実行前に引当ファイル全体を自動バックアップします。\n'+
'解除後は、P列書き直し→EMS在庫更新→②→全件検算の順で確認します。\n\n続行しますか？'
```

- [ ] **Step 5: 解除完了後は案内だけを表示する**

`SpreadsheetApp.flush()` 以降を次へ置き換え、`引当実行_本体_();` を削除する。

```javascript
SpreadsheetApp.flush();
const 手順=旧棚卸解除後手順_();
ss.toast(
  '旧棚卸割当 '+対象.length+'明細を解除（入荷日 '+入荷クリア+'／履歴 '+履歴クリア+'）',
  '🧹棚卸解除',
  8
);
ui.alert(
  '旧棚卸割当の解除完了',
  'バックアップ: '+backup+'\n'+
  '解除明細: '+対象.length+'\n'+
  '入荷日クリア: '+入荷クリア+'\n'+
  '履歴クリア: '+履歴クリア+'\n\n'+
  '次の手順:\n'+手順.map((s,i)=>(i+1)+') '+s).join('\n'),
  ui.ButtonSet.OK
);
```

- [ ] **Step 6: メニュー表示を実際の動作へ合わせる**

`Project_24/引当.js` のメニュー項目を次へ変更する。関数名は既存の図形・メニュー割当互換性のため変更しない。

```javascript
.addItem('🧹 旧棚卸割当だけを解除（再引当前）', '旧棚卸割当だけを解除して再引当')
```

- [ ] **Step 7: 集中テストと既存テストを実行する**

Run:

```powershell
node tests\project24_real_ems_only.test.js
node tests\project24_zenken_kensan.test.js
node tests\project24_arrived_box_color.test.js
```

Expected: 全テストが `PASS`。

- [ ] **Step 8: 自動②停止をコミットする**

Run:

```powershell
git add Project_24\入荷日チェック.js Project_24\引当.js tests\project24_real_ems_only.test.js
git commit -m "fix(hikiate): 棚卸解除後の再引当を確認手順に分離"
```

---

### Task 4: 全Project_24検証と安全反映を行う

**Files:**
- Verify: `Project_24/*.js`
- Verify: `tests/project24_real_ems_only.test.js`
- Verify: `tests/project24_zenken_kensan.test.js`
- Verify: `tests/project24_arrived_box_color.test.js`
- May update automatically: `tools/sync_state/Project_24.json`

**Interfaces:**
- Consumes: Task 1からTask 3のコミット済み変更。
- Produces: テスト・構文・差分確認済みのApps Scriptオンライン版と同期基準。

- [ ] **Step 1: 3つのProject_24テストを連続実行する**

Run:

```powershell
node tests\project24_real_ems_only.test.js
node tests\project24_zenken_kensan.test.js
node tests\project24_arrived_box_color.test.js
```

Expected: すべて `PASS`、終了コード0。

- [ ] **Step 2: 全Project_24 JavaScriptファイルの構文を検査する**

Run:

```powershell
Get-ChildItem Project_24 -Filter *.js -File | ForEach-Object { node --check $_.FullName; if($LASTEXITCODE -ne 0){ throw "syntax error: $($_.Name)" } }
```

Expected: 出力なし、終了コード0。

- [ ] **Step 3: コミット済み範囲と作業ツリーを確認する**

Run:

```powershell
git diff --check
git status --short
git log -5 --oneline
```

Expected: v3資料2件以外に未コミット変更がなく、直近にTask 1からTask 3のコミットが並ぶ。

- [ ] **Step 4: Project_24を安全pushする**

Run:

```powershell
powershell -ExecutionPolicy Bypass -File tools\gas_safe_push.ps1 Project_24
```

Expected: 両方変更の衝突がなく、Project_24がオンラインへ反映され、同期基準が更新される。衝突が出た場合は上書きを選ばず中止し、オンライン差分を取り込んで再確認する。

- [ ] **Step 5: 安全push後の状態を確認する**

Run:

```powershell
powershell -ExecutionPolicy Bypass -File tools\gas_pull_sync.ps1 Project_24
git status --short
```

Expected: `Project_24 はオンラインとローカルが一致しています。`。未追跡v3資料2件以外はクリーン。

- [ ] **Step 6: GitHubへコミットを共有する**

Run:

```powershell
git push origin main
```

Expected: `main` の直近コミットと安全push同期コミットが `origin/main` へ送信される。

---

### Task 5: ライブデータから旧棚卸を除去して実EMSで再引当する

**Files:**
- Live data: 引当ファイル `受注明細`, `引当履歴`, `EMS在庫`, `全件検算`
- Live data: 発注共有ファイル `EMSリスト`

**Interfaces:**
- Consumes: 安全反映済みProject_24、発注共有EMSリスト、現在の受注明細。
- Produces: 棚卸数量を含まない実EMSだけの引当結果。

- [ ] **Step 1: 発注共有ファイルをDrive上でコピーしてバックアップする**

バックアップ名を次の形式にする。

```text
発注共有_旧棚卸削除前_20260713_HHmmss
```

Expected: コピーを開け、元ファイルと `EMSリスト` の行数が一致する。

- [ ] **Step 2: 削除対象の棚卸行だけを抽出する**

`EMSリスト` のEMS番号列で `棚卸` から始まる行をフィルタし、各行の次の値を確認する。

```text
ステータス列 / EMS到着日 / 商品コード / 数量 / EMS番号 / 注文番号
```

Expected: 実在するEMS追跡番号を持つ行が1件も混ざっていない。対象行数を控える。

- [ ] **Step 3: 確認済みの棚卸行を削除する**

Expected: EMS番号列を再検索して `棚卸` 始まりが0件になる。実EMS行の数量と到着日は変更しない。

- [ ] **Step 4: 引当ファイルで旧棚卸割当だけを解除する**

Run from menu:

```text
🏠 メイン引当(EMS在庫)
→ 🧹 旧棚卸割当だけを解除（再引当前）
→ 確認文字「棚卸だけ」
```

Expected: Driveバックアップ名、解除明細数、入荷日クリア数、履歴クリア数が表示される。②は自動実行されない。

- [ ] **Step 5: 到着済み実EMS行のP列を書き直す**

Run from menu:

```text
♻️ P列を書き直す
```

Expected: 到着済み実EMS行だけが再計算され、棚卸行・EMS番号空欄行へP列が書かれない。

- [ ] **Step 6: EMS在庫を更新する**

Run from menu:

```text
🔄 EMS在庫を更新(色クリア＋最新化)
```

Expected: EMS在庫に `棚卸...` とEMS番号空欄の供給が存在しない。

- [ ] **Step 7: ②引当を実行する**

Run from menu:

```text
② 引き当て実行(実EMSリストのみ)
```

Expected: 完了メッセージに棚卸・EMS番号なしの供給除外件数が表示され、実EMS番号だけが受注明細へ書き戻される。

- [ ] **Step 8: 全件検算でOFW4コードを確認する**

Run from menu:

```text
🧮 全件検算レポート
```

確認対象:

```text
OFW304-1
OFW304-2
OFW305-1
OFW305-2
```

Expected:

```text
⚠️超過消費 = 0
⚠️入荷日ズレ = 0
```

`📦供給不足` が残る場合は、該当数量が実EMSへ未登録の待ち注文数と一致することを確認する。一致しない場合は自動補正せず、商品診断と受注番号診断で対象行を特定する。

- [ ] **Step 9: 完了記録を残す**

次を作業記録へ保存する。

```text
発注共有バックアップ名
引当ファイルバックアップ名
削除した棚卸行数
解除した受注明細数
P列書き直し結果
OFW4コードの全件検算結果
```

Expected: 再実行やロールバックに必要な2つのバックアップ名と検算結果を追跡できる。
