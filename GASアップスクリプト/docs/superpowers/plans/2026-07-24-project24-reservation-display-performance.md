# Project_24 取り置き表示整合・高速化 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 発送済み分割行へ現在確保を重複表示しないようにし、取り置き登録の入力規則を一括設定して6分タイムアウトを解消する。

**Architecture:** 引当台帳・P列・EMS供給数量は変更せず、受注番号・SKU単位の現在確保数を未発送の受注明細行へ上から配分する純粋関数を追加する。入力規則は候補行から二次元配列を作る純粋関数を追加し、管理範囲の一括クリアと列単位の `setDataValidations()` に置き換える。

**Tech Stack:** Google Apps Script V8、Google Sheets Spreadsheet Service、Node.js `assert`/`vm` による純粋ロジックテスト、clasp安全同期スクリプト

## Global Constraints

- 数量の正本である引当台帳、P列、EMS数量を変更しない。
- 発送済み行の現在引当表示は空欄にするが、引当台帳の発送履歴は削除しない。
- 未発送行の表示確保数合計を `min(現在有効確保数, 未発送注文数量合計)` と一致させる。
- お取り置きメモ、自動色分け、列幅、固定行、フィルタ、既存のプルダウン内容を保持する。
- 入力規則の行単位Range操作は禁止し、管理範囲だけを一括更新する。
- Project_24の反映は `tools/gas_safe_push.ps1 Project_24` だけを使用し、素の `clasp push` は使用しない。
- 実装前に `git pull` と `tools/gas_pull_sync.ps1 Project_24` を完了させる。
- 対象外のラベンダー色テスト失敗と、`../claude自動引当ツール/reports/` の変更には触れない。

---

### Task 1: 発送済み分割行を除外する行別表示配分

**Files:**
- Modify: `Project_24/引当.js:787-814`
- Modify: `Project_24/取り置き台帳.js:1051-1068`
- Test: `tests/project24_torioki.test.js`

**Interfaces:**
- Consumes: `引当_行出荷済み_(row, M)`, `取り置き_行キー_(row)`, candidate fields `確保済み`, `確保内訳`
- Produces: `受注明細_現在確保を行配分_(rows, totalSecured)` returning `{rows, 未配分}`; candidate field `現在EMS`; updated `受注明細_確保列を書く_(candidates)`

- [ ] **Step 1: 行別配分の失敗テストを追加する**

`tests/project24_torioki.test.js` のP列テスト直後へ追加する。

```javascript
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
  assert.deepStrictEqual(one.rows.map(r => [r.確保済み, r.不足]), [[1,1],[0,1]]);

  const full = context.受注明細_現在確保を行配分_([
    {qty:2, 発送済み:false, rowNumber:20},
    {qty:1, 発送済み:false, rowNumber:21}
  ], 3);
  assert.deepStrictEqual(full.rows.map(r => [r.確保済み, r.不足]), [[2,0],[1,0]]);
});

test('現在確保が未発送注文数を超えても表示は注文数まで', () => {
  const result = context.受注明細_現在確保を行配分_([
    {qty:2, 発送済み:false, rowNumber:30},
    {qty:1, 発送済み:false, rowNumber:31}
  ], 4);
  assert.deepStrictEqual(result.rows.map(r => r.確保済み), [2,1]);
  assert.strictEqual(result.未配分, 1);
});

test('注文数量が数値でなければ受注明細行番号を含めて停止する', () => {
  assert.throws(
    () => context.受注明細_現在確保を行配分_([{qty:'不明', 発送済み:false, rowNumber:44}], 1),
    /44行目.*注文数量/
  );
});
```

- [ ] **Step 2: テストが未定義で失敗することを確認する**

Run:

```powershell
node tests/project24_torioki.test.js
```

Expected: `受注明細_現在確保を行配分_ is not a function` を含むFAIL。

- [ ] **Step 3: 純粋な行別配分関数を追加する**

`Project_24/引当.js` の `受注明細_確保列を書く_` より前へ追加する。

```javascript
function 受注明細_現在確保を行配分_(rows, totalSecured){
  let rest=Math.max(0,Number(totalSecured)||0);
  const out=(rows||[]).map(item=>{
    const source=item||{}, rowNumber=Number(source.rowNumber)||0;
    if(source.発送済み) return {確保済み:'',不足:'',現在表示:false,発送済み:true};
    const qty=Number(source.qty);
    if(!Number.isFinite(qty) || qty<0)
      throw new Error((rowNumber?rowNumber+'行目 ':'')+'注文数量が数値ではありません: '+String(source.qty));
    const secured=Math.min(qty,rest);
    rest=Math.max(0,rest-secured);
    return {確保済み:secured,不足:Math.max(0,qty-secured),現在表示:secured>0,発送済み:false};
  });
  return {rows:out,未配分:rest};
}
```

- [ ] **Step 4: 候補へ現在EMS番号を付与する**

`Project_24/取り置き台帳.js` の `c.確保内訳=...` の直後へ追加する。

```javascript
    c.現在EMS=Array.from(new Set((自動確保行[key]||[])
      .map(r=>String(r.元EMS番号||'').trim())
      .filter(v=>typeof 実EMS番号_!=='function'||実EMS番号_(v)))).join(', ');
```

これは表示用の一時プロパティであり、台帳へ保存しない。

- [ ] **Step 5: 受注明細の書き戻しをグループ配分へ変更する**

`受注明細_確保列を書く_` を次の実装へ置き換える。

```javascript
function 受注明細_確保列を書く_(candidates){
  const ss=SpreadsheetApp.getActive(), recv=ss.getSheetByName(HIKIATE_CFG.受注);
  if(!recv) return;
  const M=列マップ_(recv), last=recv.getLastRow();
  if(last<=M.hr) return;
  const cols=受注明細_確保列を用意_(recv), emsCol=EMS番号列を用意_(recv);
  const byKey={};
  (candidates||[]).forEach(c=>{
    if(String(c.判定||'')==='棚戻し待ち') return;
    const k=取り置き_行キー_(c);
    if(!(k in byKey)) byKey[k]=c;
  });
  const R=recv.getRange(M.hr+1,1,last-M.hr,recv.getLastColumn()).getValues();
  const 確保out=R.map(()=>['']),不足out=R.map(()=>['']),内訳out=R.map(()=>['']);
  const EMSout=recv.getRange(M.hr+1,emsCol,R.length,1).getValues();
  const groups={};
  R.forEach((row,index)=>{
    const ban=String(row[M.番号]||'').trim();
    if(!ban) return;
    const key=取り置き_行キー_({ban,sku:M.SKU>=0?row[M.SKU]:'',code:row[M.コード]});
    if(!byKey[key]) return;
    (groups[key]=groups[key]||[]).push({
      index,
      qty:M.個数>=0?row[M.個数]:'',
      発送済み:引当_行出荷済み_(row,M),
      rowNumber:M.hr+1+index
    });
  });
  Object.keys(groups).forEach(key=>{
    const c=byKey[key], sourceRows=groups[key];
    const result=受注明細_現在確保を行配分_(sourceRows,c.確保済み);
    const activeCount=sourceRows.filter(r=>!r.発送済み).length;
    result.rows.forEach((line,pos)=>{
      const index=sourceRows[pos].index;
      if(line.発送済み){ EMSout[index][0]=''; return; }
      確保out[index][0]=line.確保済み;
      不足out[index][0]=line.不足;
      if(!line.現在表示){ EMSout[index][0]=''; return; }
      const detail=String(c.確保内訳||'');
      内訳out[index][0]=activeCount>1&&detail
        ? detail+'［この行'+line.確保済み+'／SKU合計'+Math.min(Number(c.確保済み)||0,sourceRows.filter(r=>!r.発送済み).reduce((s,r)=>s+Number(r.qty),0))+'］'
        : detail;
      EMSout[index][0]=String(c.現在EMS||'');
    });
    if(result.未配分>0) console.warn('受注明細の表示上限超過: '+key+' 未配分'+result.未配分);
  });
  recv.getRange(M.hr+1,cols.確保済み,確保out.length,1).setValues(確保out);
  recv.getRange(M.hr+1,cols.不足,不足out.length,1).setValues(不足out);
  recv.getRange(M.hr+1,cols.確保内訳,内訳out.length,1).setValues(内訳out);
  recv.getRange(M.hr+1,emsCol,EMSout.length,1).setValues(EMSout);
}
```

- [ ] **Step 6: 行別配分テストと構文検査を通す**

Run:

```powershell
node tests/project24_torioki.test.js
node --check "Project_24\引当.js"
node --check "Project_24\取り置き台帳.js"
```

Expected: `ALL TESTS PASSED`、構文検査はいずれも終了コード0。

- [ ] **Step 7: Task 1をコミットする**

```powershell
git add -- Project_24/引当.js Project_24/取り置き台帳.js tests/project24_torioki.test.js
git commit -m "fix(Project_24): 分割発送行の現在確保表示を整合"
```

### Task 2: 取り置き登録の入力規則を一括設定する

**Files:**
- Modify: `Project_24/取り置き台帳.js:987-1150`
- Modify: `tests/project24_torioki.test.js:17-23`
- Test: `tests/project24_torioki.test.js`

**Interfaces:**
- Consumes: candidates with field `判定`, `TORIOKI_棚確認`, `TORIOKI_戻し処理`
- Produces: `取り置き_入力規則計画_(candidates, rowCount, shelfRule, returnRule)` returning `{棚確認, 処理}` matrices

- [ ] **Step 1: テストで取り置き台帳ロジックを読み込む**

`tests/project24_torioki.test.js` の読み込み配列へ追加する。

```javascript
  'Project_24/引当.js',
  'Project_24/取り置き台帳.js',
  'Project_24/P列自動記入.js',
```

- [ ] **Step 2: 入力規則計画の失敗テストを追加する**

```javascript
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
```

- [ ] **Step 3: テストが未定義で失敗することを確認する**

Run:

```powershell
node tests/project24_torioki.test.js
```

Expected: `取り置き_入力規則計画_ is not a function` を含むFAIL。

- [ ] **Step 4: 入力規則計画の純粋関数を追加する**

`Project_24/取り置き台帳.js` の `取り置き初期登録を作成本体_` より前へ追加する。

```javascript
function 取り置き_入力規則計画_(candidates,rowCount,棚規則,戻し規則){
  const rows=candidates||[], count=Math.max(rows.length,Math.max(0,Number(rowCount)||0));
  const 棚確認=[],処理=[];
  for(let i=0;i<count;i++){
    const c=rows[i], judge=String(c&&c.判定||'');
    棚確認.push([c&&judge!=='棚戻し待ち'&&judge!=='即納'?棚規則:null]);
    処理.push([c&&judge==='棚戻し待ち'?戻し規則:null]);
  }
  return {棚確認,処理};
}
```

- [ ] **Step 5: 旧管理行数を保存前に取得する**

`取り置き初期登録を作成本体_` の開始時へ計測開始を追加する。

```javascript
  const 処理開始ms=Date.now();
```

`取り置き_表を保存_` の直前で旧行数と計測値を取得する。

```javascript
  const 旧登録シート=ss.getSheetByName(TORIOKI_CFG.初期);
  const 旧管理行数=旧登録シート?Math.max(0,旧登録シート.getLastRow()-TORIOKI_CFG.初期HDR行):0;
  const データ収集ms=Date.now()-処理開始ms;
  const 値書込開始ms=Date.now();
```

- [ ] **Step 6: 全シート・行単位の入力規則操作を一括設定へ置き換える**

現在の `sh2.getRange(1,1,...).clearDataValidations()` から候補行 `forEach` までを次へ置き換える。

```javascript
  const 値書込ms=Date.now()-値書込開始ms;
  const 入力規則開始ms=Date.now();
  const 棚確認列=TORIOKI_CFG.初期HDR.indexOf('棚確認')+1, 処理列=TORIOKI_CFG.初期HDR.indexOf('処理')+1;
  const 規則行数=Math.max(candidates.length,旧管理行数);
  if(規則行数>0){
    const 棚規則=SpreadsheetApp.newDataValidation().requireValueInList(TORIOKI_棚確認.slice(),true).setAllowInvalid(true).build();
    const 戻し規則=SpreadsheetApp.newDataValidation().requireValueInList(TORIOKI_戻し処理.slice(),true).setAllowInvalid(false).build();
    const plan=取り置き_入力規則計画_(candidates,規則行数,棚規則,戻し規則);
    sh2.getRange(初期データ行,1,規則行数,TORIOKI_CFG.初期HDR.length).clearDataValidations();
    sh2.getRange(初期データ行,棚確認列,規則行数,1).setDataValidations(plan.棚確認);
    sh2.getRange(初期データ行,処理列,規則行数,1).setDataValidations(plan.処理);
  }
  const 入力規則ms=Date.now()-入力規則開始ms;
```

行単位の `getRange()`、`clearDataValidations()`、`setDataValidation()` は残さない。

- [ ] **Step 7: 書式と合計時間のログを追加する**

`取り置き_登録行書式を更新_` の直前・直後を計測し、直後へログを追加する。

```javascript
  const 書式開始ms=Date.now();
  取り置き_登録行書式を更新_(sh2,candidates);
  const 書式ms=Date.now()-書式開始ms;
  const 合計ms=Date.now()-処理開始ms;
  console.log('取り置き登録更新 処理時間ms='+合計ms
    +' データ収集='+データ収集ms
    +' 値書込='+値書込ms
    +' 入力規則='+入力規則ms
    +' 書式='+書式ms);
```

- [ ] **Step 8: 一括入力規則テストと構文検査を通す**

Run:

```powershell
node tests/project24_torioki.test.js
node --check "Project_24\取り置き台帳.js"
```

Expected: `ALL TESTS PASSED`、構文検査は終了コード0。

- [ ] **Step 9: Task 2をコミットする**

```powershell
git add -- Project_24/取り置き台帳.js tests/project24_torioki.test.js
git commit -m "perf(Project_24): 取り置き入力規則を一括設定"
```

### Task 3: 回帰検証・安全反映・実シート確認

**Files:**
- Verify: `Project_24/引当.js`
- Verify: `Project_24/取り置き台帳.js`
- Verify: `tests/project24_torioki.test.js`

**Interfaces:**
- Consumes: Task 1とTask 2の実装
- Produces: Project_24オンライン版と実シートの検証結果

- [ ] **Step 1: 関連テストをすべて実行する**

Run:

```powershell
node tests/project24_torioki.test.js
node tests/project24_zenken.test.js
node tests/project24_daniel.test.js
node --check "Project_24\引当.js"
node --check "Project_24\取り置き台帳.js"
```

Expected: 3テストが `ALL TESTS PASSED`、構文検査が終了コード0。

- [ ] **Step 2: 対象差分と作業ツリーを確認する**

Run:

```powershell
git diff HEAD~2 -- Project_24/引当.js Project_24/取り置き台帳.js tests/project24_torioki.test.js
git status --short --branch
```

Expected: 対象3ファイル以外をコミットしていない。`../claude自動引当ツール/reports/` の既存変更は未変更のまま残る。

- [ ] **Step 3: Project_24を安全pushする**

Run:

```powershell
powershell -ExecutionPolicy Bypass -File tools\gas_safe_push.ps1 Project_24
```

Expected: オンライン競合なしでpush完了。競合があれば上書きせず停止する。

- [ ] **Step 4: 実シートで取り置き登録単独更新を実行する**

Google Sheetsのメニューから `📋 取り置き登録を更新` を1回実行する。

Expected:

- 6分以内に完了する。
- 目標は120秒以内。
- お取り置きメモと自動色分けが更新前後で残る。
- 棚確認と返品処理のプルダウンが従来どおり選択できる。

- [ ] **Step 5: 受注10117477と数量正本を確認する**

Expected:

- `10117477 | MRBLUE41` の発送済み2個行: 確保済み・不足・確保内訳・EMS番号が空欄。
- 同キーの未発送1個行: 確保済み1、不足0、EMS番号 `EG050152967KR`。
- 発注共有P列: `10117477:1`。
- 引当台帳: 現在有効1個、手動解放2個、発送済み履歴3個のまま。

- [ ] **Step 6: 完了状態を記録する**

実行時間と10117477の確認結果を最終報告へ記載する。実シート更新が6分を超えた場合は、数量修正を追加せず、計測ログの区間別時間を報告して非同期分割を別設計にする。
