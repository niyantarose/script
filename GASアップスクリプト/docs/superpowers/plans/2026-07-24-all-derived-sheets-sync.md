# 数量変更後の全派生シート同期 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Project_24で数量を変更する全ての日常操作の終了時に、②引当を1回だけ実行して全派生シートを同期し、Yahoo CSVは日本在庫確認後の手動ボタン出力に限定する。

**Architecture:** 新しい `Project_24/全シート同期.js` に中央同期オーケストレーターを置き、低レベル保存関数ではなく利用者起点の最上位処理からだけ呼ぶ。②本体は計算と主要出力を担当し、中央同期が管理シート更新と整合状態の確定・失敗通知を担当する。⑤便締めはCSVを生成せず、手動出力記録が現在の対象数量を含むことを検証してから状態を変更する。

**Tech Stack:** Google Apps Script V8、Node.js `assert`/`vm` テスト、clasp、安全pushスクリプト

## Global Constraints

- 正データはEMSリスト、取り置き台帳、EMS在庫移動台帳、GoQ受注明細・発送実績、別ルート入荷情報である。
- 分類シートや引当状況一覧の前回値を計算入力に使わない。
- 1操作につき②引当は最大1回。
- `取り置き台帳_保存_` と `EMS在庫移動台帳_保存_` から同期を呼ばない。
- DocumentLockを入れ子にしない。
- 全件検算、全件再計算プレビュー、バックアップ、Yahoo CSV作成、GoQステータス変更は自動同期しない。
- Yahoo CSVは全同期後の日本在庫を利用者が確認し、既存ボタンを押した時だけ作成する。
- 既存の取り置きメモ、棚確認、未反映の追加数量・マイナス数量、自動色分けを同期で消さない。

---

### Task 1: 中央同期オーケストレーター

**Files:**
- Create: `Project_24/全シート同期.js`
- Create: `tests/project24_all_sheet_sync.test.js`
- Modify: `Project_24/引当.js:1713-1714, 2108-2130`

**Interfaces:**
- Consumes: `EMS在庫を更新_本体_()`, `引当実行_本体_(options)`, `取り置き初期登録を作成本体_(options)`, `キャンセル戻し確認を更新本体_()`, `Yahoo戻し候補を更新_()`
- Produces: `引当_数値変更後全同期_(options)`, `引当_同期管理シート_(options)`, `引当_同期状態_読む_()`, `引当_同期状態_保存_(state)`

- [ ] **Step 1: 中央同期の失敗テストを書く**

`tests/project24_all_sheet_sync.test.js` にVMコンテキストを作り、次を検証する。

```javascript
test('中央同期はEMS更新→②→管理シートの順に各1回実行する', () => {
  const calls=[];
  const ctx=loadSyncContext({
    EMS在庫を更新_本体_:()=>calls.push('ems'),
    引当実行_本体_:()=>{ calls.push('allocate'); return {success:true}; },
    取り置き初期登録を作成本体_:()=>calls.push('register'),
    キャンセル戻し確認を更新本体_:()=>calls.push('return'),
    Yahoo戻し候補を更新_:()=>calls.push('yahoo')
  });
  const result=ctx.引当_数値変更後全同期_({理由:'test',EMS更新:true,完了表示:false});
  assert.strictEqual(result.success,true);
  assert.deepStrictEqual(calls,['ems','allocate','register','return','yahoo']);
});

test('管理シート失敗時は整合状態を削除して同期済みにしない', () => {
  const ctx=loadSyncContext({
    引当実行_本体_:()=>({success:true}),
    取り置き初期登録を作成本体_:()=>{ throw new Error('register failed'); }
  });
  const result=ctx.引当_数値変更後全同期_({理由:'test',完了表示:false});
  assert.strictEqual(result.success,false);
  assert.strictEqual(ctx.__props.has('引当_整合状態'),false);
  assert.strictEqual(JSON.parse(ctx.__props.get('引当_全シート同期状態')).status,'failed');
});
```

- [ ] **Step 2: テストを実行してREDを確認する**

Run: `node tests/project24_all_sheet_sync.test.js`

Expected: FAIL because `Project_24/全シート同期.js` and `引当_数値変更後全同期_` do not exist.

- [ ] **Step 3: 最小の中央同期を実装する**

`Project_24/全シート同期.js` に次の責務を持たせる。

```javascript
const HIKIATE_SYNC_STATE_KEY='引当_全シート同期状態';

function 引当_同期状態_保存_(state){
  PropertiesService.getDocumentProperties().setProperty(HIKIATE_SYNC_STATE_KEY,JSON.stringify(state||{}));
}

function 引当_同期状態_読む_(){
  try{ return JSON.parse(PropertiesService.getDocumentProperties().getProperty(HIKIATE_SYNC_STATE_KEY)||'null'); }
  catch(e){ return null; }
}

function 引当_同期管理シート_(){
  if(typeof 取り置き初期登録を作成本体_==='function') 取り置き初期登録を作成本体_({silent:true});
  if(typeof キャンセル戻し確認を更新本体_==='function') キャンセル戻し確認を更新本体_();
  if(typeof Yahoo戻し候補を更新_==='function') Yahoo戻し候補を更新_();
}

function 引当_数値変更後全同期_(options){
  const opt=options||{}, reason=String(opt.理由||'数量変更'), started=Date.now();
  const props=PropertiesService.getDocumentProperties();
  props.deleteProperty('引当_整合状態');
  引当_同期状態_保存_({status:'running',reason,startedAt:started});
  try{
    if(opt.EMS更新) EMS在庫を更新_本体_();
    const result=引当実行_本体_({preview:false,silentSummary:true,skipManagementRefresh:true});
    if(!result || result.success!==true) throw new Error('②引当が完了結果を返しませんでした');
    引当_同期管理シート_();
    const integrity=JSON.parse(props.getProperty('引当_整合状態')||'null');
    const ledgerAt=Number(props.getProperty('取り置き台帳_最終更新')||0);
    if(!integrity || Number(integrity.ts||0)<ledgerAt) throw new Error('台帳更新後の整合時刻を確認できません');
    引当_同期状態_保存_({status:'synced',reason,startedAt:started,finishedAt:Date.now()});
    return {success:true,result};
  }catch(error){
    props.deleteProperty('引当_整合状態');
    引当_同期状態_保存_({status:'failed',reason,startedAt:started,failedAt:Date.now(),error:String(error&&error.message||error)});
    if(opt.完了表示!==false) SpreadsheetApp.getUi().alert('全シート同期に失敗しました',
      reason+'の元データは保存済みですが、派生シートは未同期です。\n'+String(error&&error.message||error)+'\n\nメニューから②を実行し直してください。',SpreadsheetApp.getUi().ButtonSet.OK);
    return {success:false,error};
  }
}
```

`引当.js` では、公開 `引当実行()` を中央同期へ接続する。`引当実行_本体_` の末尾にある取り置き登録更新は `skipManagementRefresh` が偽の時だけ `引当_同期管理シート_()` を呼ぶ形へ置き換え、全件再計算など既存の直接呼出しも全管理シートを更新できるようにする。

- [ ] **Step 4: 中央同期テストをGREENにする**

Run: `node tests/project24_all_sheet_sync.test.js`

Expected: all central-sync tests PASS.

- [ ] **Step 5: コミットする**

```bash
git add Project_24/全シート同期.js Project_24/引当.js tests/project24_all_sheet_sync.test.js
git commit -m "feat(Project_24): 数量変更後の全派生シート同期を追加"
```

---

### Task 2: 数量変更入口を中央同期へ接続

**Files:**
- Modify: `Project_24/取り置き台帳.js:1194-1244, 1257-1265, 1371-1435, 1475-1491, 1536-1544`
- Modify: `Project_24/受注明細個別ボタン.js:180-300`
- Modify: `Project_24/現物確認移行.js:176-222`
- Modify: `Project_24/消込台帳.js:227-303, 374-405`
- Modify: `Project_24/引当.js:469-685`
- Modify: `tests/project24_all_sheet_sync.test.js`

**Interfaces:**
- Consumes: `引当_数値変更後全同期_({理由,EMS更新,完了表示})`
- Produces: 各最上位操作が数量保存後に同期をちょうど1回呼ぶ契約

- [ ] **Step 1: 接続契約の失敗テストを書く**

ソースから関数本体を取り出す `functionBody(source,name)` をテストへ追加し、次を検証する。

```javascript
[
  ['取り置き台帳.js','取り置き登録を反映本体_'],
  ['取り置き台帳.js','キャンセル戻し確認を確定本体_'],
  ['取り置き台帳.js','キャンセル戻しをYahoo反映済みにする本体_'],
  ['取り置き台帳.js','選択した取り置きを手動解除本体_'],
  ['取り置き台帳.js','orphanBulkRelease'],
  ['受注明細個別ボタン.js','選択行を個別引当_本体_'],
  ['受注明細個別ボタン.js','選択行の引当キャンセル_本体_'],
  ['現物確認移行.js','現物確認移行を反映本体_'],
  ['消込台帳.js','注文をキャンセル扱い本体_'],
  ['引当.js','取込_実行_']
].forEach(([file,name])=>{
  test(name+'は数量変更後に全同期を1回呼ぶ',()=>{
    const body=functionBody(read(file),name);
    assert.strictEqual((body.match(/引当_数値変更後全同期_\s*\(/g)||[]).length,1);
  });
});
```

- [ ] **Step 2: 接続テストを実行してREDを確認する**

Run: `node tests/project24_all_sheet_sync.test.js`

Expected: FAIL with zero coordinator calls in the listed entrypoints.

- [ ] **Step 3: 各入口を最小変更で接続する**

- 取り置き登録反映: 既存の直接 `引当実行_本体_` を中央同期へ置換する。
- キャンセル戻し確認、Yahoo戻し確定、手動解除、孤児解除、現物移行: 台帳保存と監査書込みが全て完了した後に同期する。
- 個別引当・解除: ループ内では同期せず `変更あり` を立て、ループ終了後に1回だけ同期する。
- 手動キャンセル: `キャンセル処理_` の戻り値へ `取り置き更新` を追加し、P列除去・台帳更新・消込更新のいずれかがあれば同期する。CSV取込からの内部呼出しは同期せず、外側の `取込_実行_` が最後に1回同期する。
- GoQ CSV取込: 受注明細、台帳遷移、消込、キャンセル証跡の保存後に同期する。従来の取り置き登録だけの更新は削除し、同期結果を完了表示へ含める。
- 同期失敗時は中央同期が整合状態を無効化して警告する。保存済み入力を自動で戻さない。

- [ ] **Step 4: 接続テストと既存ワークフローテストをGREENにする**

Run:

```bash
node tests/project24_all_sheet_sync.test.js
node tests/project24_reservation_workflow.test.js
node tests/project24_physical_migration.test.js
```

Expected: all PASS.

- [ ] **Step 5: コミットする**

```bash
git add Project_24/取り置き台帳.js Project_24/受注明細個別ボタン.js Project_24/現物確認移行.js Project_24/消込台帳.js Project_24/引当.js tests/project24_all_sheet_sync.test.js
git commit -m "feat(Project_24): 全ての数量変更入口で派生シートを同期"
```

---

### Task 3: Yahoo CSVを手動出力だけに限定して便締めを同期

**Files:**
- Modify: `Project_24/Yahoo在庫変更出力.js:85-185`
- Modify: `Project_24/引当履歴.js:360-476`
- Modify: `tests/project24_yahoo_stock_export.test.js`
- Modify: `tests/project24_all_sheet_sync.test.js`

**Interfaces:**
- Consumes: 日本在庫の対象行、`YAHOO出力記録`
- Produces: `Yahoo変更_出力記録が対象を含む_(targets,record)`, 手動出力記録 `items`, 便締め後の `引当_数値変更後全同期_({EMS更新:true})`

- [ ] **Step 1: 手動出力契約の失敗テストを書く**

```javascript
test('全便手動出力記録は対象便の行を数量一致で包含できる',()=>{
  const targets=[{EMS番号:'EG1',商品コード:'AAA',余り数:2}];
  const record={items:[
    {EMS番号:'EG1',商品コード:'AAA',余り数:2},
    {EMS番号:'EG2',商品コード:'BBB',余り数:1}
  ]};
  assert.strictEqual(context.Yahoo変更_出力記録が対象を含む_(targets,record),true);
  record.items[0].余り数=1;
  assert.strictEqual(context.Yahoo変更_出力記録が対象を含む_(targets,record),false);
});

test('便締めはYahoo CSVを自動生成しない',()=>{
  const body=functionBody(read('引当履歴.js'),'到着済を在庫反映済みへ本体_');
  assert.doesNotMatch(body,/Yahoo在庫変更を出力本体_\s*\(/);
  assert.match(body,/Yahoo変更_出力記録が対象を含む_/);
  assert.match(body,/引当_数値変更後全同期_\s*\(\s*\{[^}]*EMS更新\s*:\s*true/);
});
```

- [ ] **Step 2: テストを実行してREDを確認する**

Run:

```bash
node tests/project24_yahoo_stock_export.test.js
node tests/project24_all_sheet_sync.test.js
```

Expected: FAIL because output records do not contain `items`, the containment helper is absent, and ⑤ still auto-calls the exporter.

- [ ] **Step 3: 手動出力記録と便締めガードを実装する**

`Yahoo在庫変更出力.js`:

- `YAHOO出力記録` に従来の `sig`, `at`, `要確認コード` と、正規化前の `{EMS番号,商品コード,余り数}` 配列 `items` を保存する。
- `Yahoo変更_出力記録が対象を含む_` は対象行を `EMS番号|商品コード|数量` で比較する。
- 旧記録は対象全体の `sig` が完全一致する場合だけ互換扱いする。
- `日本在庫CSVを作成` と既存の出力ボタンだけが `Yahoo在庫変更を出力本体_` を呼ぶ。

`引当履歴.js`:

- ⑤内部の `Yahoo在庫変更を出力本体_(targets)` 自動呼出しを削除する。
- Yahoo出力対象があるのに現在の記録が対象行を含まない場合は、副作用を起こす前に便締めを中止し、「日本在庫を確認してCSV作成ボタンを押す」案内を出す。
- 出力済みなら現行どおり要確認・対象外を移動対象から除く。
- EMS状態・移動台帳の保存後に `引当_数値変更後全同期_({理由:'便締め',EMS更新:true,完了表示:false})` を呼ぶ。
- 同期失敗時は便締め自体を成功表示せず、中央同期の未同期警告を表示する。

- [ ] **Step 4: Yahoo・便締めテストをGREENにする**

Run:

```bash
node tests/project24_yahoo_stock_export.test.js
node tests/project24_all_sheet_sync.test.js
node tests/project24_physical_operation_guard.test.js
node tests/project24_reservation_workflow.test.js
```

Expected: all PASS.

- [ ] **Step 5: コミットする**

```bash
git add Project_24/Yahoo在庫変更出力.js Project_24/引当履歴.js tests/project24_yahoo_stock_export.test.js tests/project24_all_sheet_sync.test.js
git commit -m "fix(Project_24): Yahoo CSVを日本在庫確認後の手動出力に限定"
```

---

### Task 4: 全回帰検証・反映・引継ぎ

**Files:**
- Modify: `HANDOFF.md`
- Modify if needed: `docs/superpowers/plans/2026-07-24-all-derived-sheets-sync.md`

**Interfaces:**
- Consumes: Tasks 1-3の実装と全Project_24テスト
- Produces: 検証済みコミット、安全push済みProject_24、Claude Code用引継ぎ

- [ ] **Step 1: Project_24全テストを実行する**

Run:

```powershell
$files = Get-ChildItem tests -Filter 'project24*.test.js' | Sort-Object Name
$failed = 0
foreach ($file in $files) {
  node $file.FullName
  if ($LASTEXITCODE -ne 0) { $failed++ }
}
"TOTAL=$($files.Count) FAILED=$failed"
if ($failed -gt 0) { exit 1 }
```

Expected: `FAILED=0`.

- [ ] **Step 2: 変更JavaScriptの構文を確認する**

Run:

```powershell
node --check Project_24/全シート同期.js
node --check Project_24/引当.js
node --check Project_24/取り置き台帳.js
node --check Project_24/受注明細個別ボタン.js
node --check Project_24/現物確認移行.js
node --check Project_24/消込台帳.js
node --check Project_24/Yahoo在庫変更出力.js
node --check Project_24/引当履歴.js
```

Expected: all commands exit 0.

- [ ] **Step 3: 差分と危険な自動CSV呼出しが無いことを確認する**

Run:

```bash
git diff --check
rg -n "Yahoo在庫変更を出力本体_\(targets\)" Project_24
git diff --stat
```

Expected: `git diff --check` is clean and `rg` finds no ⑤ auto-output call.

- [ ] **Step 4: HANDOFF.mdを更新する**

記載内容:

- 中央同期関数と対象入口
- Yahoo CSVの新しい手動運用
- テスト結果
- 安全push結果
- 未解決事項があれば具体的なファイル・関数・再現手順

- [ ] **Step 5: 最終コミットする**

```bash
git add HANDOFF.md docs/superpowers/plans/2026-07-24-all-derived-sheets-sync.md
git commit -m "docs(Project_24): 全シート同期の実装結果を引き継ぎ"
```

- [ ] **Step 6: オンライン差分を再同期して安全pushする**

Run:

```powershell
powershell -ExecutionPolicy Bypass -File tools\gas_pull_sync.ps1 Project_24
powershell -ExecutionPolicy Bypass -File tools\gas_safe_push.ps1 Project_24
```

Expected: 共同作業者の変更を保護し、Project_24だけが安全にpushされる。
