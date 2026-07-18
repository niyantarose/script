# 引当取り置き台帳 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 注文番号ごとの現物取り置きを永続台帳で管理し、過去の入荷日・出荷日推測ではなく、取り置き台帳と実EMS数量だけで引当、キャンセル戻し、Yahoo移動、全件検算を一致させる。

**Architecture:** SpreadsheetAppに依存しない純粋計算を `取り置き計算.js` に分離し、シート作成・初期登録・状態遷移・一括保存を `取り置き台帳.js` に集約する。①は明示的なCSVステータスを台帳状態へ反映し、④は「計画作成→全量検証→P列・台帳・出力の一括反映」の順で処理する。⑤はYahooへ実際に移した数量だけを決定的な処理IDで `EMS在庫移動台帳` に記録する。

**Tech Stack:** Google Apps Script V8、SpreadsheetApp/PropertiesService、Node.js標準 `assert` + `vm`、PowerShell、Project_24安全同期スクリプト。

## Global Constraints

- 注文別確保の正は `取り置き台帳` の状態 `取り置き中` だけとする。
- Yahoo即納在庫は比較材料であり、韓国取り寄せ注文の供給へ加えない。
- 実EMS番号を持つ到着済み行だけをEMS供給にする。EMS番号空欄、`棚卸` 行、未着行は供給にしない。
- 受注明細の入荷日、既存 `取り置き数`、消込台帳メモ、引当履歴を注文別確保数量として加算しない。
- `OFW304-1`、`OFW304-2`、`OFW305-1`、`OFW305-2` は別商品として扱い、数値枝番を削除しない。
- 販売SKU末尾の `a` / `b` だけを基底照合で外し、連結コードは推測分割しない。
- `POEM65`、`RECIPE42`、商品コード欄が受注番号の在庫は通常FIFOへ混ぜない。
- 利用者向け状態は `取り置き中`、`発送済み`、`キャンセル戻し`、`手動解除` とする。
- キャンセル戻し結果は `未確認`、`現物あり`、`再引当済み`、`Yahoo反映済み`、`在庫なし` とする。
- `在庫なし` は再利用できない確定差引きとして箱別突合へ残す。
- 検証エラー時はP列、取り置き台帳、受注明細、引当履歴、出力シートへ書き込まない。
- 同じ初期登録、④、⑤を再実行しても数量を重複追加しない。
- 作業開始時に `git pull --ff-only` と `tools\gas_pull_sync.ps1 Project_24` を実行する。
- 反映は `tools\gas_safe_push.ps1 Project_24` だけを使い、素の `clasp push` は使用しない。
- Project_24、対応テスト、承認済み設計書、本計画以外を同じコミットへ含めない。

## File Structure

| ファイル | 責務 |
| --- | --- |
| `Project_24/取り置き計算.js` | キー正規化、台帳集計、必要数、供給残、再引当、FIFO、箱別不変条件の純粋計算 |
| `Project_24/取り置き台帳.js` | 台帳シートI/O、初期登録、CSV状態遷移、キャンセル確認、Yahoo移動記録、メニュー公開関数 |
| `Project_24/引当.js` | ①と④のオーケストレーション、既存注文分類・表示への新計算結果接続 |
| `Project_24/P列自動記入.js` | P列を直接書く処理をプレビューと反映へ分離し、台帳確保を固定表示する |
| `Project_24/消込台帳.js` | 既存キャンセル操作から新台帳遷移を呼ぶ互換入口 |
| `Project_24/引当履歴.js` | ⑤の便締め時にEMS在庫移動台帳を確定してから既存履歴を維持する |
| `Project_24/全件検算.js` | EMS供給と新台帳・移動台帳を同じ式で検算する |
| `tests/project24_reservation_domain.test.js` | キー、状態、必要数、供給残、優先順位、実例回帰の純粋テスト |
| `tests/project24_reservation_workflow.test.js` | CSV遷移、初期登録、P列プレビュー、⑤冪等性のワークフローテスト |
| `tests/project24_torioki.test.js` | 旧 `取り置き数` 依存を外し、新台帳数量へ切り替えた互換テスト |
| `tests/project24_zenken_kensan.test.js` | 新しい箱別突合式とYahoo比較の回帰テスト |

---

### Task 1: 取り置き台帳の純粋ドメインモデルを固定する

**Files:**
- Create: `Project_24/取り置き計算.js`
- Create: `tests/project24_reservation_domain.test.js`
- Reference: `Project_24/引当.js` functions `normCode_`, `受注候補コード_`, `codeKeys_`

**Interfaces:**
- Consumes: `normCode_(value)`。
- Produces: `取り置き_商品コード_(sku, code)`, `取り置き_行キー_(order)`, `取り置き_供給キー_(ems, code)`, `取り置き_集計_(rows, movements)`, `取り置き_今回必要数_(order, summary)`。

- [ ] **Step 1: 作業開始時の同期と差分確認を行う**

Run:

```powershell
git pull --ff-only
powershell -ExecutionPolicy Bypass -File tools\gas_pull_sync.ps1 Project_24
git status --short
```

Expected: `Project_24 はオンラインとローカルが一致しています。` と表示され、今回の設計書・計画書以外の未コミット変更がない。別変更があれば作業を止め、所有者を確認する。

- [ ] **Step 2: キー・状態・必要数の失敗テストを作成する**

Create `tests/project24_reservation_domain.test.js` with the standard `assert` + `vm` harness used by existing Project_24 tests, load `Project_24/引当.js` then `Project_24/取り置き計算.js`, and add these exact assertions:

```javascript
test('数値枝番は別商品、販売末尾a/bだけを基底へ寄せる', () => {
  const orders = [
    {ban:'1', code:'OFW304-1', sku:'OFW304-1b'},
    {ban:'1', code:'OFW304-2', sku:'OFW304-2b'},
    {ban:'1', code:'OFW305-1', sku:'OFW305-1b'},
    {ban:'1', code:'OFW305-2', sku:'OFW305-2b'}
  ];
  const codes = orders.map(o => context.取り置き_商品コード_(o.sku, o.code));
  assert.strictEqual(JSON.stringify(codes), JSON.stringify(['OFW304-1','OFW304-2','OFW305-1','OFW305-2']));
  assert.strictEqual(new Set(codes).size, 4);
  assert.strictEqual(context.取り置き_商品コード_('', 'POEM65（10116569）'), 'POEM65');
});

test('取り置き中だけが注文必要数を減らす', () => {
  const key = context.取り置き_行キー_({ban:'101', code:'AAA-1', sku:'AAA-1b'});
  const rows = [
    {取置ID:'A', 状態:'取り置き中', 受注番号:'101', 商品コード:'AAA-1', SKU:'AAA-1b', 取り置き数量:2, 取置元種別:'開始前在庫'},
    {取置ID:'B', 状態:'発送済み', 受注番号:'101', 商品コード:'AAA-1', SKU:'AAA-1b', 取り置き数量:1, 取置元種別:'EMS', 元EMS番号:'EG1'},
    {取置ID:'C', 状態:'手動解除', 受注番号:'101', 商品コード:'AAA-1', SKU:'AAA-1b', 取り置き数量:1, 取置元種別:'開始前在庫'}
  ];
  const summary = context.取り置き_集計_(rows, []);
  assert.strictEqual(summary.activeByKey[key], 2);
  assert.strictEqual(context.取り置き_今回必要数_({ban:'101',code:'AAA-1',sku:'AAA-1b',qty:3}, summary), 1);
});

test('キャンセル戻し結果を供給使用区分へ正しく分類する', () => {
  const base = {状態:'キャンセル戻し',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:1,取置元種別:'EMS',元EMS番号:'EG1'};
  const rows = [
    Object.assign({取置ID:'A',戻し処理結果:'未確認'},base),
    Object.assign({取置ID:'B',戻し処理結果:'現物あり'},base),
    Object.assign({取置ID:'C',戻し処理結果:'在庫なし'},base),
    Object.assign({取置ID:'D',戻し処理結果:'再引当済み'},base),
    Object.assign({取置ID:'E',戻し処理結果:'Yahoo反映済み'},base)
  ];
  const summary = context.取り置き_集計_(rows, [{処理ID:'YAHOO|RETURN|E',EMS番号:'EG1',商品コード:'AAA',数量:1}]);
  const use = summary.usageBySupply['EG1|AAA'];
  assert.strictEqual(JSON.stringify(use), JSON.stringify({取り置き中:0,発送済み:0,戻し未処理:2,在庫なし確定:1,Yahoo移動済み:1}));
  assert.strictEqual(summary.confirmedReturns.length, 1);
  assert.strictEqual(summary.confirmedReturns[0].取置ID, 'B');
});

test('重複ID・非整数・注文数量超過をエラーにする', () => {
  const rows = [
    {取置ID:'X',状態:'取り置き中',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:1.5},
    {取置ID:'X',状態:'取り置き中',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:2}
  ];
  const summary = context.取り置き_集計_(rows, []);
  assert.ok(summary.errors.some(e => /重複/.test(e)));
  assert.ok(summary.errors.some(e => /正の整数/.test(e)));
});
```

- [ ] **Step 3: 新テストが未定義関数で失敗することを確認する**

Run:

```powershell
node tests\project24_reservation_domain.test.js
```

Expected: `取り置き_商品コード_ is not a function` または同等の未定義エラーでFAILする。

- [ ] **Step 4: 純粋ドメイン関数を実装する**

Create `Project_24/取り置き計算.js` with these constants and function contracts. `取り置き_集計_` はシートへアクセスせず、入力配列を変更しない。

```javascript
const TORIOKI_STATUS = Object.freeze({
  ACTIVE:'取り置き中', SHIPPED:'発送済み', RETURN:'キャンセル戻し', RELEASED:'手動解除'
});
const TORIOKI_RETURN = Object.freeze({
  UNCHECKED:'未確認', PRESENT:'現物あり', REALLOCATED:'再引当済み', YAHOO:'Yahoo反映済み', MISSING:'在庫なし'
});

function 取り置き_整数_(value){
  const n=Number(value);
  return Number.isInteger(n) && n>0 ? n : 0;
}

function 取り置き_商品コード_(sku, code){
  const c=normCode_(code);
  if(c) return c;
  return normCode_(sku).replace(/[AB]$/,'');
}

function 取り置き_行キー_(order){
  return [
    String(order&&order.ban||order&&order.受注番号||'').trim(),
    取り置き_商品コード_(order&&order.sku||order&&order.SKU, order&&order.code||order&&order.商品コード),
    normCode_(order&&order.sku||order&&order.SKU)
  ].join('|');
}

function 取り置き_供給キー_(ems, code){
  return String(ems||'').trim()+'|'+normCode_(code);
}

function 取り置き_集計_(rows, movements){
  const out={activeByKey:{}, activeRowsByKey:{}, usageBySupply:{}, confirmedReturns:[], errors:[], rows:(rows||[]).map(r=>Object.assign({},r))};
  const ids=new Set(), moveIds=new Set();
  const usage=k=>out.usageBySupply[k]||(out.usageBySupply[k]={取り置き中:0,発送済み:0,戻し未処理:0,在庫なし確定:0,Yahoo移動済み:0});
  out.rows.forEach((r,index)=>{
    const id=String(r.取置ID||'').trim(), qty=取り置き_整数_(r.取り置き数量), key=取り置き_行キー_(r);
    if(!id) out.errors.push('台帳'+(index+2)+'行: 取置IDなし');
    else if(ids.has(id)) out.errors.push('台帳'+(index+2)+'行: 取置ID重複 '+id);
    else ids.add(id);
    if(!qty) out.errors.push('台帳'+(index+2)+'行: 取り置き数量は正の整数');
    if(r.状態===TORIOKI_STATUS.ACTIVE && qty){
      out.activeByKey[key]=(out.activeByKey[key]||0)+qty;
      (out.activeRowsByKey[key]=out.activeRowsByKey[key]||[]).push(r);
    }
    const ems=String(r.元EMS番号||'').trim();
    if(!ems || !qty || r.状態===TORIOKI_STATUS.RELEASED) return;
    const u=usage(取り置き_供給キー_(ems,r.元EMS商品コード||r.商品コード));
    if(r.状態===TORIOKI_STATUS.ACTIVE) u.取り置き中+=qty;
    else if(r.状態===TORIOKI_STATUS.SHIPPED) u.発送済み+=qty;
    else if(r.状態===TORIOKI_STATUS.RETURN){
      const result=String(r.戻し処理結果||TORIOKI_RETURN.UNCHECKED);
      if(result===TORIOKI_RETURN.UNCHECKED || result===TORIOKI_RETURN.PRESENT) u.戻し未処理+=qty;
      if(result===TORIOKI_RETURN.MISSING) u.在庫なし確定+=qty;
      if(result===TORIOKI_RETURN.PRESENT) out.confirmedReturns.push(r);
    }
  });
  (movements||[]).forEach((m,index)=>{
    const qty=取り置き_整数_(m.数量), id=String(m.処理ID||'').trim();
    if(!id || !qty) out.errors.push('移動台帳'+(index+2)+'行: 処理IDまたは数量が不正');
    else if(moveIds.has(id)) out.errors.push('移動台帳'+(index+2)+'行: 処理ID重複 '+id);
    else moveIds.add(id);
    if(qty) usage(取り置き_供給キー_(m.EMS番号,m.商品コード)).Yahoo移動済み+=qty;
  });
  return out;
}

function 取り置き_今回必要数_(order, summary){
  const qty=Math.max(0,Number(order&&order.qty)||0);
  return Math.max(0,qty-Number(summary&&summary.activeByKey[取り置き_行キー_(order)]||0));
}
```

- [ ] **Step 5: 新テストを通し、構文を検査する**

Run:

```powershell
node tests\project24_reservation_domain.test.js
node --check Project_24\取り置き計算.js
```

Expected: 全テストPASS、構文エラーなし。

- [ ] **Step 6: Task 1をコミットする**

```powershell
git add Project_24\取り置き計算.js tests\project24_reservation_domain.test.js
git commit -m "feat(hikiate): 取り置き台帳の純粋計算を追加"
```

---

---

### Task 2: キャンセル再引当・EMS供給残・FIFOを純粋計算にする

**Files:**
- Modify: `Project_24/取り置き計算.js`
- Modify: `tests/project24_reservation_domain.test.js`

**Interfaces:**
- Consumes: Task 1の `取り置き_集計_`, `取り置き_今回必要数_`, `取り置き_行キー_`, `取り置き_供給キー_`。
- Produces: `取り置き_割当計算_(input)` と `取り置き_割当検証_(plan)`。
- `input.orders`: `{ban, code, sku, qty, sortKey, i, keys, paid}` 配列。
- `input.supplies`: `{ems, code, qty, arrival, directBan}` 配列。
- `input.explicit`: `{ems, code, ban, qty}` 配列。
- `plan`: `{orders, newRows, returnUpdates, remainingBySupply, surplus, errors}`。

- [ ] **Step 1: 実例と優先順位の失敗テストを追加する**

Append these assertions:

```javascript
test('現物ありキャンセル戻しを今回EMSより先に最古注文へ再引当する', () => {
  const result = context.取り置き_割当計算_({
    orders:[
      {ban:'100',code:'AAA',sku:'AAAb',qty:1,sortKey:100,i:0,keys:['AAA']},
      {ban:'200',code:'AAA',sku:'AAAb',qty:1,sortKey:200,i:1,keys:['AAA']}
    ],
    ledger:[{取置ID:'OLD',状態:'キャンセル戻し',受注番号:'050',商品コード:'AAA',SKU:'AAAb',取り置き数量:1,取置元種別:'EMS',元EMS番号:'EG0',戻し処理結果:'現物あり'}],
    movements:[],
    supplies:[{ems:'EG1',code:'AAA',qty:1,arrival:'2026-07-12'}],
    explicit:[]
  });
  assert.strictEqual(result.newRows[0].受注番号, '100');
  assert.strictEqual(result.newRows[0].取置元種別, 'キャンセル再引当');
  assert.strictEqual(result.newRows[0].元取置ID, 'OLD');
  assert.strictEqual(result.newRows[1].受注番号, '200');
  assert.strictEqual(result.newRows[1].元EMS番号, 'EG1');
});

test('YMNGD09は3個供給・1個確保なら余り2個', () => {
  const result = context.取り置き_割当計算_({
    orders:[{ban:'101',code:'YMNGD09',sku:'YMNGD09b',qty:1,sortKey:101,i:0,keys:['YMNGD09']}],
    ledger:[], movements:[],
    supplies:[{ems:'EG049827401KR',code:'YMNGD09',qty:3,arrival:'2026-07-12'}], explicit:[]
  });
  assert.strictEqual(result.newRows.reduce((s,r)=>s+r.取り置き数量,0),1);
  assert.strictEqual(result.surplus[0].qty,2);
});

test('JMEE167の10個を既存7個と新規3個に分けても超過しない', () => {
  const result = context.取り置き_割当計算_({
    orders:[
      {ban:'10117284',code:'JMEE167',sku:'JMEE167b',qty:7,sortKey:10117284,i:0,keys:['JMEE167']},
      {ban:'10117602',code:'JMEE167',sku:'JMEE167b',qty:3,sortKey:10117602,i:1,keys:['JMEE167']}
    ],
    ledger:[{取置ID:'EMS|EG1|10117284',状態:'取り置き中',受注番号:'10117284',商品コード:'JMEE167',SKU:'JMEE167b',取り置き数量:7,取置元種別:'EMS',元EMS番号:'EG1'}],
    movements:[], supplies:[{ems:'EG1',code:'JMEE167',qty:10,arrival:'2026-07-12'}], explicit:[]
  });
  assert.strictEqual(result.newRows.length,1);
  assert.strictEqual(result.newRows[0].受注番号,'10117602');
  assert.strictEqual(result.newRows[0].取り置き数量,3);
  assert.strictEqual(JSON.stringify(result.errors),'[]');
});

test('注文番号在庫10117375の2個を指定注文だけへ割り当てる', () => {
  const result = context.取り置き_割当計算_({
    orders:[{ban:'10117375',code:'KAGURA08W-PG',sku:'KAGURA08W-PGb',qty:2,sortKey:10117375,i:0,keys:['KAGURA08W-PG']}],
    ledger:[], movements:[],
    supplies:[{ems:'EG1',code:'10117375',qty:2,arrival:'2026-07-12',directBan:'10117375'}], explicit:[]
  });
  assert.strictEqual(result.newRows.length,1);
  assert.strictEqual(result.newRows[0].受注番号,'10117375');
  assert.strictEqual(result.newRows[0].取り置き数量,2);
  assert.strictEqual(result.newRows[0].商品コード,'KAGURA08W-PG');
  assert.strictEqual(result.newRows[0].元EMS商品コード,'10117375');
});

test('同じ入力で再計算しても既存EMS確保を追加しない', () => {
  const input={orders:[{ban:'101',code:'AAA',sku:'AAAb',qty:1,sortKey:101,i:0,keys:['AAA']}],movements:[],supplies:[{ems:'EG1',code:'AAA',qty:1,arrival:'2026-07-12'}],explicit:[]};
  const first=context.取り置き_割当計算_(Object.assign({ledger:[]},input));
  const second=context.取り置き_割当計算_(Object.assign({ledger:first.newRows},input));
  assert.strictEqual(first.newRows.length,1);
  assert.strictEqual(second.newRows.length,0);
  assert.strictEqual(second.surplus.length,0);
});
```

- [ ] **Step 2: テストが未定義関数で失敗することを確認する**

Run: `node tests\project24_reservation_domain.test.js`

Expected: `取り置き_割当計算_ is not a function` でFAILする。

- [ ] **Step 3: 決定的IDと割当計算を実装する**

Add these helpers and keep the allocation order exactly `既存取り置き→現物あり戻し→注文番号指定→P列確定→EMS FIFO`:

```javascript
function 取り置き_決定ID_(source, ems, sourceCode, key, originId){
  return [source,String(ems||''),normCode_(sourceCode),String(originId||''),key].join('|');
}

function 取り置き_新規行_(order, qty, source, ems, originId, sourceCode){
  const key=取り置き_行キー_(order);
  return {
    取置ID:取り置き_決定ID_(source,ems,sourceCode||order.code,key,originId), 状態:TORIOKI_STATUS.ACTIVE,
    受注番号:String(order.ban), 商品コード:取り置き_商品コード_(order.sku,order.code), SKU:String(order.sku||''),
    取り置き数量:qty, 取置元種別:source, 元EMS番号:String(ems||''), 元EMS商品コード:normCode_(sourceCode||order.code), 元取置ID:String(originId||''),
    戻し処理結果:'', 終了理由・メモ:''
  };
}

function 取り置き_使用合計_(usage){
  return ['取り置き中','発送済み','戻し未処理','在庫なし確定','Yahoo移動済み']
    .reduce((sum,key)=>sum+(Number(usage&&usage[key])||0),0);
}

function 取り置き_割当計算_(input){
  const summary=取り置き_集計_(input.ledger||[],input.movements||[]);
  const orders=(input.orders||[]).map(o=>Object.assign({},o,{need:取り置き_今回必要数_(o,summary)}))
    .sort((a,b)=>(a.sortKey||0)-(b.sortKey||0)||(a.i||0)-(b.i||0));
  const newRows=[], returnUpdates=[], errors=summary.errors.slice();
  const matches=(order,code)=> (order.keys||[]).indexOf(normCode_(code))>=0 || 取り置き_商品コード_(order.sku,order.code)===normCode_(code);
  summary.confirmedReturns.slice().sort((a,b)=>String(a.登録日時||'').localeCompare(String(b.登録日時||''))).forEach(source=>{
    const originalQty=取り置き_整数_(source.取り置き数量); let left=originalQty;
    for(const order of orders){
      if(!left || !order.need || !matches(order,source.商品コード)) continue;
      const take=Math.min(left,order.need); order.need-=take; left-=take;
      newRows.push(取り置き_新規行_(order,take,'キャンセル再引当',source.元EMS番号,source.取置ID,source.元EMS商品コード||source.商品コード));
    }
    if(left<originalQty) returnUpdates.push(left===0
      ? {取置ID:source.取置ID,戻し処理結果:TORIOKI_RETURN.REALLOCATED}
      : {取置ID:source.取置ID,戻し処理結果:TORIOKI_RETURN.PRESENT,取り置き数量:left,終了理由・メモ:(originalQty-left)+'個を再引当済み'});
  });
  const supplyByKey={};
  (input.supplies||[]).forEach(s=>{
    const key=取り置き_供給キー_(s.ems,s.code);
    if(!supplyByKey[key]) supplyByKey[key]=Object.assign({},s,{qty:0});
    supplyByKey[key].qty+=(Number(s.qty)||0);
    if(s.directBan) supplyByKey[key].directBan=String(s.directBan);
  });
  const supplies=Object.keys(supplyByKey).map(key=>supplyByKey[key]);
  const remainingBySupply={};
  supplies.forEach(s=>{
    const key=取り置き_供給キー_(s.ems,s.code), used=取り置き_使用合計_(summary.usageBySupply[key]);
    remainingBySupply[key]=Math.max(0,(Number(s.qty)||0)-used);
  });
  const takeSupply=(s,order,qty,source)=>{
    const key=取り置き_供給キー_(s.ems,s.code), take=Math.min(qty,order.need,remainingBySupply[key]||0);
    if(take<=0) return 0;
    remainingBySupply[key]-=take; order.need-=take;
    newRows.push(取り置き_新規行_(order,take,source,s.ems,'',s.code));
    return take;
  };
  supplies.filter(s=>s.directBan).forEach(s=>{
    const order=orders.find(o=>String(o.ban)===String(s.directBan));
    if(!order) errors.push('注文番号指定の対象なし: '+s.directBan);
    else if(takeSupply(s,order,Number(s.qty)||0,'EMS')!==(Number(s.qty)||0)) errors.push('注文番号指定が注文数または供給数を超過: '+s.directBan);
  });
  (input.explicit||[]).forEach(e=>{
    const order=orders.find(o=>String(o.ban)===String(e.ban) && matches(o,e.code));
    const supply=supplies.find(s=>String(s.ems)===String(e.ems) && normCode_(s.code)===normCode_(e.code));
    if(!order || !supply) errors.push('P列確定を特定できない: '+e.ban+' '+e.code);
    else takeSupply(supply,order,Number(e.qty)||0,'EMS');
  });
  supplies.filter(s=>!s.directBan).forEach(s=>{
    for(const order of orders){
      if(!(remainingBySupply[取り置き_供給キー_(s.ems,s.code)]>0) || !order.need) break;
      if(matches(order,s.code)) takeSupply(s,order,order.need,'EMS');
    }
  });
  const surplus=supplies.map(s=>({ems:s.ems,code:normCode_(s.code),qty:remainingBySupply[取り置き_供給キー_(s.ems,s.code)]||0,arrival:s.arrival}))
    .filter(s=>s.qty>0);
  const plan={orders,newRows,returnUpdates,remainingBySupply,surplus,errors};
  plan.errors=plan.errors.concat(取り置き_割当検証_(plan,input,summary));
  return plan;
}
```

Implement the validation as a projected-ledger check so partial cancellation reallocations are counted exactly once:

```javascript
function 取り置き_割当検証_(plan,input){
  const projected=(input.ledger||[]).map(r=>Object.assign({},r));
  const byId={}; projected.forEach((r,index)=>byId[String(r.取置ID||'')]=index);
  (plan.returnUpdates||[]).forEach(update=>{
    const index=byId[String(update.取置ID||'')];
    if(index!==undefined) projected[index]=Object.assign({},projected[index],update);
  });
  (plan.newRows||[]).forEach(row=>{
    const id=String(row.取置ID||''), index=byId[id];
    if(index===undefined){ byId[id]=projected.length; projected.push(Object.assign({},row)); }
    else projected[index]=Object.assign({},projected[index],row);
  });
  const projectedSummary=取り置き_集計_(projected,input.movements||[]), errors=projectedSummary.errors.slice();
  const orderQty={};
  (input.orders||[]).forEach(order=>{
    const key=取り置き_行キー_(order); orderQty[key]=(orderQty[key]||0)+(Number(order.qty)||0);
  });
  Object.keys(projectedSummary.activeByKey).forEach(key=>{
    if(!(key in orderQty)) errors.push('取り置き中の対象注文なし: '+key);
    else if(projectedSummary.activeByKey[key]>orderQty[key]) errors.push('注文数量超過: '+key+' 取り置き'+projectedSummary.activeByKey[key]+' / 注文'+orderQty[key]);
  });
  const supplyQty={};
  (input.supplies||[]).forEach(s=>{
    const key=取り置き_供給キー_(s.ems,s.code); supplyQty[key]=(supplyQty[key]||0)+(Number(s.qty)||0);
  });
  Object.keys(projectedSummary.usageBySupply).forEach(key=>{
    if(!(key in supplyQty)) return; // 締め済み過去EMSは全件検算で確認し、現在④の供給上限には使わない
    const used=取り置き_使用合計_(projectedSummary.usageBySupply[key]), supplied=supplyQty[key]||0;
    if(used>supplied) errors.push('EMS供給超過: '+key+' 使用'+used+' / 供給'+supplied);
  });
  (plan.newRows||[]).forEach(row=>{
    if(!取り置き_整数_(row.取り置き数量)) errors.push('新規取り置き数量が正の整数でない: '+row.取置ID);
    const order=(input.orders||[]).find(o=>取り置き_行キー_(o)===取り置き_行キー_(row));
    if(order && 取り置き_商品コード_(order.sku,order.code)!==normCode_(row.商品コード)) errors.push('数値枝番または商品コード不一致: '+row.取置ID);
  });
  return Array.from(new Set(errors));
}
```

- [ ] **Step 4: ドメインテストを通す**

Run:

```powershell
node tests\project24_reservation_domain.test.js
node --check Project_24\取り置き計算.js
```

Expected: 全テストPASS、構文エラーなし。

- [ ] **Step 5: Task 2をコミットする**

```powershell
git add Project_24\取り置き計算.js tests\project24_reservation_domain.test.js
git commit -m "feat(hikiate): 取り置き優先の割当計算を追加"
```

---

---

### Task 3: 台帳シートI/Oと初回取り置き登録を実装する

**Files:**
- Create: `Project_24/取り置き台帳.js`
- Create: `tests/project24_reservation_workflow.test.js`
- Modify: `Project_24/引当.js` function `onOpen`

**Interfaces:**
- Consumes: Task 1のキー・集計関数、既存 `列マップ_(sheet)`, `区分_(value)`, `HIKIATE_CFG`。
- Produces: `取り置き台帳_読む_()`, `取り置き台帳_保存_(rows)`, `EMS在庫移動台帳_読む_()`, `取り置き初期登録を作成()`, `取り置き初期登録を確定()`。

- [ ] **Step 1: 初期候補・入力検証・冪等IDの失敗テストを作成する**

Create the workflow test with the same VM harness and add:

```javascript
test('部分在庫と希望日待ちの受注だけを初期候補にする', () => {
  const orders=[
    {ban:'101',code:'AAA',sku:'AAAb',qty:2,kbn:'取り寄せ'},
    {ban:'102',code:'BBB',sku:'BBBb',qty:1,kbn:'取り寄せ'},
    {ban:'103',code:'CCC',sku:'CCCb',qty:1,kbn:'取り寄せ'}
  ];
  const rows=context.取り置き_初期候補_(orders,new Set(['101']),new Set(['102']));
  assert.strictEqual(rows.length,2);
  assert.strictEqual(rows[0].取置ID,'INIT|101|AAA|AAAB');
  assert.strictEqual(rows[1].取置ID,'INIT|102|BBB|BBBB');
});

test('初回入力は空欄と0を除外し注文数量超過を止める', () => {
  const result=context.取り置き_初期確定計画_([
    {取置ID:'INIT|1',受注番号:'101',商品コード:'AAA',SKU:'AAAb',注文数量:2,現物取り置き数量:2},
    {取置ID:'INIT|2',受注番号:'102',商品コード:'BBB',SKU:'BBBb',注文数量:1,現物取り置き数量:0},
    {取置ID:'INIT|3',受注番号:'103',商品コード:'CCC',SKU:'CCCb',注文数量:1,現物取り置き数量:2}
  ],[],'2026-07-15 10:00:00');
  assert.ok(result.errors.some(e=>/101|102/.test(e))===false);
  assert.ok(result.errors.some(e=>/103/.test(e)));
  assert.strictEqual(result.rows.length,0, '1件でもエラーなら保存対象を返さない');
});

test('同じINITキーは追加せず目標数量へ洗い替える', () => {
  const existing=[{取置ID:'INIT|101|AAA|AAAB',状態:'取り置き中',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:1,取置元種別:'開始前在庫'}];
  const result=context.取り置き_初期確定計画_([
    {取置ID:'INIT|101|AAA|AAAB',受注番号:'101',商品コード:'AAA',SKU:'AAAb',注文数量:2,現物取り置き数量:2}
  ],existing,'2026-07-15 10:00:00');
  assert.strictEqual(JSON.stringify(result.errors),'[]');
  assert.strictEqual(result.rows.length,1);
  assert.strictEqual(result.rows[0].取り置き数量,2);
});
```

- [ ] **Step 2: テストが未定義関数で失敗することを確認する**

Run: `node tests\project24_reservation_workflow.test.js`

Expected: `取り置き_初期候補_ is not a function` でFAILする。

- [ ] **Step 3: シート設定と純粋な初期登録計画を実装する**

Create `Project_24/取り置き台帳.js` with this exact schema:

```javascript
const TORIOKI_CFG = Object.freeze({
  台帳:'取り置き台帳', 初期:'取り置き初期登録', 戻し:'キャンセル戻し確認', Yahoo候補:'Yahoo戻し候補', 移動:'EMS在庫移動台帳',
  台帳HDR:['取置ID','状態','受注番号','商品コード','SKU','取り置き数量','取置元種別','元EMS番号','元EMS商品コード','元取置ID','登録日時','更新日時','戻し処理結果','終了理由・メモ'],
  初期HDR:['取置ID','受注番号','商品コード','SKU','注文数量','現物取り置き数量','メモ','判定'],
  移動HDR:['処理ID','EMS番号','商品コード','数量','移動先','処理日時']
});

function 取り置き_初期候補_(orders, partialBans, holdBans){
  return (orders||[]).filter(o=>partialBans.has(String(o.ban))||holdBans.has(String(o.ban))).map(o=>{
    const key=取り置き_行キー_(o);
    return {取置ID:'INIT|'+key,受注番号:String(o.ban),商品コード:取り置き_商品コード_(o.sku,o.code),SKU:String(o.sku||''),注文数量:Number(o.qty)||0,現物取り置き数量:'',メモ:'',判定:''};
  });
}

function 取り置き_初期確定計画_(inputRows, existingRows, now){
  const errors=[], targets={}, inputIds=new Set();
  (inputRows||[]).forEach((r,index)=>{
    inputIds.add(String(r.取置ID||''));
    const entered=Number(r.現物取り置き数量)||0, ordered=Number(r.注文数量)||0;
    if(entered<0 || !Number.isInteger(entered)) errors.push('初期登録'+(index+2)+'行: 数量は0以上の整数');
    if(entered>ordered) errors.push('受注'+r.受注番号+': 現物'+entered+'が注文'+ordered+'を超過');
    if(entered>0) targets[r.取置ID]=Object.assign({},r,{取り置き数量:entered});
  });
  if(errors.length) return {rows:[],errors};
  const kept=(existingRows||[]).filter(r=>r.取置元種別!=='開始前在庫' || !inputIds.has(String(r.取置ID||'')));
  Object.keys(targets).forEach(id=>{
    const r=targets[id];
    kept.push({取置ID:id,状態:TORIOKI_STATUS.ACTIVE,受注番号:r.受注番号,商品コード:r.商品コード,SKU:r.SKU,
      取り置き数量:r.取り置き数量,取置元種別:'開始前在庫',元EMS番号:'',元EMS商品コード:'',元取置ID:'',登録日時:now,更新日時:now,
      戻し処理結果:'',終了理由・メモ:String(r.メモ||'')});
  });
  return {rows:kept,errors:[]};
}
```

Implement one generic header-based adapter and reuse it for both ledgers:

```javascript
function 取り置き_表を読む_(sheetName, headers){
  const sh=SpreadsheetApp.getActive().getSheetByName(sheetName);
  if(!sh || sh.getLastRow()<2) return [];
  const head=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(v=>String(v||'').trim());
  const index={}; headers.forEach(h=>index[h]=head.indexOf(h));
  if(headers.some(h=>index[h]<0)) throw new Error(sheetName+'の見出し不足: '+headers.filter(h=>index[h]<0).join(','));
  return sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues().map(row=>{
    const obj={}; headers.forEach(h=>obj[h]=row[index[h]]); return obj;
  }).filter(obj=>String(obj[headers[0]]||'').trim());
}

function 取り置き_表を保存_(sheetName, headers, rows){
  const ss=SpreadsheetApp.getActive(); let sh=ss.getSheetByName(sheetName); if(!sh) sh=ss.insertSheet(sheetName);
  if(sh.getMaxColumns()<headers.length) sh.insertColumnsAfter(sh.getMaxColumns(),headers.length-sh.getMaxColumns());
  if(sh.getMaxRows()<rows.length+1) sh.insertRowsAfter(sh.getMaxRows(),rows.length+1-sh.getMaxRows());
  sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
  if(sh.getMaxRows()>1) sh.getRange(2,1,sh.getMaxRows()-1,sh.getMaxColumns()).clearContent();
  if(rows.length) sh.getRange(2,1,rows.length,headers.length).setValues(rows.map(r=>headers.map(h=>r[h]==null?'':r[h])));
  sh.setFrozenRows(1);
}

function 取り置き台帳_読む_(){ return 取り置き_表を読む_(TORIOKI_CFG.台帳,TORIOKI_CFG.台帳HDR); }
function 取り置き台帳_保存_(rows){ 取り置き_表を保存_(TORIOKI_CFG.台帳,TORIOKI_CFG.台帳HDR,rows); }
function EMS在庫移動台帳_読む_(){ return 取り置き_表を読む_(TORIOKI_CFG.移動,TORIOKI_CFG.移動HDR); }
function EMS在庫移動台帳_保存_(rows){ 取り置き_表を保存_(TORIOKI_CFG.移動,TORIOKI_CFG.移動HDR,rows); }
```

Add the public initial-registration functions:

```javascript
function 取り置き_受注番号集合_(sheetName){
  const sh=SpreadsheetApp.getActive().getSheetByName(sheetName), out=new Set();
  if(!sh || sh.getLastRow()<2) return out;
  const head=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(v=>String(v||'').trim());
  const col=head.indexOf('受注番号'); if(col<0) throw new Error(sheetName+'に受注番号見出しがありません');
  sh.getRange(2,col+1,sh.getLastRow()-1,1).getDisplayValues().forEach(r=>{ const ban=String(r[0]||'').trim(); if(ban) out.add(ban); });
  return out;
}

function 取り置き初期登録を作成(){
  const ss=SpreadsheetApp.getActive(), recv=ss.getSheetByName(HIKIATE_CFG.受注), ui=SpreadsheetApp.getUi();
  if(!recv){ ui.alert('受注明細がありません'); return; }
  const M=列マップ_(recv), values=recv.getDataRange().getValues(), orders=[];
  for(let i=M.hr;i<values.length;i++){
    const row=values[i], ban=String(row[M.番号]||'').trim(), qty=Number(row[M.個数])||0;
    if(!ban || qty<=0 || 区分_(row[M.選択肢])!=='取り寄せ') continue;
    orders.push({ban,code:String(row[M.コード]||''),sku:M.SKU>=0?String(row[M.SKU]||''):'',qty});
  }
  const candidates=取り置き_初期候補_(orders,取り置き_受注番号集合_(HIKIATE_CFG.部分),取り置き_受注番号集合_(HIKIATE_CFG.希望));
  取り置き_表を保存_(TORIOKI_CFG.初期,TORIOKI_CFG.初期HDR,candidates);
  ui.alert('取り置き初期登録を作成しました','候補'+candidates.length+'行です。棚の現物を確認し「現物取り置き数量」だけ入力してください。',ui.ButtonSet.OK);
}

function 取り置き初期登録を確定(){
  const ui=SpreadsheetApp.getUi(), inputs=取り置き_表を読む_(TORIOKI_CFG.初期,TORIOKI_CFG.初期HDR);
  const plan=取り置き_初期確定計画_(inputs,取り置き台帳_読む_(),new Date());
  if(plan.errors.length){ ui.alert('初期登録を中止しました',plan.errors.join('\n'),ui.ButtonSet.OK); return; }
  const selected=inputs.filter(r=>(Number(r.現物取り置き数量)||0)>0), qty=selected.reduce((s,r)=>s+(Number(r.現物取り置き数量)||0),0);
  const answer=ui.alert('初期取り置きを確定します','対象'+selected.length+'行 / 合計'+qty+'個です。確定しますか？',ui.ButtonSet.OK_CANCEL);
  if(answer!==ui.Button.OK) return;
  取り置き台帳_保存_(plan.rows);
  SpreadsheetApp.getActive().toast('初期取り置き '+selected.length+'行 / '+qty+'個を確定しました','取り置き台帳',7);
}
```

- [ ] **Step 4: メニューへ初回登録入口を追加する**

In `onOpen()` under `📥 受注・共通`, add:

```javascript
.addItem('📋 取り置き初期登録を作成', '取り置き初期登録を作成')
.addItem('✅ 取り置き初期登録を確定', '取り置き初期登録を確定')
```

- [ ] **Step 5: テスト・構文・既存回帰を通す**

Run:

```powershell
node tests\project24_reservation_workflow.test.js
node tests\project24_reservation_domain.test.js
node tests\project24_torioki.test.js
node --check Project_24\取り置き台帳.js
node --check Project_24\引当.js
```

Expected: 全テストPASS、構文エラーなし。

- [ ] **Step 6: Task 3をコミットする**

```powershell
git add Project_24\取り置き台帳.js Project_24\引当.js tests\project24_reservation_workflow.test.js
git commit -m "feat(hikiate): 取り置き台帳と初期登録を追加"
```

---

---

### Task 4: ①CSV更新で発送・キャンセルを台帳へ遷移させる

**Files:**
- Modify: `Project_24/取り置き台帳.js`
- Modify: `Project_24/引当.js` functions `取込_実行_`, `列マップ_`
- Modify: `Project_24/消込台帳.js` function `キャンセル処理_`
- Modify: `tests/project24_reservation_workflow.test.js`

**Interfaces:**
- Consumes: 受注明細CSVの全行と現在台帳。
- Produces: `取り置き_CSV遷移計画_(csvRows, ledgerRows)`, `取り置き_CSV遷移を反映_(plan)`。
- Missing order behavior: 状態を変えず `要確認` に出す。消滅だけで発送済みにしない。

- [ ] **Step 1: 明示ステータス遷移の失敗テストを追加する**

```javascript
test('処理済だけ発送済み、キャンセルだけキャンセル戻しへ遷移する', () => {
  const ledger=[
    {取置ID:'A',状態:'取り置き中',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:1,戻し処理結果:''},
    {取置ID:'B',状態:'取り置き中',受注番号:'102',商品コード:'BBB',SKU:'BBBb',取り置き数量:1,戻し処理結果:''},
    {取置ID:'C',状態:'取り置き中',受注番号:'103',商品コード:'CCC',SKU:'CCCb',取り置き数量:1,戻し処理結果:''}
  ];
  const plan=context.取り置き_CSV遷移計画_([
    {受注番号:'101',商品コード:'AAA',SKU:'AAAb',受注ステータス:'処理済'},
    {受注番号:'102',商品コード:'BBB',SKU:'BBBb',受注ステータス:'キャンセル'}
  ],ledger,'2026-07-15 10:00:00');
  assert.strictEqual(plan.rows.find(r=>r.取置ID==='A').状態,'発送済み');
  assert.strictEqual(plan.rows.find(r=>r.取置ID==='B').状態,'キャンセル戻し');
  assert.strictEqual(plan.rows.find(r=>r.取置ID==='B').戻し処理結果,'未確認');
  assert.strictEqual(plan.rows.find(r=>r.取置ID==='C').状態,'取り置き中');
  assert.ok(plan.review.some(x=>x.取置ID==='C'));
});

test('列不足または重複する矛盾ステータスでは全遷移を止める', () => {
  const ledger=[{取置ID:'A',状態:'取り置き中',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:1}];
  const plan=context.取り置き_CSV遷移計画_([
    {受注番号:'101',商品コード:'AAA',SKU:'AAAb',受注ステータス:'処理済'},
    {受注番号:'101',商品コード:'AAA',SKU:'AAAb',受注ステータス:'キャンセル'}
  ],ledger,'2026-07-15 10:00:00');
  assert.ok(plan.errors.some(e=>/競合/.test(e)));
  assert.strictEqual(plan.rows.length,0);
});
```

- [ ] **Step 2: 純粋なCSV遷移計画を実装する**

```javascript
function 取り置き_CSV遷移計画_(csvRows, ledgerRows, now){
  const statuses={}, errors=[], review=[];
  (csvRows||[]).forEach(r=>{
    const ban=String(r.受注番号||'').replace(/^niyantarose-/i,'').trim();
    const key=取り置き_行キー_({ban,code:r.商品コード,sku:r.SKU});
    if(!ban || !取り置き_商品コード_(r.SKU,r.商品コード)) return;
    const next=/キャンセル/.test(String(r.受注ステータス||''))?'キャンセル':/処理済|発送済|出荷済/.test(String(r.受注ステータス||''))?'発送済み':'継続';
    if(statuses[key] && statuses[key]!==next) errors.push('同じ受注行にステータス競合: '+key);
    statuses[key]=next;
  });
  if(errors.length) return {rows:[],review,errors};
  const rows=(ledgerRows||[]).map(r=>{
    const copy=Object.assign({},r); if(copy.状態!==TORIOKI_STATUS.ACTIVE) return copy;
    const key=取り置き_行キー_(copy), next=statuses[key];
    if(next==='発送済み'){ copy.状態=TORIOKI_STATUS.SHIPPED; copy.更新日時=now; }
    else if(next==='キャンセル'){ copy.状態=TORIOKI_STATUS.RETURN; copy.戻し処理結果=TORIOKI_RETURN.UNCHECKED; copy.更新日時=now; }
    else if(!next) review.push({取置ID:copy.取置ID,受注番号:copy.受注番号,商品コード:copy.商品コード,理由:'最新CSVに受注行なし'});
    return copy;
  });
  return {rows,review,errors:[]};
}
```

- [ ] **Step 3: `取込_実行_` を計画→検証→反映の順に接続する**

Immediately after CSV normalization and before clearing `受注明細`, build objects from the original full-status CSV:

```javascript
const 取り置き遷移=取り置き_CSV遷移計画_(CSV行を受注行オブジェクトへ_(data[0],data.slice(1)),取り置き台帳_読む_(),new Date());
if(取り置き遷移.errors.length){
  ui.alert('取り置き台帳の更新を中止しました',取り置き遷移.errors.join('\n'),ui.ButtonSet.OK);
  return;
}
```

Add this adapter to `取り置き台帳.js`:

```javascript
function CSV行を受注行オブジェクトへ_(header, rows){
  const head=(header||[]).map(v=>String(v||'').trim()), index=name=>head.indexOf(name);
  const cBan=index('受注番号'), cStatus=index('受注ステータス'), cCode=index('商品コード'), cQty=index('個数');
  const cSku=index('商品SKU')>=0?index('商品SKU'):index('SKU');
  const missing=[];
  if(cBan<0) missing.push('受注番号'); if(cStatus<0) missing.push('受注ステータス');
  if(cCode<0) missing.push('商品コード'); if(cQty<0) missing.push('個数'); if(cSku<0) missing.push('商品SKU/SKU');
  if(missing.length) throw new Error('全ステータスCSVの見出し不足: '+missing.join(','));
  return (rows||[]).map(row=>({受注番号:String(row[cBan]||'').replace(/^niyantarose-/i,''),受注ステータス:String(row[cStatus]||''),
    商品コード:String(row[cCode]||''),SKU:String(row[cSku]||''),個数:Number(row[cQty])||0}));
}
```

After the existing受注明細 write succeeds, call `取り置き台帳_保存_(取り置き遷移.rows)` once and render `review` to a `取り置き要確認` sheet without changing those ledger rows.

Remove `退避取置` and the old `取り置き数` restore block from `取込_実行_`; keep the physical column only as legacy display if it already exists, but never read it into calculation.

- [ ] **Step 4: 手動キャンセル入口も同じ遷移へ統一する**

At the end of `キャンセル処理_(bans)`, call:

```javascript
const ledger=取り置き台帳_読む_();
const now=new Date();
const rows=ledger.map(r=>{
  if(r.状態!=='取り置き中' || bans.indexOf(String(r.受注番号))<0) return r;
  return Object.assign({},r,{状態:'キャンセル戻し',戻し処理結果:'未確認',更新日時:now});
});
取り置き台帳_保存_(rows);
```

This does not add Yahoo inventory and does not create a replacement reservation.

- [ ] **Step 5: テストと構文を通す**

Run:

```powershell
node tests\project24_reservation_workflow.test.js
node tests\project24_reservation_domain.test.js
node tests\project24_torioki.test.js
node --check Project_24\取り置き台帳.js
node --check Project_24\引当.js
node --check Project_24\消込台帳.js
```

Expected: 全テストPASS、構文エラーなし。

- [ ] **Step 6: Task 4をコミットする**

```powershell
git add Project_24\取り置き台帳.js Project_24\引当.js Project_24\消込台帳.js tests\project24_reservation_workflow.test.js
git commit -m "feat(hikiate): CSV状態を取り置き台帳へ反映"
```

---

---

### Task 5: キャンセル戻しの現物確認と優先再引当を実装する

**Files:**
- Modify: `Project_24/取り置き台帳.js`
- Modify: `Project_24/引当.js` function `onOpen`
- Modify: `tests/project24_reservation_workflow.test.js`
- Modify: `tests/project24_reservation_domain.test.js`

**Interfaces:**
- Consumes: 状態 `キャンセル戻し` の台帳行。
- Produces: `キャンセル戻し確認を更新()`, `キャンセル戻し確認を確定()`, `Yahoo戻し候補を更新_()`, `選択した取り置きを手動解除()`。
- `現物あり` は次回④で最古注文へ優先再引当、`在庫なし` は供給へ戻さない。

- [ ] **Step 1: 現物確認入力の失敗テストを追加する**

```javascript
test('現物ありと在庫なしだけを戻し結果へ反映する', () => {
  const ledger=[
    {取置ID:'A',状態:'キャンセル戻し',受注番号:'101',商品コード:'AAA',取り置き数量:1,戻し処理結果:'未確認'},
    {取置ID:'B',状態:'キャンセル戻し',受注番号:'102',商品コード:'BBB',取り置き数量:1,戻し処理結果:'未確認'}
  ];
  const result=context.取り置き_戻し確認計画_([{取置ID:'A',現物確認:'現物あり'},{取置ID:'B',現物確認:'在庫なし'}],ledger,'2026-07-15 10:00:00');
  assert.strictEqual(JSON.stringify(result.errors),'[]');
  assert.strictEqual(result.rows.find(r=>r.取置ID==='A').戻し処理結果,'現物あり');
  assert.strictEqual(result.rows.find(r=>r.取置ID==='B').戻し処理結果,'在庫なし');
});

test('未確認のまま、または不明な選択肢では確定しない', () => {
  const ledger=[{取置ID:'A',状態:'キャンセル戻し',受注番号:'101',商品コード:'AAA',取り置き数量:1,戻し処理結果:'未確認'}];
  const result=context.取り置き_戻し確認計画_([{取置ID:'A',現物確認:'たぶんある'}],ledger,'2026-07-15 10:00:00');
  assert.ok(result.errors.length>0);
  assert.strictEqual(result.rows.length,0);
});
```

- [ ] **Step 2: 戻し確認計画と2つの運用シートを実装する**

Use this schema:

```javascript
const 戻しHDR=['取置ID','受注番号','商品コード','数量','元EMS番号','現物確認','メモ'];
const Yahoo候補HDR=['取置ID','商品コード','数量','元EMS番号','処理ID','確認'];

function 取り置き_戻し確認計画_(inputs, ledger, now){
  const byId={}; (inputs||[]).forEach(r=>byId[String(r.取置ID||'')]=String(r.現物確認||'').trim());
  const errors=[];
  Object.keys(byId).forEach(id=>{ if(byId[id] && byId[id]!=='現物あり' && byId[id]!=='在庫なし') errors.push(id+': 現物確認は「現物あり」か「在庫なし」'); });
  if(errors.length) return {rows:[],errors};
  return {rows:(ledger||[]).map(r=>{
    const choice=byId[String(r.取置ID||'')]; if(!choice) return Object.assign({},r);
    if(r.状態!=='キャンセル戻し' || r.戻し処理結果!=='未確認'){ errors.push(r.取置ID+': 未確認のキャンセル戻しではない'); return Object.assign({},r); }
    return Object.assign({},r,{戻し処理結果:choice,更新日時:now});
  }),errors};
}
```

Add the three sheet functions:

```javascript
function キャンセル戻し確認を更新(){
  const rows=取り置き台帳_読む_().filter(r=>r.状態==='キャンセル戻し'&&r.戻し処理結果==='未確認').map(r=>({
    取置ID:r.取置ID,受注番号:r.受注番号,商品コード:r.商品コード,数量:r.取り置き数量,元EMS番号:r.元EMS番号,現物確認:'',メモ:r['終了理由・メモ']||''
  }));
  取り置き_表を保存_(TORIOKI_CFG.戻し,戻しHDR,rows);
  const sh=SpreadsheetApp.getActive().getSheetByName(TORIOKI_CFG.戻し);
  if(rows.length) sh.getRange(2,6,rows.length,1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['現物あり','在庫なし'],true).setAllowInvalid(false).build());
  SpreadsheetApp.getActive().toast('未確認のキャンセル戻し '+rows.length+'件','取り置き台帳',6);
}

function キャンセル戻し確認を確定(){
  const ui=SpreadsheetApp.getUi(), inputs=取り置き_表を読む_(TORIOKI_CFG.戻し,戻しHDR);
  const plan=取り置き_戻し確認計画_(inputs,取り置き台帳_読む_(),new Date());
  if(plan.errors.length){ ui.alert('キャンセル戻しを確定できません',plan.errors.join('\n'),ui.ButtonSet.OK); return; }
  const answer=ui.alert('現物確認を確定します','入力済み'+inputs.filter(r=>r.現物確認).length+'件を台帳へ反映します。',ui.ButtonSet.OK_CANCEL);
  if(answer!==ui.Button.OK) return;
  取り置き台帳_保存_(plan.rows); キャンセル戻し確認を更新(); Yahoo戻し候補を更新_();
}

function Yahoo戻し候補を更新_(){
  const rows=取り置き台帳_読む_().filter(r=>r.状態==='キャンセル戻し'&&r.戻し処理結果==='現物あり').map(r=>({
    取置ID:r.取置ID,商品コード:r.商品コード,数量:r.取り置き数量,元EMS番号:r.元EMS番号,処理ID:'YAHOO|RETURN|'+r.取置ID,確認:''
  }));
  取り置き_表を保存_(TORIOKI_CFG.Yahoo候補,Yahoo候補HDR,rows);
  const sh=SpreadsheetApp.getActive().getSheetByName(TORIOKI_CFG.Yahoo候補);
  if(rows.length) sh.getRange(2,6,rows.length,1).insertCheckboxes();
}
```

Because Task 2 changes a partially reallocated original return row to its leftover quantity, `Yahoo戻し候補を更新_()` displays only the unassigned physical remainder after④.

- [ ] **Step 3: メニューへ確認入口を追加する**

```javascript
.addItem('📦 キャンセル戻し確認を更新', 'キャンセル戻し確認を更新')
.addItem('✅ キャンセル戻し確認を確定', 'キャンセル戻し確認を確定')
.addItem('🔓 選択した取り置きを手動解除', '選択した取り置きを手動解除')
```

Implement manual release only from a selected row on `取り置き台帳`; require a non-empty reason:

```javascript
function 選択した取り置きを手動解除(){
  const ss=SpreadsheetApp.getActive(), sh=ss.getActiveSheet(), ui=SpreadsheetApp.getUi(), row=sh.getActiveRange().getRow();
  if(sh.getName()!==TORIOKI_CFG.台帳 || row<2){ ui.alert('取り置き台帳の解除する行を選択してください'); return; }
  const ledger=取り置き台帳_読む_(), target=ledger[row-2];
  if(!target || target.状態!=='取り置き中'){ ui.alert('取り置き中の行だけ手動解除できます'); return; }
  const response=ui.prompt('手動解除の理由','登録間違い、現物不足などの理由を入力してください。',ui.ButtonSet.OK_CANCEL);
  if(response.getSelectedButton()!==ui.Button.OK) return;
  const reason=String(response.getResponseText()||'').trim(); if(!reason){ ui.alert('解除理由は必須です'); return; }
  ledger[row-2]=Object.assign({},target,{状態:'手動解除','終了理由・メモ':reason,更新日時:new Date()});
  取り置き台帳_保存_(ledger);
}
```

- [ ] **Step 4: テスト・構文・再引当回帰を通す**

Run:

```powershell
node tests\project24_reservation_workflow.test.js
node tests\project24_reservation_domain.test.js
node --check Project_24\取り置き台帳.js
node --check Project_24\取り置き計算.js
node --check Project_24\引当.js
```

Expected: `現物あり` がTask 2の最古注文へ一度だけ再引当され、`在庫なし` は `surplus` に出ず、全テストPASS。

- [ ] **Step 5: Task 5をコミットする**

```powershell
git add Project_24\取り置き台帳.js Project_24\引当.js tests\project24_reservation_workflow.test.js tests\project24_reservation_domain.test.js
git commit -m "feat(hikiate): キャンセル戻し確認を追加"
```

---

---

### Task 6: P列自動記入を台帳基準のプレビュー→反映へ分離する

**Files:**
- Modify: `Project_24/P列自動記入.js` functions `発注共有P列記入_`, `P列書き直し実行_`
- Modify: `Project_24/取り置き計算.js`
- Modify: `tests/project24_reservation_workflow.test.js`

**Interfaces:**
- Consumes: 現役注文、取り置き台帳集計、EMS在庫移動台帳、発注共有EMSリスト。
- Produces: `発注共有P列計画_(options)`, `発注共有P列計画を反映_(plan)`, `P列計画_確定割当_(plan)`。`options.currentP` はP列書き直しプレビュー用で、省略時は現在値を使う。
- Compatibility: public `発注共有P列記入_()` remains and becomes plan+apply wrapper.

- [ ] **Step 1: 台帳確保を保持し、発送推測を混ぜない失敗テストを追加する**

```javascript
test('P列計画は同じEMSの取り置き中を固定し、残数だけ新規FIFOへ回す', () => {
  const rows=[{ems:'EG1',code:'AAA',qty:3,pOriginal:'101',arrival:'2026-07-12'}];
  const orders=[
    {ban:'101',code:'AAA',sku:'AAAb',qty:1,need:0,date:new Date('2026-07-01'),keys:['AAA']},
    {ban:'102',code:'AAA',sku:'AAAb',qty:2,need:2,date:new Date('2026-07-02'),keys:['AAA']}
  ];
  const fixed={'EG1|AAA':[{ban:'101',qty:1}]};
  const usage={'EG1|AAA':{取り置き中:1,発送済み:0,戻し未処理:0,在庫なし確定:0,Yahoo移動済み:0}};
  const result=context.P列計画_純計算_(rows,orders,fixed,usage);
  assert.strictEqual(result.rows[0].nextP,'101, 102:2');
  assert.strictEqual(result.rows[0].left,0);
});

test('発送済みと戻し未処理は供給を塞ぐがP列の現役注文表示には出さない', () => {
  const rows=[{ems:'EG1',code:'AAA',qty:3,pOriginal:'',arrival:'2026-07-12'}];
  const orders=[{ban:'102',code:'AAA',sku:'AAAb',qty:2,need:2,date:new Date('2026-07-02'),keys:['AAA']}];
  const usage={'EG1|AAA':{取り置き中:0,発送済み:1,戻し未処理:1,在庫なし確定:0,Yahoo移動済み:0}};
  const result=context.P列計画_純計算_(rows,orders,{},usage);
  assert.strictEqual(result.rows[0].nextP,'102');
  assert.strictEqual(result.rows[0].entries[0].qty,1);
  assert.strictEqual(orders[0].need,2,'入力注文を変更しない');
});

test('P列計画を作っただけではsetValuesを呼ばない', () => {
  let writes=0;
  const plan={sheet:{getRange:()=>({setValues:()=>{writes++;}})},startRow:7,colP:16,values:[['101']],writes:[0]};
  assert.strictEqual(writes,0);
  context.発注共有P列計画を反映_(plan);
  assert.strictEqual(writes,1);
});
```

- [ ] **Step 2: 純計算とP列計画オブジェクトを実装する**

Add to `取り置き計算.js`:

```javascript
function P列計画_純計算_(emsRows, inputOrders, fixedBySupply, usageBySupply){
  const orders=(inputOrders||[]).map(o=>Object.assign({},o));
  const rows=(emsRows||[]).map(r=>Object.assign({},r,{entries:[],left:0,nextP:''}));
  const fixedRemaining={}, usageRemaining={};
  Object.keys(fixedBySupply||{}).forEach(key=>fixedRemaining[key]=(fixedBySupply[key]||[]).map(e=>Object.assign({},e)));
  Object.keys(usageBySupply||{}).forEach(key=>usageRemaining[key]=取り置き_使用合計_(usageBySupply[key]));
  const takeFixed=(key,limit)=>{
    const out=[], queue=fixedRemaining[key]||[]; let left=limit;
    while(left>0 && queue.length){
      const take=Math.min(left,Number(queue[0].qty)||0);
      if(take>0) out.push({ban:String(queue[0].ban),qty:take});
      queue[0].qty-=take; left-=take;
      if(queue[0].qty<=0) queue.shift();
    }
    return out;
  };
  rows.forEach(r=>{
    const key=取り置き_供給キー_(r.ems,r.code), qty=Math.max(0,Number(r.qty)||0);
    const fixed=takeFixed(key,qty);
    const fixedQty=fixed.reduce((s,e)=>s+(Number(e.qty)||0),0);
    usageRemaining[key]=Math.max(0,(usageRemaining[key]||0)-fixedQty);
    const nonDisplayUsed=Math.min(qty-fixedQty,usageRemaining[key]||0);
    usageRemaining[key]=Math.max(0,(usageRemaining[key]||0)-nonDisplayUsed);
    let capacity=Math.max(0,qty-fixedQty-nonDisplayUsed);
    r.entries=fixed;
    for(const order of orders.sort((a,b)=>a.date-b.date)){
      if(capacity<=0) break;
      if((order.keys||[]).indexOf(normCode_(r.code))<0 || !(order.need>0)) continue;
      const take=Math.min(capacity,order.need); capacity-=take; order.need-=take;
      const prev=r.entries.find(e=>String(e.ban)===String(order.ban));
      if(prev) prev.qty+=take; else r.entries.push({ban:String(order.ban),qty:take});
    }
    r.left=capacity;
    r.nextP=P列指定文字列_(r.entries,qty);
  });
  return {rows,orders};
}
```

In `P列自動記入.js`, extract all reads and calculations from `発注共有P列記入_()` into `発注共有P列計画_(options)`, beginning with `options=options||{}`. It must not call `setValues`, `setBackground`, or mutate any sheet. Use `options.currentP` when supplied, otherwise read the current P column. Build `fixedBySupply` from `取り置き_集計_().activeRowsByKey`: include only `状態=取り置き中`, `取置元種別=EMS|キャンセル再引当`, and non-empty `元EMS番号`; its supply key is `元EMS番号|元EMS商品コード`（旧行だけ `商品コード` fallback）。Build `usageBySupply` from the same summary. Return:

```javascript
{
  error:'', sheet:ems, startRow:hr+1, colP, rowCount:n,
  values:pColumn, backgrounds:[{a1,color}], writes:[rowIndex],
  rows:calculated.rows, 到着実績, 到着便, 到着実績取得済,
  summary:{記入,分割,在庫,既存,過剰除外:0,解析警告:0}
}
```

Add apply and compatibility functions:

```javascript
function 発注共有P列計画を反映_(plan){
  if(plan.error || !plan.writes.length) return plan.summary||plan;
  plan.sheet.getRange(plan.startRow,plan.colP,plan.rowCount,1).setValues(plan.values);
  plan.backgrounds.forEach(item=>plan.sheet.getRange(item.a1).setBackground(item.color));
  return Object.assign({},plan.summary,{到着実績:plan.到着実績,到着便:plan.到着便,到着実績取得済:plan.到着実績取得済});
}

function P列計画_確定割当_(plan){
  const out=[];
  (plan.rows||[]).forEach(r=>(r.entries||[]).forEach(e=>out.push({ems:r.ems,code:r.code,ban:e.ban,qty:e.qty})));
  return out;
}

function 発注共有P列記入_(){
  const plan=発注共有P列計画_();
  if(plan.error) return {error:plan.error};
  return 発注共有P列計画を反映_(plan);
}
```

Remove `消込台帳_出荷済み行_()` and `引当履歴_需要を差し引く_()` from P列 need construction. Their exact source and quantity now come from `usageBySupply`, not date inference.

- [ ] **Step 3: P列書き直しもプレビュー計画を経由させる**

`P列書き直し実行_()` must create a clear-plan in memory, call `発注共有P列計画_({currentP:clearedValues})`, validate the plan, then perform one P-column `setValues`. Do not clear P first and calculate second.

- [ ] **Step 4: テスト・構文・既存P回帰を通す**

Run:

```powershell
node tests\project24_reservation_workflow.test.js
node tests\project24_reservation_domain.test.js
node tests\project24_arrived_box_color.test.js
node --check Project_24\P列自動記入.js
node --check Project_24\取り置き計算.js
```

Expected: 全テストPASS。計画作成だけでは書込0回、apply時だけP列を1回書く。

- [ ] **Step 5: Task 6をコミットする**

```powershell
git add Project_24\P列自動記入.js Project_24\取り置き計算.js tests\project24_reservation_workflow.test.js
git commit -m "refactor(hikiate): P列更新を台帳基準の計画反映へ分離"
```

---

---

### Task 7: ④引当実行を取り置き台帳の計画へ接続する

**Files:**
- Modify: `Project_24/引当.js` functions `残必要計算_`, `引当行状態_`, `注文出荷準備OK_`, `引当実行_本体_`
- Modify: `Project_24/取り置き台帳.js`
- Modify: `tests/project24_torioki.test.js`
- Modify: `tests/project24_arrived_box_color.test.js`
- Modify: `tests/project24_reservation_domain.test.js`

**Interfaces:**
- Consumes: `取り置き台帳_読む_()`, `EMS在庫移動台帳_読む_()`, `発注共有P列計画_(options)`, `取り置き_割当計算_(input)`。
- Produces: `取り置き台帳_割当計画を反映_(plan, existingRows, now)`, `引当実行_本体_(options)`, `引当切替差分を作成()` and updated output rows.
- Commit boundary: `plan.errors.length===0` before the first write.

- [ ] **Step 1: 旧列・入荷日を必要数へ加えない失敗テストへ更新する**

Replace the old `取り置き数` assertions in `project24_torioki.test.js` with:

```javascript
test('必要数は取り置き台帳数量と今回計画数量だけを差し引く', () => {
  assert.strictEqual(context.残必要計算_({qty:3,取り置き中数量:1,alloc:0}),2);
  assert.strictEqual(context.残必要計算_({qty:3,取り置き中数量:1,alloc:2}),0);
});

test('旧取り置き数・入荷日・履歴数量は必要数へ加えない', () => {
  assert.strictEqual(context.残必要計算_({qty:3,取り置き数:3,入荷:true,履歴Alloc:3,取り置き中数量:0,alloc:0}),3);
});

test('取り置き中で全数確保された取り寄せ行は出荷準備OK', () => {
  assert.strictEqual(context.注文出荷準備OK_([{kbn:'取り寄せ',qty:2,取り置き中数量:2,alloc:0,キャンセル:false}]),true);
});
```

Add these exact regressions:

```javascript
test('開始前取り置き5商品を今回EMSへ再引当しない', () => {
  const codes=['KRBLCM16-02','OFW300','OFW301','MZBGD03','JNXGD01'];
  const orders=codes.map((code,index)=>({ban:String(10100+index),code,sku:code+'b',qty:1,sortKey:10100+index,i:index,keys:[code]}));
  const ledger=orders.map(order=>({取置ID:'INIT|'+context.取り置き_行キー_(order),状態:'取り置き中',受注番号:order.ban,
    商品コード:order.code,SKU:order.sku,取り置き数量:1,取置元種別:'開始前在庫'}));
  const supplies=codes.map(code=>({ems:'EGNEW',code,qty:1,arrival:'2026-07-12'}));
  const result=context.取り置き_割当計算_({orders,ledger,movements:[],supplies,explicit:[]});
  assert.strictEqual(result.newRows.length,0);
  assert.strictEqual(result.surplus.length,5);
  result.surplus.forEach(row=>assert.strictEqual(row.qty,1));
});

test('POEM65とRECIPE42の受注番号タグは指定注文へだけ入る', () => {
  const orders=[
    {ban:'10116569',code:'POEM65',sku:'POEM65b',qty:1,sortKey:1,i:0,keys:['POEM65']},
    {ban:'10117126',code:'RECIPE42',sku:'RECIPE42b',qty:1,sortKey:2,i:1,keys:['RECIPE42']},
    {ban:'10199999',code:'POEM65',sku:'POEM65b',qty:1,sortKey:3,i:2,keys:['POEM65']}
  ];
  const supplies=[
    {ems:'EG1',code:'POEM65',qty:1,directBan:'10116569'},
    {ems:'EG1',code:'RECIPE42',qty:1,directBan:'10117126'}
  ];
  const result=context.取り置き_割当計算_({orders,ledger:[],movements:[],supplies,explicit:[]});
  assert.strictEqual(JSON.stringify(result.newRows.map(r=>r.受注番号)),JSON.stringify(['10116569','10117126']));
  assert.strictEqual(result.orders.find(o=>o.ban==='10199999').need,1);
});
```

- [ ] **Step 2: テストが現行推測ロジックで失敗することを確認する**

Run:

```powershell
node tests\project24_torioki.test.js
node tests\project24_reservation_domain.test.js
```

Expected: 旧 `取り置き数` または `入荷/履歴` が必要数へ入るためFAILする。

- [ ] **Step 3: 必要数と表示状態を台帳数量へ切り替える**

Replace the helpers with:

```javascript
function 残必要計算_(l){
  return Math.max(0,(Number(l&&l.qty)||0)-(Number(l&&l.取り置き中数量)||0)-(Number(l&&l.alloc)||0));
}

function 注文出荷準備OK_(arr){
  const active=(arr||[]).filter(l=>l&&!l.キャンセル);
  return active.length>0 && active.every(l=>l.kbn==='即納'||(l.kbn==='取り寄せ'&&残必要計算_(l)===0));
}

function 引当行状態_(l,cfg){
  if(l.キャンセル) return {st:'キャンセル',color:cfg.色_グレー};
  if(l.kbn==='即納') return {st:'即納',color:cfg.色_水};
  if(l.kbn==='指定なし') return {st:'要確認',color:cfg.色_橙};
  if((Number(l.alloc)||0)>0) return {st:'引当(今回)',color:cfg.色_黄};
  if((Number(l.取り置き中数量)||0)>0) return {st:'取り置き中',color:cfg.色_着};
  return {st:'在庫待ち',color:null};
}
```

- [ ] **Step 4: `引当実行_本体_` を読み取り→計画→検証→反映へ並べ替える**

Add the supply adapter next to `EMS明細_`:

```javascript
function EMS供給オブジェクト_(rows, cols, arrivalFn){
  return (rows||[]).map(row=>{
    const ems=String(row[cols.EMS番号]||'').trim(), raw=String(row[cols.コード]||'').trim(), qty=Number(row[cols.数量])||0;
    if(!実EMS番号_(ems) || !raw || qty<=0) return null;
    return {ems,code:normCode_(raw),qty,arrival:arrivalFn(row),directBan:注文番号在庫コード_(raw)||タグ受注番号_(raw)};
  }).filter(Boolean);
}
```

Change the signature to `引当実行_本体_(options)` with `options=options||{}` and use this exact high-level order inside the function:

```javascript
// A. 受注明細をlinesへ一括読込
const ledgerRows=取り置き台帳_読む_();
const movementRows=EMS在庫移動台帳_読む_();
const ledgerSummary=取り置き_集計_(ledgerRows,movementRows);
lines.forEach(l=>{ l.取り置き中数量=ledgerSummary.activeByKey[取り置き_行キー_(l)]||0; });

// B. P列は書かずに計画だけ作成
const pPlan=発注共有P列計画_();
if(pPlan.error){ ui.alert(pPlan.error); return; }

// C. EMS明細を supplies へ変換して純粋割当
const allocationPlan=取り置き_割当計算_({
  orders:lines.filter(l=>l.kbn==='取り寄せ'&&!l.キャンセル).map(l=>({
    ban:l.ban,code:l.code,sku:l.sku,qty:l.qty,sortKey:l.sortKey,i:l.i,keys:candKeys(l),paid:l.paid
  })),
  ledger:ledgerRows, movements:movementRows, supplies:EMS供給オブジェクト_(E,EC,到着_), explicit:P列計画_確定割当_(pPlan)
});
if(allocationPlan.errors.length){
  ui.alert('引当を中止しました',allocationPlan.errors.join('\n'),ui.ButtonSet.OK);
  return;
}

// preview=true は計画を返すだけで、P列・台帳・受注明細・出力へ書かない
if(options.preview) return {allocationPlan,pPlan,lines,ledgerRows,movementRows};

// D. 検証成功後だけ書込
const now=new Date();
取り置き台帳_割当計画を反映_(allocationPlan,ledgerRows,now);
発注共有P列計画を反映_(pPlan);
```

Keep the public wrapper as `function 引当実行(){ 直列_(()=>引当実行_本体_({preview:false})); }`. Add `引当切替差分を作成()` which calls preview mode, compares the returned plan with the current output by `受注番号|商品コード|SKU`, and writes only `引当切替差分`. Preview mode must not call `消込台帳更新_`, `発注共有P列計画を反映_`, any ledger save, any受注明細 write, or output-sheet write.

`取り置き台帳_割当計画を反映_` must apply every field present in `returnUpdates`（戻し結果、部分再引当後の残数量、メモを含む）, then upsert each `newRows` by deterministic `取置ID`; existing same ID is replaced, not appended. Set `登録日時` only for new IDs and `更新日時` on every changed row. Save the whole ledger once.

Delete these quantity deductions from the allocation path:

```text
入荷日あり行を今回在庫から差し引く処理
消込台帳の出荷済みを到着日・発送日で今回在庫から差し引く処理
引当履歴を現在注文の必要数から差し引く処理
取り置き出荷メモによる今回箱消費の補正
```

Keep the sheets and data for audit only. Current EMS consumption is reconstructed from `取り置き台帳` source EMS and `EMS在庫移動台帳`.

- [ ] **Step 5: 出力シートを同じ計画から生成する**

For each order line, copy the post-plan active quantity from existing ledger plus `allocationPlan.newRows`. `今回入荷EMSの在庫` consumers must come from ledger rows whose `元EMS番号` matches the row EMS; `日本在庫` must be exactly `allocationPlan.surplus`. Continue to show actual matched branch code through `注文一覧表示コード_`; do not modify the original受注明細 product code.

Set `PropertiesService['引当_整合状態']` only after ledger, P, and output writes all complete. Store `{ts,要確認:0,台帳版:'v1'}`.

- [ ] **Step 6: ④回帰テストを通す**

Run:

```powershell
node tests\project24_reservation_domain.test.js
node tests\project24_reservation_workflow.test.js
node tests\project24_torioki.test.js
node tests\project24_arrived_box_color.test.js
node tests\project24_zenken_kensan.test.js
node tests\project24_daniel_amari.test.js
node --check Project_24\引当.js
node --check Project_24\取り置き台帳.js
node --check Project_24\取り置き計算.js
```

Expected: 全テストPASS。`KRBLCM16-02` が余り、`YMNGD09` が余り2、`JMEE167` が7+3、注文番号 `10117375` が2個、OFW4枝番が独立する。

- [ ] **Step 7: Task 7をコミットする**

```powershell
git add Project_24\引当.js Project_24\取り置き台帳.js tests\project24_torioki.test.js tests\project24_arrived_box_color.test.js tests\project24_reservation_domain.test.js
git commit -m "feat(hikiate): 引当実行を取り置き台帳基準へ切替"
```

---

---

### Task 8: ⑤Yahoo移動を冪等なEMS在庫移動台帳へ記録する

**Files:**
- Modify: `Project_24/取り置き台帳.js`
- Modify: `Project_24/引当履歴.js` function `到着済を在庫反映済みへ本体_`
- Modify: `Project_24/引当.js` function `onOpen`
- Modify: `tests/project24_reservation_workflow.test.js`

**Interfaces:**
- Consumes: ④検証済み `日本在庫` surplus、選択EMS番号、現物確認済みYahoo戻し候補。
- Produces: `EMS在庫移動_箱計画_(surplus, existing)`, `EMS在庫移動_戻し計画_(returns, existing)`, `EMS在庫移動台帳_保存_(rows)`。
- Process IDs: box surplus `YAHOO|EMS|<EMS>|<CODE>`; cancellation return `YAHOO|RETURN|<取置ID>`。

- [ ] **Step 1: 同じ処理IDを二重登録しない失敗テストを追加する**

```javascript
test('箱余りのYahoo移動は同じEMS・商品で一度だけ記録する', () => {
  const surplus=[{ems:'EG1',code:'AAA',qty:2}];
  const first=context.EMS在庫移動_箱計画_(surplus,[],'2026-07-15 10:00:00');
  const second=context.EMS在庫移動_箱計画_(surplus,first.rows,'2026-07-15 10:01:00');
  assert.strictEqual(JSON.stringify(first.errors),'[]');
  assert.strictEqual(first.added.length,1);
  assert.strictEqual(first.added[0].処理ID,'YAHOO|EMS|EG1|AAA');
  assert.strictEqual(second.added.length,0);
  assert.strictEqual(second.rows.length,1);
});

test('キャンセル戻しは取置IDごとに一度だけYahoo移動する', () => {
  const returns=[{取置ID:'OLD1',状態:'キャンセル戻し',戻し処理結果:'現物あり',商品コード:'AAA',取り置き数量:1,元EMS番号:'EG0'}];
  const first=context.EMS在庫移動_戻し計画_(returns,[],'2026-07-15 10:00:00');
  const second=context.EMS在庫移動_戻し計画_(returns,first.rows,'2026-07-15 10:01:00');
  assert.strictEqual(first.added[0].処理ID,'YAHOO|RETURN|OLD1');
  assert.strictEqual(second.added.length,0);
});
```

- [ ] **Step 2: 純粋な移動計画を実装する**

```javascript
function EMS在庫移動_追加計画_(candidates, existing, now){
  const ids=new Set((existing||[]).map(r=>String(r.処理ID||''))), added=[], errors=[];
  (candidates||[]).forEach(c=>{
    if(!c.処理ID || !取り置き_整数_(c.数量)) errors.push('Yahoo移動の処理IDまたは数量が不正: '+String(c.処理ID||''));
    else if(!ids.has(c.処理ID)){ ids.add(c.処理ID); added.push(Object.assign({},c,{移動先:'Yahoo即納',処理日時:now})); }
  });
  return {rows:(existing||[]).map(r=>Object.assign({},r)).concat(added),added,errors};
}

function EMS在庫移動_箱計画_(surplus, existing, now){
  return EMS在庫移動_追加計画_((surplus||[]).filter(s=>s.qty>0).map(s=>({
    処理ID:'YAHOO|EMS|'+String(s.ems||'').trim()+'|'+normCode_(s.code),EMS番号:String(s.ems||'').trim(),商品コード:normCode_(s.code),数量:s.qty
  })),existing,now);
}

function EMS在庫移動_戻し計画_(returns, existing, now){
  return EMS在庫移動_追加計画_((returns||[]).filter(r=>r.状態==='キャンセル戻し'&&r.戻し処理結果==='現物あり').map(r=>({
    処理ID:'YAHOO|RETURN|'+r.取置ID,EMS番号:String(r.元EMS番号||''),商品コード:normCode_(r.元EMS商品コード||r.商品コード),数量:r.取り置き数量
  })),existing,now);
}
```

- [ ] **Step 3: 便締めの確認後だけ箱余りを記録する**

In `到着済を在庫反映済みへ本体_`, after target EMS validation and before changing external EMS statuses:

1. Read the latest④ timestamp and require `要確認===0` and `台帳版==='v1'`; remove the current “それでも締める” override for ledger validation errors.
2. Read `日本在庫` rows for selected EMS only.
3. Show `EMS番号 / 商品コード / Yahooへ移す数量` and the text `Yahoo在庫への反映が完了している場合だけOK`.
4. Build the movement plan and stop if errors exist.
5. After the external status change succeeds, save the movement rows once.

Wrap external status change and movement save in one `try/catch`. If external status update fails, do not save movement rows. If movement save fails after external update, immediately restore the same `flipA1` cells to `到着済`, call `SpreadsheetApp.flush()`, and stop before history promotion or EMS refresh. The deterministic process IDs still prevent duplicates on retry.

- [ ] **Step 4: キャンセル戻しのYahoo確定入口を追加する**

Add menu item:

```javascript
.addItem('🛒 Yahoo戻しを反映済みにする', 'キャンセル戻しをYahoo反映済みにする')
```

Implement the public function as follows. It does not call a Yahoo API or add stock automatically.

```javascript
function キャンセル戻しをYahoo反映済みにする(){
  const ui=SpreadsheetApp.getUi(), candidates=取り置き_表を読む_(TORIOKI_CFG.Yahoo候補,Yahoo候補HDR).filter(r=>r.確認===true||String(r.確認).toUpperCase()==='TRUE');
  if(!candidates.length){ ui.alert('Yahoo戻し候補にチェックがありません'); return; }
  const answer=ui.alert('Yahoo戻しの最終確認','Yahoo在庫へ実際に加算済みの'+candidates.length+'件だけを確定します。',ui.ButtonSet.OK_CANCEL);
  if(answer!==ui.Button.OK) return;
  const ledger=取り置き台帳_読む_(), selectedIds=new Set(candidates.map(r=>String(r.取置ID)));
  const returns=ledger.filter(r=>selectedIds.has(String(r.取置ID)));
  if(returns.some(r=>r.状態!=='キャンセル戻し'||r.戻し処理結果!=='現物あり')){ ui.alert('現物ありでない候補が含まれるため中止しました'); return; }
  const existingMoves=EMS在庫移動台帳_読む_(), plan=EMS在庫移動_戻し計画_(returns,existingMoves,new Date());
  if(plan.errors.length){ ui.alert('Yahoo移動を中止しました',plan.errors.join('\n'),ui.ButtonSet.OK); return; }
  const updatedLedger=ledger.map(r=>selectedIds.has(String(r.取置ID))?Object.assign({},r,{戻し処理結果:'Yahoo反映済み',更新日時:new Date()}):r);
  try{
    EMS在庫移動台帳_保存_(plan.rows);
    try{ 取り置き台帳_保存_(updatedLedger); }
    catch(error){ EMS在庫移動台帳_保存_(existingMoves); throw error; }
  }catch(error){ ui.alert('Yahoo戻し記録に失敗したため台帳を元へ戻しました',error.message,ui.ButtonSet.OK); return; }
  Yahoo戻し候補を更新_();
}
```

- [ ] **Step 5: テストと構文を通す**

Run:

```powershell
node tests\project24_reservation_workflow.test.js
node tests\project24_reservation_domain.test.js
node --check Project_24\取り置き台帳.js
node --check Project_24\引当履歴.js
node --check Project_24\引当.js
```

Expected: 同じ処理IDの2回目は追加0、全テストPASS、構文エラーなし。

- [ ] **Step 6: Task 8をコミットする**

```powershell
git add Project_24\取り置き台帳.js Project_24\引当履歴.js Project_24\引当.js tests\project24_reservation_workflow.test.js
git commit -m "feat(hikiate): Yahoo移動を冪等台帳へ記録"
```

---

---

### Task 9: 全件検算と運用出力を新台帳の同一式へ統一する

**Files:**
- Modify: `Project_24/全件検算.js`
- Modify: `tests/project24_zenken_kensan.test.js`
- Modify: `tests/project24_reservation_domain.test.js`

**Interfaces:**
- Consumes: EMSリスト全状態、取り置き台帳、EMS在庫移動台帳、Yahoo CSV。
- Produces: `全件検算_台帳集計_(src)` with per supply key invariant and product summary.
- Invariant: `供給 = 取り置き中 + 発送済み + 戻し未処理 + 在庫なし確定 + Yahoo移動済み + 余り`。

- [ ] **Step 1: 新しい突合式の失敗テストを追加する**

```javascript
test('箱別供給が新台帳の全区分と一致すればOK', () => {
  const result=context.全件検算_台帳集計_({
    supply:[{ems:'EG1',code:'AAA',qty:10,status:'到着済'}],
    ledger:[
      {取置ID:'A',状態:'取り置き中',受注番号:'101',商品コード:'AAA',SKU:'AAAb',取り置き数量:3,取置元種別:'EMS',元EMS番号:'EG1'},
      {取置ID:'B',状態:'発送済み',受注番号:'102',商品コード:'AAA',SKU:'AAAb',取り置き数量:2,取置元種別:'EMS',元EMS番号:'EG1'},
      {取置ID:'C',状態:'キャンセル戻し',戻し処理結果:'現物あり',受注番号:'103',商品コード:'AAA',SKU:'AAAb',取り置き数量:1,取置元種別:'EMS',元EMS番号:'EG1'},
      {取置ID:'D',状態:'キャンセル戻し',戻し処理結果:'在庫なし',受注番号:'104',商品コード:'AAA',SKU:'AAAb',取り置き数量:1,取置元種別:'EMS',元EMS番号:'EG1'}
    ],
    movements:[{処理ID:'YAHOO|EMS|EG1|AAA',EMS番号:'EG1',商品コード:'AAA',数量:2}], yahoo:null
  });
  const row=result.rows[0];
  assert.strictEqual(row.供給,10);
  assert.strictEqual(row.取り置き中,3);
  assert.strictEqual(row.発送済み,2);
  assert.strictEqual(row.戻し未処理,1);
  assert.strictEqual(row.在庫なし確定,1);
  assert.strictEqual(row.Yahoo移動済み,2);
  assert.strictEqual(row.余り,1);
  assert.strictEqual(row.判定,'OK');
});

test('再引当済み元行とYahoo反映済み元行を二重計上しない', () => {
  const result=context.全件検算_台帳集計_({
    supply:[{ems:'EG1',code:'AAA',qty:2,status:'到着済'}],
    ledger:[
      {取置ID:'OLD1',状態:'キャンセル戻し',戻し処理結果:'再引当済み',受注番号:'100',商品コード:'AAA',取り置き数量:1,取置元種別:'EMS',元EMS番号:'EG1'},
      {取置ID:'NEW1',状態:'取り置き中',受注番号:'200',商品コード:'AAA',取り置き数量:1,取置元種別:'キャンセル再引当',元EMS番号:'EG1',元取置ID:'OLD1'},
      {取置ID:'OLD2',状態:'キャンセル戻し',戻し処理結果:'Yahoo反映済み',受注番号:'300',商品コード:'AAA',取り置き数量:1,取置元種別:'EMS',元EMS番号:'EG1'}
    ],
    movements:[{処理ID:'YAHOO|RETURN|OLD2',EMS番号:'EG1',商品コード:'AAA',数量:1}],yahoo:null
  });
  assert.strictEqual(result.rows[0].取り置き中,1);
  assert.strictEqual(result.rows[0].Yahoo移動済み,1);
  assert.strictEqual(result.rows[0].余り,0);
});
```

- [ ] **Step 2: 純粋な台帳検算を実装する**

Add this pure aggregation. A missing ledger/movement header in the sheet-reading wrapper is a blocking error; Yahoo read failure remains comparison-only warning.

```javascript
function 全件検算_台帳集計_(src){
  const supplyByKey={}, meta={};
  (src.supply||[]).forEach(r=>{
    const ems=String(r.ems||''), code=normCode_(r.code), qty=Number(r.qty)||0;
    if(!実EMS番号_(ems) || !code || qty<=0) return;
    const key=取り置き_供給キー_(ems,code);
    supplyByKey[key]=(supplyByKey[key]||0)+qty; meta[key]={ems,code};
  });
  const summary=取り置き_集計_(src.ledger||[],src.movements||[]);
  if(summary.errors.length) return {rows:[],errors:summary.errors};
  const rows=Object.keys(supplyByKey).sort().map(key=>{
    const u=summary.usageBySupply[key]||{取り置き中:0,発送済み:0,戻し未処理:0,在庫なし確定:0,Yahoo移動済み:0};
    const supplied=supplyByKey[key], used=取り置き_使用合計_(u), rest=supplied-used;
    const yahoo=src.yahoo&&src.yahoo[meta[key].code]!=null?Number(src.yahoo[meta[key].code]):null;
    return {EMS番号:meta[key].ems,商品コード:meta[key].code,判定:rest<0?'⚠️超過消費':'OK',供給:supplied,
      取り置き中:u.取り置き中||0,発送済み:u.発送済み||0,戻し未処理:u.戻し未処理||0,在庫なし確定:u.在庫なし確定||0,
      Yahoo移動済み:u.Yahoo移動済み||0,余り:rest,'Yahoo a在庫':yahoo,差:yahoo==null?'':yahoo-rest};
  });
  return {rows,errors:[]};
}
```

Change report headers to:

```javascript
const HDR=['EMS番号','商品コード','判定','供給','取り置き中','発送済み','戻し未処理','在庫なし確定','Yahoo移動済み','余り','Yahoo a在庫','差'];
```

Remove the old arrival-date buckets `出荷済(到着箱)`, `出荷済(過去便)`, `確保済(ズレ)` from the main judgment. If historical visibility is still useful, render it below the new report as `参考情報` and do not use it in `判定`.

- [ ] **Step 3: ④完了ダイアログと今回EMS出力の式を同じ文言へ統一する**

Use exactly:

```text
供給N＝取り置き中A＋発送済みB＋戻し未処理C＋在庫なし確定D＋Yahoo移動済みE＋余りF
```

If the equation fails, store `引当_整合状態.要確認 > 0`; ⑤ must remain blocked.

- [ ] **Step 4: 検算・既存回帰・構文を通す**

Run:

```powershell
node tests\project24_zenken_kensan.test.js
node tests\project24_reservation_domain.test.js
node tests\project24_reservation_workflow.test.js
node tests\project24_torioki.test.js
node tests\project24_arrived_box_color.test.js
node tests\project24_daniel_amari.test.js
node --check Project_24\全件検算.js
```

Expected: 全テストPASS。Yahoo CSVがなくても箱別数量判定は実行され、Yahoo比較だけが警告になる。

- [ ] **Step 5: Task 9をコミットする**

```powershell
git add Project_24\全件検算.js tests\project24_zenken_kensan.test.js tests\project24_reservation_domain.test.js
git commit -m "feat(hikiate): 全件検算を取り置き台帳式へ統一"
```

---

---

### Task 10: 全回帰・速度・安全反映・初回移行を完了する

**Files:**
- Verify: every `Project_24/*.js`
- Verify: every `tests/project24*.test.js`
- May update automatically: `tools/sync_state/Project_24.json`
- Live sheets created: `取り置き台帳`, `取り置き初期登録`, `キャンセル戻し確認`, `Yahoo戻し候補`, `EMS在庫移動台帳`, `取り置き要確認`

**Interfaces:**
- Consumes: Tasks 1-9のコードと承認済み設計。
- Produces: 安全反映済みProject_24、利用者が現物数を入力できる初期登録シート、旧方式との差分プレビュー。

- [ ] **Step 1: すべてのProject_24テストを列挙して実行する**

Run:

```powershell
Get-ChildItem tests -Filter 'project24*.test.js' -File | Sort-Object Name | ForEach-Object { node $_.FullName; if($LASTEXITCODE -ne 0){ throw "test failed: $($_.Name)" } }
```

Expected: each test file exits 0 and prints only PASS lines.

- [ ] **Step 2: Project_24全JavaScriptの構文を検査する**

Run:

```powershell
Get-ChildItem Project_24 -Filter '*.js' -File | Sort-Object Name | ForEach-Object { node --check $_.FullName; if($LASTEXITCODE -ne 0){ throw "syntax error: $($_.Name)" } }
```

Expected: exit 0, no syntax errors.

- [ ] **Step 3: 速度計測ログとAPI回数を確認する**

Confirm by source inspection and local workflow tests:

```text
①: 受注明細本体1回書込 + 取り置き台帳1回書込 + 要確認シート最大1回書込
④: 発注共有P列1回書込 + 取り置き台帳1回書込 + 各出力シート1回書込
ループ内のgetRange/setValueは禁止
console.logに「①受注明細更新 処理時間ms」「④引当実行 処理時間ms」を残す
```

Run `rg -n "forEach.*getRange|for\s*\(.*getRange|setValue\(" Project_24\取り置き台帳.js Project_24\取り置き計算.js Project_24\引当.js Project_24\P列自動記入.js` and inspect every hit. Row-by-row writes found in the new paths must be converted to array `setValues` or `RangeList` before continuing.

- [ ] **Step 4: 差分と作業ツリーを確認する**

Run:

```powershell
git status --short
git diff --check
git log --oneline --max-count=12
```

Expected: planned Project_24/tests/docs changes only; `git diff --check` has no output.

- [ ] **Step 5: Project_24を安全反映する**

Run:

```powershell
powershell -ExecutionPolicy Bypass -File tools\gas_safe_push.ps1 Project_24
```

Expected: online-only change conflictなしでProject_24へ反映され、sync stateが更新される。競合が出た場合は上書きを選ばず中止し、`gas_pull_sync.ps1` で取り込んでテストをやり直す。

- [ ] **Step 6: 初期登録シートを作り、利用者入力で停止する**

In the live spreadsheet:

1. Run `取り置き初期登録を作成`.
2. Confirm candidates come only from current `部分在庫` and `希望日待ち`.
3. Do not run ④.
4. Ask the user to inspect the shelf and enter exact `現物取り置き数量`.

Expected: this is the only manual full entry. After confirmation, ①/④ maintain the ledger; future manual work is limited to cancellation physical confirmation and exceptional `手動解除`.

- [ ] **Step 7: 初期登録を確定し、新旧差分をプレビューする**

After the user finishes quantities:

1. Run `取り置き初期登録を確定`.
2. Run the public function `引当切替差分を作成`; it calls `引当実行_本体_({preview:true})` and performs no operational writes.
3. Confirm it exports `受注番号 / 商品コード / 旧割当数量 / 新取り置き中 / 新規EMS割当 / 余り` to `引当切替差分`.
4. Verify the named regression rows: `KRBLCM16-02`, `YMNGD09`, `JMEE167`, `OFW304-1`, `OFW304-2`, `OFW305-1`, `OFW305-2`, `POEM65`, `RECIPE42`, `10117375`.
5. If any unexplained difference exists, do not run live ④; fix the calculation and rerun all tests.

- [ ] **Step 8: 本番④を1回実行し、再実行の冪等性を確認する**

1. Run ④ once and record counts and processing seconds.
2. Without changing source data, run ④ again.
3. Confirm `取り置き台帳` active total, EMS movement total, and each box surplus are unchanged.
4. Confirm ① and④ are each below two minutes; target is ① under 30 seconds and④ under 60 seconds on the current row count. If either exceeds two minutes, stop rollout and capture execution logs before optimization.

- [ ] **Step 9: 全件検算と安全同期を確認する**

Run the live `全件検算レポート`, then:

```powershell
powershell -ExecutionPolicy Bypass -File tools\gas_pull_sync.ps1 Project_24
git status --short
```

Expected: all box equations are OK, Project_24 is online/local identical, and no unexplained worktree changes remain.

- [ ] **Step 10: rollout結果をコミットする**

If `tools/sync_state/Project_24.json` changed:

```powershell
git add tools\sync_state\Project_24.json
git commit -m "chore(Project_24): 取り置き台帳版の同期基準を更新"
```

Do not commit spreadsheet data exports containing customer names, addresses, phone numbers, or order details.

---

## Execution Checkpoints

1. Tasks 1-5 complete: ledger creation, initial entry, CSV transitions, cancellation confirmation work without changing ④.
2. Tasks 6-7 complete: P preview and④ use the new ledger; stop and review domain tests before any live push.
3. Tasks 8-9 complete: Yahoo movement and full reconciliation use the same source-of-truth equation.
4. Task 10 Step 6: implementation pauses for the user's one-time physical quantity entry.
5. Task 10 Steps 7-10: preview, live cutover, rerun idempotency, and safe sync finish the migration.
