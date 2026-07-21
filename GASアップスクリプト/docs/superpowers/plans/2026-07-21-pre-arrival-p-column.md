# 未着便のP列先行記入＋受注明細「入荷予定」列 実装計画

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 未着（発送前・輸送中）のEMSリスト行に予定の注文番号を薄青背景で先行記入し、受注明細に「入荷予定」列（`7/23(…9766)` 形式）を自動生成して、箱が来る前に「誰の分か・いつ着くか」を見えるようにする。

**Architecture:** 既存のP列計画エンジン（P列計画_純計算_・無変更）に渡す供給行を「確定対象(到着済)→計画対象(未着)」の2パス順に並べ替えるだけで先行割当を実現。確定(②P列確定)・出荷(④)は従来どおり到着済のみ読むので物理保証は不変。受注明細側は計画結果から純粋関数でマップを作り、1列を毎回全書き直しする。

**Tech Stack:** Google Apps Script（V8）、テストは素のnode＋vm（`node tests/<file>.test.js`）。

**Spec:** `docs/superpowers/specs/2026-07-21-pre-arrival-p-column-design.md`

## Global Constraints

- 作業場所: worktree `GASアップスクリプト/.worktrees/full-allocation-rebuild-v3`（ブランチ `codex/full-allocation-rebuild-v3`）。作業前に `git pull origin codex/full-allocation-rebuild-v3` → `tools/gas_pull_sync.ps1 Project_24`。編集前に対象ファイルのmtimeを確認（ローカルCodexの並行編集対策。数分以内の個別更新があればユーザーに確認）
- 素の `clasp push` は禁止。本番反映は `tools/gas_safe_push.ps1 Project_24` ＋ scratchへ `.clasp.json` をコピーして `clasp pull` し全ファイル一致を独立検証
- **本番pushは「全件再計算の反映（残作業A）」がユーザー操作で完了してから**。それまでコミットのみ
- テスト実行はworktreeの `GASアップスクリプト` 直下で `node tests/<file>.test.js`。既存8テストファイルの回帰全PASS必須
- GASの書き込み系エントリポイントは `直列_` 必須（本計画は既存の直列_内フローに乗るだけで新規エントリポイントなし）
- google.script.run から呼ぶ関数名は末尾 `_` 禁止（本計画では該当なし）
- プロパティ名の「・」(U+30FB) は引用符で書く（本計画では該当なし）
- コミットは日本語1行サマリ＋ `Co-Authored-By: Claude Fable 5 <noreply@anthropic.com>`

## 前提（コード現状・2026-07-21時点）

- `Project_24/P列自動記入.js:156-158` `P列処理対象EMS_(status)` = `status==='到着済'` のみ
- `Project_24/P列自動記入.js:254-259` `発注共有P列計画を反映_(plan)` が唯一の反映funnel（呼び元: 引当.js:1627 と P列自動記入.js:416）。`P列書き直し実行_`（33-68行）だけは直接setValuesする別経路
- `Project_24/P列自動記入.js:261-403` `発注共有P列計画_(options)` が計画本体。345行で `status` 抽出済み、352-355行でrows構築、365-371行が手動名指し救済、382-383行が純計算呼び出し、390-391行が背景、392-398行がsummary、399-402行がreturn
- `Project_24/取り置き計算.js:124` `P列計画_純計算_(emsRows, inputOrders, fixedBySupply, usageBySupply)` は供給行を**渡された順**に消費する純粋関数（変更しない）
- `P列指定文字列_` の出力: 全量1注文=`10117602`、分割=`10117602:16, 10117731:4`、余りはP列に書かない
- 受注明細の取込（引当.js:515-537 取込_実行_）は**ヘッダー行を保持**しデータ域だけ入れ替える。カスタム列の値は消えるがヘッダーは残る→「入荷予定」は右端自動作成・②で毎回全再生成でよい
- テストvmコンテキストは `tests/project24_reservation_workflow.test.js:7-30` の形式（SpreadsheetApp等スタブ＋7ファイル読込）を踏襲する

---

### Task 1: 純粋関数4つ＋新テストファイル

**Files:**
- Modify: `Project_24/P列自動記入.js`（`P列処理対象EMS_` の直後、現156-158行付近に追記）
- Create: `tests/project24_pre_arrival_plan.test.js`

**Interfaces:**
- Produces: `P列計画対象EMS_(status:string):boolean` — 未着判定（到着済/在庫反映済み以外）
- Produces: `P列計画行順_(rows:Array):Array` — (対象||計画)&&!全件再計算ブロック を対象先→計画後で返す
- Produces: `入荷予定表記_(boxes:[{arrival:'YYYY-MM-DD'|'',ems:string}]):string` — `7/23(…9766)+7/26(…9011)ほか` 形式
- Produces: `入荷予定マップ_(planRows, lines):{[sheetRow:number]:string}` — 計画行entriesを受注行へ引き当てる
- Consumes: `P列計画_純計算_`（取り置き計算.js・既存）

- [ ] **Step 1: 新テストファイルを書く（RED想定）**

`tests/project24_pre_arrival_plan.test.js` を以下の完全な内容で作成:

```js
// 未着便のP列先行記入（計画対象）と受注明細「入荷予定」の純粋ロジックテスト
// 実行: GASアップスクリプト直下で node tests/project24_pre_arrival_plan.test.js
const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

const activeSpreadsheet = { getSheetByName: () => null };
const context = {
  console,
  Date,
  Set,
  SpreadsheetApp: {
    openById: () => ({ getSheetByName: () => null }),
    getActive: () => activeSpreadsheet
  },
  LockService: { getDocumentLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
  PropertiesService: { getDocumentProperties: () => ({ getProperty: () => null, setProperty: () => {}, deleteProperty: () => {} }) },
  Utilities: { formatDate: () => '' },
  Logger: { log: () => {} }
};
vm.createContext(context);
[
  'Project_24/引当.js',
  'Project_24/引当履歴.js',
  'Project_24/取り置き計算.js',
  'Project_24/取り置き台帳.js',
  'Project_24/消込台帳.js',
  'Project_24/P列自動記入.js',
  'Project_24/全件再計算.js'
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

test('計画対象判定: 到着済/在庫反映済み以外(空欄含む)を未着扱いにする', () => {
  assert.strictEqual(context.P列計画対象EMS_('到着済'), false);
  assert.strictEqual(context.P列計画対象EMS_('在庫反映済み'), false);
  assert.strictEqual(context.P列計画対象EMS_('発送済み'), true);
  assert.strictEqual(context.P列計画対象EMS_(''), true);
  assert.strictEqual(context.P列計画対象EMS_(null), true);
});

test('行順: 到着済が先・未着が後・ブロックSKUと対象外行は除外', () => {
  const rows = [
    {i: 0, 計画: true},
    {i: 1, 対象: true},
    {i: 2, 対象: true, 全件再計算ブロック: true},
    {i: 3} // 在庫反映済みなど(どちらでもない)
  ];
  const ordered = context.P列計画行順_(rows);
  assert.deepStrictEqual(ordered.map(r => r.i), [1, 0]);
});

test('未着FIFO: 到着済で足りない古い注文が未着便へ流れ、到着済の割当は単独時と同じ', () => {
  const arrived = {ems: 'EG-A', code: 'AAA', qty: 1, pOriginal: '', arrival: '2026-07-18', 対象: true};
  const future = {ems: 'EG-F', code: 'AAA', qty: 2, pOriginal: '', arrival: '2026-07-25', 計画: true};
  const orders = () => ([
    {ban: '101', code: 'AAA', sku: 'AAAb', qty: 2, need: 2, date: new Date('2026-07-01'), keys: ['AAA'], row: 10},
    {ban: '102', code: 'AAA', sku: 'AAAb', qty: 1, need: 1, date: new Date('2026-07-02'), keys: ['AAA'], row: 11}
  ]);
  const both = context.P列計画_純計算_(context.P列計画行順_([arrived, future]), orders(), {}, {});
  assert.strictEqual(both.rows[0].nextP, '101'); // 到着済: 全量101
  assert.strictEqual(both.rows[1].nextP, '101:1, 102:1'); // 未着: 101の残り1+102の1
  const alone = context.P列計画_純計算_(context.P列計画行順_([Object.assign({}, arrived)]), orders(), {}, {});
  assert.strictEqual(alone.rows[0].nextP, '101'); // 未着を混ぜても到着済の結果は不変
});

test('昇格: 同じ行が未着→到着済に変わっても割当テキストは同値', () => {
  const orders = () => ([{ban: '101', code: 'AAA', sku: 'AAAb', qty: 2, need: 2, date: new Date('2026-07-01'), keys: ['AAA'], row: 10}]);
  const asPlan = context.P列計画_純計算_(context.P列計画行順_([{ems: 'EG-F', code: 'AAA', qty: 2, pOriginal: '', arrival: '2026-07-25', 計画: true}]), orders(), {}, {});
  const asArrived = context.P列計画_純計算_(context.P列計画行順_([{ems: 'EG-F', code: 'AAA', qty: 2, pOriginal: '', arrival: '2026-07-25', 対象: true}]), orders(), {}, {});
  assert.strictEqual(asPlan.rows[0].nextP, asArrived.rows[0].nextP);
  assert.strictEqual(asPlan.rows[0].nextP, '101');
});

test('入荷予定表記: 単便・複数便の到着日順・日付なし・3便以上はほか', () => {
  assert.strictEqual(context.入荷予定表記_([{arrival: '2026-07-23', ems: 'EG050049766KR'}]), '7/23(…9766)');
  assert.strictEqual(
    context.入荷予定表記_([
      {arrival: '2026-07-26', ems: 'EG050099011KR'},
      {arrival: '2026-07-23', ems: 'EG050049766KR'}
    ]),
    '7/23(…9766)+7/26(…9011)');
  assert.strictEqual(context.入荷予定表記_([{arrival: '', ems: 'EG050049766KR'}]), '(…9766)');
  assert.strictEqual(
    context.入荷予定表記_([
      {arrival: '2026-07-23', ems: 'EG050049766KR'},
      {arrival: '2026-07-26', ems: 'EG050099011KR'},
      {arrival: '2026-07-30', ems: 'EG050111222KR'}
    ]),
    '7/23(…9766)+7/26(…9011)ほか');
});

test('入荷予定マップ: 計画行のentriesだけを受注行へ引き当て、到着済行は含めない', () => {
  const planRows = [
    {計画: true, code: 'AAA', arrival: '2026-07-25', ems: 'EG050049766KR', entries: [{ban: '101', qty: 1}]},
    {対象: true, code: 'AAA', arrival: '2026-07-18', ems: 'EG049827401KR', entries: [{ban: '102', qty: 1}]}
  ];
  const lines = [
    {ban: '101', row: 10, keys: ['AAA']},
    {ban: '102', row: 11, keys: ['AAA']}
  ];
  const map = context.入荷予定マップ_(planRows, lines);
  assert.strictEqual(map[10], '7/25(…9766)');
  assert.strictEqual(map[11], undefined);
});

process.exitCode = failures ? 1 : 0;
console.log(failures ? `${failures} FAILED` : 'ALL PASS');
```

- [ ] **Step 2: 実行してREDを確認**

Run: `node tests/project24_pre_arrival_plan.test.js`
Expected: FAIL多数（`context.P列計画対象EMS_ is not a function` 等）

- [ ] **Step 3: 純粋関数4つを実装**

`Project_24/P列自動記入.js` の `P列処理対象EMS_`（現156-158行）の直後に追加:

```js
// 未着(発送前・輸送中・ステータス空欄含む)の計画対象。確定(②)には使わずP列の先行記入だけに使う
function P列計画対象EMS_(status){
  const s=String(status==null?'':status).trim();
  return s!=='到着済' && s!=='在庫反映済み';
}

// 供給の消費順: 確定対象(到着済)を先に、計画対象(未着)を後に。ブロックSKUと対象外は除く
function P列計画行順_(rows){
  const eligible=(rows||[]).filter(r=>r && (r.対象||r.計画) && !r.全件再計算ブロック);
  return eligible.filter(r=>r.対象).concat(eligible.filter(r=>!r.対象));
}

// 未着便の入荷予定表記: 「7/23(…9766)+7/26(…9011)ほか」(到着日昇順・最大2便・EMS番号は数字部末尾4桁)
function 入荷予定表記_(boxes){
  const seen={}, list=[];
  (boxes||[]).forEach(b=>{
    const arrival=String(b&&b.arrival||''), ems=String(b&&b.ems||'');
    const key=ems+'|'+arrival;
    if(seen[key]) return; seen[key]=1; list.push({arrival:arrival,ems:ems});
  });
  list.sort((a,b)=>a.arrival<b.arrival?-1:a.arrival>b.arrival?1:0);
  const fmt=b=>{
    const digits=b.ems.match(/(\d{4})\D*$/);
    const tail=digits?'…'+digits[1]:b.ems;
    const d=b.arrival.match(/^\d{4}-(\d{2})-(\d{2})$/);
    return (d? Number(d[1])+'/'+Number(d[2]) : '')+'('+tail+')';
  };
  return list.slice(0,2).map(fmt).join('+')+(list.length>2?'ほか':'');
}

// 計画対象(未着)行の割当entriesを、受注明細の行番号→入荷予定表記のマップへ変換する純粋関数
function 入荷予定マップ_(planRows, lines){
  const byBanCode={}; // 受注番号|正規化コード → [{arrival,ems}]
  (planRows||[]).forEach(r=>{
    if(!r || !r.計画) return;
    (r.entries||[]).forEach(e=>{
      if(!e || !e.ban) return;
      const key=String(e.ban)+'|'+String(r.code||'');
      (byBanCode[key]=byBanCode[key]||[]).push({arrival:String(r.arrival||''),ems:String(r.ems||'')});
    });
  });
  const out={};
  (lines||[]).forEach(l=>{
    if(!l || !l.row) return;
    const keys=l.keys instanceof Set?Array.from(l.keys):(l.keys||[]);
    let boxes=[];
    keys.forEach(k=>{ boxes=boxes.concat(byBanCode[String(l.ban)+'|'+k]||[]); });
    if(boxes.length) out[l.row]=入荷予定表記_(boxes);
  });
  return out;
}
```

- [ ] **Step 4: テストがGREENになるのを確認**

Run: `node tests/project24_pre_arrival_plan.test.js`
Expected: `ALL PASS`（7件）

- [ ] **Step 5: 既存テストの回帰**

Run: `for f in tests/*.test.js; do node "$f" | grep -c PASS; done`（bash）または各ファイル個別実行
Expected: 全ファイルFAILなし（合計236件PASS＋新規7件）

- [ ] **Step 6: コミット**

```bash
git add "GASアップスクリプト/Project_24/P列自動記入.js" "GASアップスクリプト/tests/project24_pre_arrival_plan.test.js"
git commit -m "feat(Project_24): 未着便P列先行記入の純粋関数4つ(計画対象判定/行順/入荷予定表記/マップ)"
```

---

### Task 2: 発注共有P列計画_ を計画対象へ拡張

**Files:**
- Modify: `Project_24/P列自動記入.js`（発注共有P列計画_ 現261-403行の中の5か所）

**Interfaces:**
- Consumes: Task 1 の `P列計画対象EMS_` / `P列計画行順_`
- Produces: plan オブジェクトに `lines`（row付き受注行）が追加され、`plan.rows` の各行に `計画:boolean` が付く。`summary.予定` = 予定を書いた未着行数。`plan.backgrounds` が全計算行をカバー（自己修復色）

- [ ] **Step 1: rows構築に計画フラグを追加（現352-355行）**

変更前:
```js
    rows.push({
      i,ems:emsNo,code:normCode_(code),sourceCode:code,directBan:注文番号在庫コード_(code)||タグ受注番号_(code),arrival:ymd_(at(row,colA)),qty:Number(at(row,colQ))||0,
      pOriginal,対象:P列処理対象EMS_(at(row,colSt)),status,全件再計算ブロック:rebuildBlocked.has(全件再計算_SKU正規化_(code,'EMS'))
    });
```
変更後（345行で抽出済みの `status` を使う）:
```js
    rows.push({
      i,ems:emsNo,code:normCode_(code),sourceCode:code,directBan:注文番号在庫コード_(code)||タグ受注番号_(code),arrival:ymd_(at(row,colA)),qty:Number(at(row,colQ))||0,
      pOriginal,対象:P列処理対象EMS_(status),計画:P列計画対象EMS_(status),status,全件再計算ブロック:rebuildBlocked.has(全件再計算_SKU正規化_(code,'EMS'))
    });
```

- [ ] **Step 2: lines に受注明細の行番号を追加（現298-301行）**

変更前:
```js
    const line={
      ban, code, sku, qty, kbn:'取り寄せ', キャンセル:false,
      keys, need:0, date, seq:lines.length
    };
```
変更後:
```js
    const line={
      ban, code, sku, qty, kbn:'取り寄せ', キャンセル:false,
      keys, need:0, date, seq:lines.length, row:i+1
    };
```

- [ ] **Step 3a: 名指し救済の失敗テストを追加（RED）**

`tests/project24_pre_arrival_plan.test.js` の `process.exitCode` 行の前に追加:

```js
test('名指し救済: コード不一致の手書き名指しは未着行でも保持しコード一致は保持しない', () => {
  const rows = [
    {計画: true, code: 'ZZZ', pOriginal: '10117999', directBan: ''},
    {計画: true, code: 'AAA', pOriginal: '10117888', directBan: ''},
    {対象: true, code: 'YYY', pOriginal: '10117777', directBan: ''},
    {code: 'XXX', pOriginal: '10117666', directBan: ''} // 在庫反映済み等(どちらでもない)は触らない
  ];
  context.P列名指し救済適用_(rows, new Set(['AAA']));
  assert.strictEqual(rows[0].directBan, '10117999');
  assert.strictEqual(rows[0].手動名指し, true);
  assert.strictEqual(rows[1].directBan, ''); // コード一致=通常ルート(計画が正)
  assert.strictEqual(rows[2].directBan, '10117777'); // 到着済も従来どおり
  assert.strictEqual(rows[3].directBan, '');
});
```

Run: `node tests/project24_pre_arrival_plan.test.js`
Expected: この1件だけFAIL（`P列名指し救済適用_ is not a function`）

- [ ] **Step 3b: 救済ループを純粋関数へ抽出し計画対象にも適用（現360-371行）**

変更前:
```js
  // 【コード不一致の名指し救済】箱コードが説明文(핫 토픽…等)でどの注文とも一致せず、
  // P列に受注番号が手書きされている行は、その番号へのdirect名指しとして扱い、名指しを消さない。
  // (コードが注文と一致する行の手動Pは従来通り計画が正=上書きされ得る)
  const 全注文キー=new Set();
  lines.forEach(l=>{ (l.keys instanceof Set?Array.from(l.keys):l.keys||[]).forEach(k=>全注文キー.add(k)); });
  rows.forEach(r=>{
    if(!r.対象 || r.directBan) return;
    const ban=P手動名指し解析_(r.pOriginal);
    if(!ban) return;
    if(全注文キー.has(r.code)) return; // コードが一致する=通常ルートで扱う
    r.directBan=ban; r.手動名指し=true;
  });
```
変更後（関数本体は `P列処理対象EMS_` 群の近くに置く）:
```js
  const 全注文キー=new Set();
  lines.forEach(l=>{ (l.keys instanceof Set?Array.from(l.keys):l.keys||[]).forEach(k=>全注文キー.add(k)); });
  P列名指し救済適用_(rows, 全注文キー);
```
新関数:
```js
// 【コード不一致の名指し救済】箱コードが説明文(핫 토픽…等)でどの注文とも一致せず、
// P列に受注番号が手書きされている行は、その番号へのdirect名指しとして扱い、名指しを消さない。
// (コードが注文と一致する行の手動Pは従来通り計画が正=上書きされ得る)。未着(計画対象)行にも適用
function P列名指し救済適用_(rows, 全注文キー){
  (rows||[]).forEach(r=>{
    if(!r || (!r.対象 && !r.計画) || r.directBan) return;
    const ban=P手動名指し解析_(r.pOriginal);
    if(!ban) return;
    if(全注文キー.has(r.code)) return; // コードが一致する=通常ルートで扱う
    r.directBan=ban; r.手動名指し=true;
  });
}
```

- [ ] **Step 3c: GREENを確認**

Run: `node tests/project24_pre_arrival_plan.test.js`
Expected: `ALL PASS`（8件）

- [ ] **Step 4: 純計算の入力を2パス順へ・ブロック行は両対象（現382-383行）**

変更前:
```js
  const calculated=P列計画_純計算_(rows.filter(r=>r.対象&&!r.全件再計算ブロック),lines,fixedBySupply,ledgerSummary.usageBySupply);
  rows.filter(r=>r.対象&&r.全件再計算ブロック).forEach(r=>calculated.rows.push(Object.assign({},r,{entries:[],left:r.qty,nextP:''})));
```
変更後:
```js
  const calculated=P列計画_純計算_(P列計画行順_(rows),lines,fixedBySupply,ledgerSummary.usageBySupply);
  rows.filter(r=>(r.対象||r.計画)&&r.全件再計算ブロック).forEach(r=>calculated.rows.push(Object.assign({},r,{entries:[],left:r.qty,nextP:''})));
```

- [ ] **Step 5: 背景を全計算行の自己修復方式へ（現390-391行）**

変更前:
```js
  const backgrounds=calculated.rows.filter(r=>writes.indexOf(r.i)>=0 && r.nextP && (/[:：,、]/.test(r.nextP)||r.left>0))
    .map(r=>({a1:P列セルA1_(hr+1+r.i,colP),color:'#fff2cc'}));
```
変更後（ファイル先頭付近に `const P列予定色='#c9daf8'; // 未着行の先行記入(予定)の印` を追加した上で）:
```js
  // 背景は毎回全計算行に引き直す(自己修復)。未着→到着済の昇格でテキスト不変でも青が通常色へ戻るように
  const backgrounds=calculated.rows.map(r=>{
    const text=writes.indexOf(r.i)>=0? r.nextP : String(r.pOriginal||'');
    let color=null;
    if(r.計画) color=text? P列予定色 : null;
    else if(text && (/[:：,、]/.test(text) || r.left>0)) color='#fff2cc';
    return {a1:P列セルA1_(hr+1+r.i,colP),color};
  });
```

- [ ] **Step 6: summaryに予定件数・returnにlines（現392-402行）**

変更前:
```js
  const summary={
    記入:writes.filter(i=>pColumn[i][0]).length,
    分割:calculated.rows.filter(r=>r.nextP && /[:：,、]/.test(r.nextP)).length,
    在庫:calculated.rows.filter(r=>r.qty>0 && !r.nextP).length,
    既存:calculated.rows.filter(r=>String(r.pOriginal||'').trim()).length,
    過剰除外:0,解析警告:0
  };
  return {
    error:'',sheet:ems,startRow:hr+1,colP,rowCount:n,values:pColumn,backgrounds,writes,
    rows:calculated.rows,到着実績,到着便,到着実績取得済,summary,forceWrite:!!options.clearCurrentP
  };
```
変更後:
```js
  const summary={
    記入:writes.filter(i=>pColumn[i][0]).length,
    分割:calculated.rows.filter(r=>r.nextP && /[:：,、]/.test(r.nextP)).length,
    在庫:calculated.rows.filter(r=>r.qty>0 && !r.nextP).length,
    既存:calculated.rows.filter(r=>String(r.pOriginal||'').trim()).length,
    予定:calculated.rows.filter(r=>r.計画 && r.nextP).length,
    過剰除外:0,解析警告:0
  };
  return {
    error:'',sheet:ems,startRow:hr+1,colP,rowCount:n,values:pColumn,backgrounds,writes,
    rows:calculated.rows,lines,到着実績,到着便,到着実績取得済,summary,forceWrite:!!options.clearCurrentP
  };
```

- [ ] **Step 7: 回帰テスト（vm読込の構文チェック兼用）**

Run: 既存8ファイル＋新規1ファイルを全部 `node tests/<f>` 実行
Expected: 全PASS（この時点で挙動が変わるのは「未着行にnextPが計算される」ことだけで、書込side effectはまだ従来のwritesゲート内）

- [ ] **Step 8: コミット**

```bash
git add "GASアップスクリプト/Project_24/P列自動記入.js"
git commit -m "feat(Project_24): P列計画を未着行(計画対象)へ拡張(2パス消費・救済継承・予定色の自己修復背景)"
```

---

### Task 3: 反映funnel＋受注明細「入荷予定」書込

**Files:**
- Modify: `Project_24/P列自動記入.js`（発注共有P列計画を反映_ 現254-259行、＋新関数1つ）
- Modify: `Project_24/引当.js`（列マップ_ 現393行付近に1行）

**Interfaces:**
- Consumes: Task 1 の `入荷予定マップ_`、Task 2 の `plan.lines` / `plan.rows[].計画` / 全行backgrounds
- Produces: `受注明細入荷予定を更新_(plan):void` — funnel経由で毎回実行。列マップ_に `入荷予定` フィールド追加

- [ ] **Step 1: 列マップ_に入荷予定を追加**

`Project_24/引当.js` 現393行 `EMS: find('EMS番号'),` の直後に:
```js
    入荷予定: find('入荷予定'),
```

- [ ] **Step 2: 受注明細入荷予定を更新_ を実装**

`Project_24/P列自動記入.js` の `発注共有P列計画を反映_` の直前に追加:

```js
// 受注明細「入荷予定」列を毎回全書き直しする。列が無ければ右端に自動作成
// (取込_実行_はヘッダー行を保持するため、一度作れば以後の取込でも列は生き残る。値は②のたび再生成)
function 受注明細入荷予定を更新_(plan){
  if(!plan || plan.error || !plan.rows) return;
  const recv=SpreadsheetApp.getActive().getSheetByName(HIKIATE_CFG.受注);
  if(!recv) return;
  const M=列マップ_(recv);
  let col=M.入荷予定;
  if(col<0){
    col=recv.getLastColumn();
    if(recv.getMaxColumns()<col+1) recv.insertColumnsAfter(recv.getMaxColumns(),col+1-recv.getMaxColumns());
    recv.getRange(M.hr,col+1).setValue('入荷予定');
  }
  const last=recv.getLastRow(); if(last<=M.hr) return;
  const n=last-M.hr;
  const map=入荷予定マップ_(plan.rows, plan.lines);
  const values=[];
  for(let i=0;i<n;i++) values.push([map[M.hr+1+i]||'']);
  recv.getRange(M.hr+1,col+1,n,1).setValues(values);
}
```

- [ ] **Step 3: funnelを「背景と入荷予定は毎回」へ変更（現254-259行）**

変更前:
```js
function 発注共有P列計画を反映_(plan){
  if(plan.error || !plan.writes || (!plan.writes.length && !plan.forceWrite)) return plan.summary||plan;
  plan.sheet.getRange(plan.startRow,plan.colP,plan.rowCount||plan.values.length,1).setValues(plan.values);
  P列背景を反映_(plan.sheet, plan.backgrounds);
  return Object.assign({},plan.summary,{到着実績:plan.到着実績,到着便:plan.到着便,到着実績取得済:plan.到着実績取得済});
}
```
変更後:
```js
function 発注共有P列計画を反映_(plan){
  if(plan.error || !plan.writes) return plan.summary||plan;
  if(plan.writes.length || plan.forceWrite)
    plan.sheet.getRange(plan.startRow,plan.colP,plan.rowCount||plan.values.length,1).setValues(plan.values);
  P列背景を反映_(plan.sheet, plan.backgrounds); // 背景は毎回自己修復(昇格の青→通常色を含む)
  受注明細入荷予定を更新_(plan); // 到着日変化など、P列テキスト不変でも予定表示は変わり得る
  return Object.assign({},plan.summary,{到着実績:plan.到着実績,到着便:plan.到着便,到着実績取得済:plan.到着実績取得済});
}
```

- [ ] **Step 4: 回帰テスト**

Run: 全9テストファイル
Expected: 全PASS（funnelはGAS結線でvmテスト対象外、構文エラー検出のみ。呼び元 引当.js:1627 / P列自動記入.js:416 は無変更で全経路が入荷予定更新を通る）

- [ ] **Step 5: コミット**

```bash
git add "GASアップスクリプト/Project_24/P列自動記入.js" "GASアップスクリプト/Project_24/引当.js"
git commit -m "feat(Project_24): 受注明細「入荷予定」列を自動生成(未着便の到着予定を注文視点で表示)"
```

---

### Task 4: ♻️P列を書き直す の未着対応

**Files:**
- Modify: `Project_24/P列自動記入.js`（P列を書き直す 現12-30行のダイアログ/トースト、P列書き直し実行_ 現33-68行）

**Interfaces:**
- Consumes: Task 1 の `P列計画対象EMS_`、Task 3 の `受注明細入荷予定を更新_`、既存 `実EMS番号_`
- Produces: 書き直しが未着行（実EMS番号あり）もクリア→予定を入れ直す。戻り値に `予定` 件数

- [ ] **Step 1: クリア範囲の拡張（現43-55行）**

変更前:
```js
  const cSt=f('ステータス列','ステータス'), cP=f('注文番号');
  if(cP<0) return {error:'EMSリストの'+hr+'行目に「注文番号」見出しがありません'};
  const n=last-hr;
  const st=cSt>=0? sh.getRange(hr+1,cSt+1,n,1).getDisplayValues() : null;
  const pv=sh.getRange(hr+1,cP+1,n,1).getDisplayValues();
  const clearedValues=pv.map(row=>[String(!row||row[0]==null?'':row[0])]);
  const clearedRows=[];
  let cleared=0;
  for(let i=0;i<n;i++){
    if(st && String(st[i][0]||'').trim()!=='到着済') continue; // 到着済だけ(ステータス列が無ければ全行)
    if(String(clearedValues[i][0]||'').trim()==='') continue;
    clearedValues[i][0]=''; cleared++; clearedRows.push(i);
  }
```
変更後:
```js
  const cSt=f('ステータス列','ステータス'), cP=f('注文番号'), cE=f('EMS番号');
  if(cP<0) return {error:'EMSリストの'+hr+'行目に「注文番号」見出しがありません'};
  const n=last-hr;
  const st=cSt>=0? sh.getRange(hr+1,cSt+1,n,1).getDisplayValues() : null;
  const emsNos=cE>=0? sh.getRange(hr+1,cE+1,n,1).getDisplayValues() : null;
  const pv=sh.getRange(hr+1,cP+1,n,1).getDisplayValues();
  const clearedValues=pv.map(row=>[String(!row||row[0]==null?'':row[0])]);
  const clearedRows=[];
  let cleared=0;
  for(let i=0;i<n;i++){
    const status=st? String(st[i][0]||'').trim() : '';
    const 確定=st? status==='到着済' : true; // ステータス列が無ければ従来どおり全行
    const 計画=!!(st && emsNos && P列計画対象EMS_(status) && 実EMS番号_(emsNos[i][0])); // 未着は実EMS番号行だけ(在庫反映済みは不変)
    if(!確定 && !計画) continue;
    if(String(clearedValues[i][0]||'').trim()==='') continue;
    clearedValues[i][0]=''; cleared++; clearedRows.push(i);
  }
```

- [ ] **Step 2: 実行部の戻り値と入荷予定更新（現63-67行）**

変更前:
```js
  plan.sheet.getRange(plan.startRow,plan.colP,plan.rowCount,1).setValues(plan.values);
  const backgrounds=clearedRows.map(i=>({a1:P列セルA1_(plan.startRow+i,plan.colP),color:null}))
    .concat(plan.backgrounds||[]);
  P列背景を反映_(plan.sheet, backgrounds);
  return {クリア:cleared, 記入:plan.summary.記入, 分割:plan.summary.分割, 在庫:plan.summary.在庫};
```
変更後:
```js
  plan.sheet.getRange(plan.startRow,plan.colP,plan.rowCount,1).setValues(plan.values);
  const backgrounds=clearedRows.map(i=>({a1:P列セルA1_(plan.startRow+i,plan.colP),color:null}))
    .concat(plan.backgrounds||[]);
  P列背景を反映_(plan.sheet, backgrounds);
  受注明細入荷予定を更新_(plan);
  return {クリア:cleared, 記入:plan.summary.記入, 分割:plan.summary.分割, 在庫:plan.summary.在庫, 予定:plan.summary.予定};
```

- [ ] **Step 3: ダイアログとトーストの文言更新（現12-30行）**

ダイアログ本文の変更前:
```js
    'EMSリストの「到着済」行のP列(注文番号)を全部消して、今のロジックで書き直します。\n\n'+
    '・バグ時代の古い割当(残骸)が一掃されます\n'+
    '・手で書いた名指しも消えます(コード末尾の（受注番号）タグは自動で再現)\n'+
    '・在庫反映済みなど過去の行は触りません\n\n実行しますか？',
```
変更後:
```js
    'EMSリストの「到着済」行と未着行(実EMS番号あり)のP列(注文番号)を全部消して、今のロジックで書き直します。\n\n'+
    '・バグ時代の古い割当(残骸)が一掃されます\n'+
    '・手で書いた名指しも消えます(コード末尾の（受注番号）タグとコード不一致の名指しは自動で再現)\n'+
    '・未着行には予定(薄青)が入り直します\n'+
    '・在庫反映済みなど過去の行は触りません\n\n実行しますか？',
```
トーストの変更前:
```js
    'P列書き直し: クリア'+r.クリア+'行 → 記入'+r.記入+'行'+(r.分割?'（分割'+r.分割+'行）':'')+
    ' / 在庫扱い'+r.在庫+'行。仕上げに ②引き当て実行 を回してください','♻️P列',8);
```
変更後:
```js
    'P列書き直し: クリア'+r.クリア+'行 → 記入'+r.記入+'行'+(r.分割?'（分割'+r.分割+'行）':'')+
    (r.予定?'（うち予定'+r.予定+'行）':'')+
    ' / 在庫扱い'+r.在庫+'行。仕上げに ②引き当て実行 を回してください','♻️P列',8);
```

- [ ] **Step 4: 回帰テスト＋コミット**

Run: 全9テストファイル → 全PASS
```bash
git add "GASアップスクリプト/Project_24/P列自動記入.js"
git commit -m "feat(Project_24): ♻️P列書き直しを未着行対応(予定クリア&入れ直し・件数表示)"
```

---

### Task 5: 本番反映（ゲート付き）と実地確認

**Files:**
- なし（運用手順のみ）

- [ ] **Step 1: ゲート確認** — ユーザーの「✅ 全件再計算を反映」が完了済みであること（タスク#1）。未完了なら**ここで停止**しコミットのみで待機
- [ ] **Step 2: 同期** — `git pull origin codex/full-allocation-rebuild-v3` → `powershell -ExecutionPolicy Bypass -File tools\gas_pull_sync.ps1 Project_24`（オンライン一致 or sync取り込みを確認）
- [ ] **Step 3: push** — `powershell -ExecutionPolicy Bypass -File tools\gas_safe_push.ps1 Project_24`（「✅ 安全push完了」を確認。コンフリクト対話が出たら中止してユーザーへ）
- [ ] **Step 4: 独立検証** — scratchフォルダへ `.clasp.json` をコピーして `clasp pull` し、Project_24全ファイルのSHA256がworktreeと一致することを確認
- [ ] **Step 5: origin push** — `git push origin codex/full-allocation-rebuild-v3`
- [ ] **Step 6: 実地確認（ユーザーと）** — ②引き当て実行後: (a) EMSリスト未着行に薄青の予定が入る (b) 受注明細右端に「入荷予定」列が生成され `7/xx(…9766)` 形式が入る (c) JMEE167の未着20個の行に現在の取り寄せ注文が予定表示される (d) 到着済行のP列は従来と同一内容

---

## Self-Review結果（作成時実施済み）

- スペック§1-6→Task対応: §1-2=Task1-2、§3=Task2(色)+Task4(書き直し)、§4=Task3、§5=不変(コード変更なしで担保)、§6=Task1＋Task2 Step3a(救済)＋各Taskの回帰
- 型整合: `P列計画行順_(rows)` はTask1定義→Task2 Step4で使用。`plan.lines`/`計画`フラグはTask2で生成→Task3 `入荷予定マップ_(plan.rows, plan.lines)` で消費。`summary.予定` はTask2で生成→Task4トーストで表示。`P列名指し救済適用_` はTask2 Step3bで定義・同Taskで使用
- 到着済行の挙動不変の担保: Task1「未着FIFO」テストの単独実行比較＋救済テストの到着済ケース＋Task5 Step6(d)
