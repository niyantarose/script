# 取り置き登録・棚確認表示改善 実装計画

> **For Codex:** REQUIRED SUB-SKILL: Use `superpowers:subagent-driven-development` to implement this plan task by task with specification review and code-quality review.

**Goal:** 「取り置き登録」で未来発売の予約だけを商品名・選択肢から除外し、過去予約を表示したうえで、注文単位の太線と棚確認ステータス別の即時行色を追加する。

**Architecture:** 既存の候補生成・DocumentProperties記憶・一括シート書き込みを維持する。候補判定は受注ステータスだけに限定し、注文グループを壊さない純粋ソートを追加する。表示は既存の注文境界ヘルパーと、固定5本の条件付き書式を一括適用する専用ヘルパーで構成する。

**Tech Stack:** Google Apps Script / Google Sheets API、Node.js `vm` ベースの既存回帰テスト、Git worktree

---

### Task 1: 未来予約の除外ルールと手動記憶を新仕様へ揃える

**Files:**
- Modify: `tests/project24_reservation_workflow.test.js:130-150`
- Modify: `tests/project24_reservation_workflow.test.js:1164-1205`
- Modify: `Project_24/取り置き台帳.js:9-134`

**Step 1: 未来予約だけを判定する回帰テストを作る**

`取り置き_予約判定_` が次を満たすテストにする。

```javascript
test('予約判定は未来の発売予定だけを自動予約にする', () => {
  const today=new Date(2026,6,17);
  assert.strictEqual(context.取り置き_予約判定_('予約9月韓国発売予定','',today),true);
  assert.strictEqual(context.取り置き_予約判定_('予約7月末韓国発売予定','',today),true);
  assert.strictEqual(context.取り置き_予約判定_('予約5月韓国発売予定','',today),false);
  assert.strictEqual(context.取り置き_予約判定_('','予約早期完売',today),false);
});
```

初期候補では未来予約を内部的に `予約` とし、旧入荷日スタンプまたは `部分包装` がある行は空欄に戻して現物形跡を優先することも検証する。自動 `予約` 行は後段の判断済み記憶処理で候補から非表示になる。

**Step 2: 自動除外が正確に2ステータスだけである失敗テストを作る**

既存の登録絞り込みテストを次へ更新する。

- `予約中` は除外
- `■出荷GO` は除外
- `出荷待/取寄せ` は旧入荷日が空でも表示
- `予約受付終了` は未来日を含まず「予約中」でもないため表示
- 数量入力済みでも `予約中` / `出荷GO` は除外

期待IDへ `C` と新規の `J` を含める。この時点では本体が `/予約/` のため `予約受付終了` が落ち、テストはREDになることを確認する。

**Step 3: 手動「予約」の保存・非表示テストを追加する**

`取り置き_棚確認記憶を適用_` のテストに `棚確認:'予約'` の候補を追加し、数量が空なら非表示・storeへ保存されることを検証する。数量ありなら表示される既存矛盾ガードも維持する。

**Step 4: REDを確認する**

Run:

```powershell
node tests\project24_reservation_workflow.test.js
```

Expected: 未来予約判定・初期候補への反映に関するテストが失敗する。

**Step 5: 最小実装で未来予約判定と正確なステータス除外を実装する**

`取り置き_登録絞り込み_` を次の考え方へ変更する。

```javascript
if(st.indexOf('予約中')>=0) return false;
if(st.indexOf('出荷GO')>=0) return false;
return true;
```

`取り置き_予約判定_` を復元し、次の入力だけを未来予約と判定する。

- 年月日が読め、日付がtodayより後
- 年なしの月が現在月より後
- 当月の `月末` 表記で月末日がtodayより後

過去日、日付不明、予約表記なしはfalseにする。候補収集時に商品名・選択肢を判定して予約フラグへ渡し、初期候補で自動 `予約` を設定する。ただし旧入荷日または `部分包装` がある行は空欄にする。`予約` を判断済み記憶の対象へ含める変更は維持する。

**Step 6: GREENを確認する**

Run:

```powershell
node tests\project24_reservation_workflow.test.js
```

Expected: 全件PASS。

**Step 7: Task 1をコミットする**

```powershell
git add -- Project_24/取り置き台帳.js tests/project24_reservation_workflow.test.js
git commit -m "fix(hikiate): 予約除外を受注ステータスだけに限定"
```

---

### Task 2: 注文グループを保った並び順と太い罫線を追加する

**Files:**
- Modify: `tests/project24_reservation_workflow.test.js`
- Modify: `Project_24/取り置き台帳.js:102-140`
- Modify: `Project_24/取り置き台帳.js:356-373`
- Reuse: `Project_24/引当.js:713-724,2360-2369`

**Step 1: 注文グループを壊さない並び順の失敗テストを追加する**

同じ受注番号に `要棚確認` と通常行が混在しても連続したままになり、要確認を1件以上含む注文グループが上へ来ることを検証する。

```javascript
test('要棚確認を優先しても同じ注文の商品行を分断しない', () => {
  const rows=context.取り置き_注文単位で並べる_([
    {受注番号:'200',商品コード:'B1',判定:''},
    {受注番号:'100',商品コード:'A1',判定:'要棚確認'},
    {受注番号:'100',商品コード:'A2',判定:''},
    {受注番号:'300',商品コード:'C1',判定:''}
  ]);
  assert.strictEqual(JSON.stringify(rows.map(r=>r.受注番号)),JSON.stringify(['100','100','200','300']));
});
```

**Step 2: 境界範囲の失敗テストを追加する**

専用ラッパー `取り置き_注文境界A1_` が、データ開始行2・全14列で同じ注文の最終行だけを返すことを検証する。

```javascript
assert.strictEqual(JSON.stringify(Array.from(context.取り置き_注文境界A1_([
  {受注番号:'100'},{受注番号:'100'},{受注番号:'200'}
]))),JSON.stringify(['A3:N3','A4:N4']));
```

**Step 3: REDを確認する**

Run:

```powershell
node tests\project24_reservation_workflow.test.js
```

Expected: 新しい2ヘルパーが未定義で失敗する。

**Step 4: 注文単位ソートを実装する**

`取り置き_注文単位で並べる_` は次を満たす純粋関数にする。

- 受注番号ごとに最初の出現順を保持してグループ化
- グループ内の行順は保持
- `要棚確認` を1件以上含むグループを先にする
- 同じ優先度のグループ順は安定

現在の行単位 `.sort(...)` をこのヘルパー呼び出しへ置き換える。

**Step 5: 境界ラッパーと実シート罫線を追加する**

`取り置き_注文境界A1_` は既存 `注文境界A1_` へ `受注番号 -> ban` を渡す薄いラッパーにする。シート作成後は既存 `注文罫線_(sh2, 2, 1)` を呼び、B列の受注番号で薄い格子と黒い `SOLID_THICK` の下線を一括設定する。

**Step 6: GREENを確認する**

Run:

```powershell
node tests\project24_reservation_workflow.test.js
```

Expected: 全件PASS。

**Step 7: Task 2をコミットする**

```powershell
git add -- Project_24/取り置き台帳.js tests/project24_reservation_workflow.test.js
git commit -m "feat(hikiate): 取り置き登録を注文単位の太線で区切る"
```

---

### Task 3: 棚確認ステータス別の条件付き書式を追加する

**Files:**
- Modify: `tests/project24_reservation_workflow.test.js`
- Modify: `Project_24/取り置き台帳.js:9-15`
- Modify: `Project_24/取り置き台帳.js:361-381`

**Step 1: 条件付き書式定義の失敗テストを追加する**

純粋関数 `取り置き_棚確認書式定義_` が5ステータスを決められた順・色で返すことを検証する。

```javascript
const defs=context.取り置き_棚確認書式定義_(11,2);
assert.strictEqual(JSON.stringify(defs.map(d=>[d.値,d.背景])),JSON.stringify([
  ['発送待ち','#cfe2f3'],
  ['部分在庫','#d9ead3'],
  ['出荷済み','#f4cccc'],
  ['未着','#d9d9d9'],
  ['予約','#d9d2e9']
]));
assert.strictEqual(defs[2].条件,'=$K2="出荷済み"');
assert.strictEqual(defs[2].文字色,'#990000');
assert.strictEqual(defs[2].太字,true);
```

**Step 2: シート適用の失敗テストを追加する**

`SpreadsheetApp.newConditionalFormatRule()` とシートをモックし、`取り置き_棚確認書式を設定_(sheet, rowCount)` が次を満たすことを確認する。

- 全データ列・対象行だけをrangesへ設定
- 5ルールを1回の `setConditionalFormatRules` で保存
- `出荷済み`だけ濃い赤文字・太字
- 0行では空配列を設定し、前回ルールを消す

**Step 3: REDを確認する**

Run:

```powershell
node tests\project24_reservation_workflow.test.js
```

Expected: 書式定義・適用ヘルパー未定義で失敗する。

**Step 4: 書式定義と適用を最小実装する**

見出し配列から棚確認列番号と全幅を求める。条件式は列番号をA1列記号へ変換して動的に作る。

```javascript
const defs=取り置き_棚確認書式定義_(棚確認列,2);
const target=sh.getRange(2,1,rowCount,TORIOKI_CFG.初期HDR.length);
const rules=defs.map(def=>{
  let b=SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(def.条件)
    .setBackground(def.背景)
    .setRanges([target]);
  if(def.文字色) b=b.setFontColor(def.文字色);
  if(def.太字) b=b.setBold(true);
  return b.build();
});
sh.setConditionalFormatRules(rules);
```

**Step 5: 作成処理へ統合する**

- 既存の黄色い `要棚確認` 基本背景は維持
- データ行の有無にかかわらず、作成処理の最後に条件付き書式ヘルパーを呼ぶ
- ステータスが入ると条件付き書式が黄色より優先表示される
- `onEdit` は変更しない

**Step 6: GREENを確認する**

Run:

```powershell
node tests\project24_reservation_workflow.test.js
```

Expected: 全件PASS。

**Step 7: Task 3をコミットする**

```powershell
git add -- Project_24/取り置き台帳.js tests/project24_reservation_workflow.test.js
git commit -m "feat(hikiate): 棚確認ステータスを行色へ即時反映"
```

---

### Task 4: Project_24全体回帰・レビュー・引き渡し

**Files:**
- Verify: `Project_24/取り置き台帳.js`
- Verify: `tests/project24_*.test.js`
- Review: commits created by Tasks 1-3

**Step 1: 構文確認を実行する**

Run:

```powershell
node --check Project_24\取り置き台帳.js
```

Expected: exit code 0。

**Step 2: Project_24全6スイートを実行する**

Run each:

```powershell
node tests\project24_arrived_box_color.test.js
node tests\project24_daniel_amari.test.js
node tests\project24_reservation_domain.test.js
node tests\project24_reservation_workflow.test.js
node tests\project24_torioki.test.js
node tests\project24_zenken_kensan.test.js
```

Expected: 6スイートすべてPASS。

**Step 3: 差分検査を実行する**

Run:

```powershell
git diff --check HEAD~3..HEAD
git status --short
```

Expected: whitespace errorなし。対象外の新規変更なし。

**Step 4: 仕様レビューを行う**

`superpowers:requesting-code-review` を使い、設計書の各要件について次を確認する。

- 自動除外が `予約中` / `出荷GO` のみ
- 手動 `予約` が記憶・非表示対象
- 注文グループが連続し太線が末尾に付く
- 選択した商品行だけ色が変わる
- 出荷済みが赤で最も目立つ
- 引当計算ロジックに変更がない

**Step 5: コード品質レビューを行う**

API呼び出し数、既存書式との干渉、0件時の古いルール残存、列位置のハードコード、テストの脆さを確認する。指摘があれば同じTask内で修正・再テスト・追加コミットする。

**Step 6: 配備前チェックポイントで停止する**

ローカルworktreeの実装・テスト・コミット完了を報告する。実シートへ反映する `clasp push` は、差分対象がProject_24だけであることを再確認してから別手順で行う。
