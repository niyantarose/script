# 全件検算と安全な再構築(改訂版v2) 実装計画

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 旧「全件再構築」仕様の採用部品(バックアップ・検算・重複排除)だけを既存フローへ移植する。

**Architecture:** 新しい引当エンジンは作らない。読み取り専用の突き合わせレポート(全件検算.js)+全リセットの事前バックアップ+台帳読み取り時の重複排除関数。純粋ロジックは node+vm テスト(既存 tests/project24_arrived_box_color.test.js と同方式)。

**Tech Stack:** Google Apps Script (Project_24)、Node.js 標準 assert/vm テスト。

## Global Constraints

- 仕様: `docs/superpowers/specs/2026-07-10-full-allocation-rebuild-v2-design.md`
- レポートは読み取り専用(受注明細・EMSリスト・台帳・履歴へ一切書き込まない)
- ②・🔎・棚卸の既存挙動を変えない(棚卸.jsのリファクタは挙動不変)
- 編集前に `tools/gas_pull_sync.ps1 Project_24`、反映は `tools/gas_safe_push.ps1 Project_24`(素のclasp push禁止)
- コミットは `feat(hikiate):`/`docs(hikiate):` スタイル・日本語1行+本文

---

### Task 1: 旧仕様書へ不採用注記

**Files:**
- Modify: `docs/superpowers/specs/2026-07-10-full-allocation-rebuild-design.md:1`

- [x] **Step 1: 冒頭に注記を追加**

タイトル直後に:

```markdown
> **不採用(2026-07-10)**: 本設計は「②と並ぶ第2の引当エンジン」を作るため採用しない。
> 採用価値のある部品(事前バックアップ・Yahoo検算・出荷済み重複排除)は
> [改訂版v2](2026-07-10-full-allocation-rebuild-v2-design.md) へ移植した。
```

- [x] **Step 2: 仕様v2・本計画とまとめて docs コミット**

```bash
git add docs/superpowers/
git commit -m "docs(hikiate): 全件再計算の設計を改訂(v2=既存フローへの部品移植)"
```

### Task 2: 棚卸.js の Yahoo CSV 読み込みを関数へ抽出(挙動不変)

**Files:**
- Modify: `Project_24/棚卸.js`(生成本体のセクション1を置換)
- Test: `tests/project24_zenken_kensan.test.js`(新規。Task 2〜4 で同じファイルに追記)

**Interfaces:**
- Produces: `CSV行分解_(line)->string[]` / `YahooCSV集計_(text)->{a在庫,商品名,subなし,解析スキップ}|{error}` / `Yahoo在庫を読む_()->{a在庫,商品名,subなし,解析スキップ,fileName}|{error}`

- [x] **Step 1: 失敗するテストを書く**(YahooCSV集計_: a加算/b無視/引用符/壊れ行スキップ/sub無し検出)
- [x] **Step 2: `node tests/project24_zenken_kensan.test.js` が FAIL することを確認**
- [x] **Step 3: 棚卸.js に CSV行分解_/YahooCSV集計_/Yahoo在庫を読む_ を抽出し、生成本体を置き換え**
- [x] **Step 4: 新テストと既存テスト(project24_arrived_box_color.test.js)が PASS することを確認**

### Task 3: 消込台帳.js に出荷済み重複排除を追加

**Files:**
- Modify: `Project_24/消込台帳.js`(末尾に追加)

**Interfaces:**
- Produces: `受注基底コード_(sku,code)->string`(SKU優先→normCode_→末尾A/B落とし) / `出荷済み重複排除_(rows)->rows`(受注番号+基底コードで数量最大の1件)

- [x] **Step 1: 失敗するテストを追記**(表記ゆれ統合/数量最大/受注番号・コード違いは残す/ban無しは素通し)
- [x] **Step 2: FAIL 確認 → 実装 → PASS 確認**

### Task 4: 全件検算.js の純粋集計

**Files:**
- Create: `Project_24/全件検算.js`

**Interfaces:**
- Consumes: `normCode_/ymd_/区分_`(引当.js)、`受注基底コード_/出荷済み重複排除_`(消込台帳.js)
- Produces: `全件検算_集計_(src)->{rows,counts}`。src/戻りの形は仕様書「機能2」のとおり。判定は6段の先勝ち。

- [x] **Step 1: 失敗するテストを追記**(仕様のテストケース3〜9: 超過消費/入荷日ズレ/供給不足/箱残>Yahoo/EMS外在庫/OK/除外/過去便分類/並び順)
- [x] **Step 2: FAIL 確認 → 実装 → PASS 確認**

### Task 5: レポート描画とメニュー(GAS側)

**Files:**
- Modify: `Project_24/全件検算.js`(`全件検算レポート()`/`全件検算レポート本体_()` を追加)
- Modify: `Project_24/引当.js:73` 付近(メニュー1項目)

- [x] **Step 1: 本体実装**(EMSリスト/受注明細/台帳/Yahooを読み → 全件検算_集計_ → 「全件検算」シートへ描画。直列_でロック。Yahoo読めない時は判定4/5なしで注記)
- [x] **Step 2: メニュー追加** `🧮 全件検算レポート(EMS×台帳×受注×Yahooの突き合わせ)`
- [x] **Step 3: 全テスト PASS 確認**(vm ロードで構文/参照エラーも検出)

### Task 6: 全リセットの事前バックアップ

**Files:**
- Modify: `Project_24/入荷日チェック.js`(`引当データの全リセット本体_`)

- [x] **Step 1: 「リセット」入力確認の後・ロック取得の前に Drive コピーを作成**(`引当ファイル_リセット前_yyyyMMdd_HHmmss`、元と同じフォルダ、失敗時は何も消さず中止)
- [x] **Step 2: 完了ダイアログにバックアップ名を表示**
- [x] **Step 3: 全テスト PASS 確認**

### Task 7: コミットと反映

- [x] **Step 1: `node tests/project24_arrived_box_color.test.js` と `node tests/project24_zenken_kensan.test.js` の PASS を確認**
- [x] **Step 2: feat コミット**

```bash
git add Project_24/ tests/
git commit -m "feat(hikiate): 全件検算レポート(読み取り専用)と全リセットの事前バックアップを追加"
```

- [x] **Step 3: `tools/gas_safe_push.ps1 Project_24` で反映**(sync系コミットが増えたら取り込む)
