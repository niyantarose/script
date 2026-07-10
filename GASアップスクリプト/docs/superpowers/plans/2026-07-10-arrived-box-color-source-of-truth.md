# Arrived Box Color Source of Truth Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 「到着済」P列確定割当を黄色、「在庫反映済み」履歴をラベンダーとして表示する。

**Architecture:** `P列確定マップ_` に到着日とEMS番号を追加し、`引当実行_本体_` の確定引当が古い入荷日を持つ行も現在便として処理する。P列未確定の手入力値は変更しない。

**Tech Stack:** Google Apps Script JavaScript、Node.js `vm`/`assert` 回帰テスト、clasp安全同期スクリプト

## Global Constraints

- Project_24のみ変更する。
- 素の`clasp push`は禁止し、`tools/gas_safe_push.ps1 Project_24`を使う。
- P列未確定の手入力日付は上書きしない。
- 在庫数量を超える引当を作らない。

---

### Task 1: P列確定メタデータと色判定

**Files:**
- Modify: `Project_24/P列確定.js`
- Modify: `Project_24/引当.js`
- Create: `tests/project24_arrived_box_color.test.js`

**Interfaces:**
- Consumes: `P列確定マップ_(): Record<string, Array<{key:string, qty:number, arrival:string, ems:string}>>`
- Produces: `l.今回P`, `l.今回P入荷日値`, `l.今回PEMS` を持つ現在便の引当行

- [ ] **Step 1: Write the failing test**

  P列マップが到着日・EMS番号を返すことと、古い入荷日でも`今回P`なら現在便扱いになることをassertする。

- [ ] **Step 2: Run test to verify it fails**

  Run: `node tests/project24_arrived_box_color.test.js`
  Expected: arrival/ems欠落または現在便判定falseでFAIL。

- [ ] **Step 3: Write minimal implementation**

  `P列確定マップ_`の各要素へ`arrival`と`ems`を追加し、確定引当成功時に現在便フラグと訂正値を保持する。既存日付が現在便と不一致の確定行だけ再引当対象に含める。

- [ ] **Step 4: Run test and syntax checks**

  Run: `node tests/project24_arrived_box_color.test.js`
  Run: `node --check Project_24/P列確定.js`
  Run: `node --check Project_24/引当.js`
  Expected: 全てexit 0。

- [ ] **Step 5: Deploy safely**

  Run: `powershell -ExecutionPolicy Bypass -File tools/gas_safe_push.ps1 Project_24`
  Expected: オンライン競合なしでProject_24へ反映。

