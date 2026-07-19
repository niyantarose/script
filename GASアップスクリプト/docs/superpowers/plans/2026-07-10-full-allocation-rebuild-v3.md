# Full Allocation Rebuild v3 Implementation Plan

> 2026-07-19 update: Execute `2026-07-19-full-allocation-rebuild-ledger.md` instead.

> **For Codex:** Execute this plan task-by-task with test-driven development. Do not apply the rebuilt data to the live sheets until the preview has been generated and reviewed.

**Goal:** Replace the synthetic shelf-stock workflow with a source-of-truth rebuild that derives P-column assignments, allocation history, and Korean EMS arrival fields from real EMS rows, GoQ shipped records, current orders, and Yahoo `a` free inventory.

**Architecture:** Add one `全件再計算.js` module containing a pure quantity engine plus GAS adapters. The pure engine consumes normalized supplies, shipped demands, current orders, and Yahoo free-stock counts and returns allocations, blocked SKUs, summaries, and write plans. Preview stores both signatures and the serialized write plan. Apply revalidates signatures, creates two Drive backups, excludes legacy synthetic rows, replaces derived data, and runs the existing output renderer with P auto-writing disabled and blocked SKUs removed from stock.

**Tech Stack:** Google Apps Script V8, SpreadsheetApp/DriveApp/PropertiesService/LockService, Node.js `assert` + `vm` tests.

---

### Task 1: Lock the source rules with failing tests

**Files:**
- Create: `tests/project24_full_allocation_rebuild.test.js`
- Reference: `Project_24/引当.js`
- Reference: `Project_24/棚卸.js`

- [ ] Add VM loading and stubs matching the existing Project_24 tests.
- [ ] Add failing tests for strict SKU normalization: remove only Yahoo/GoQ terminal `a`/`b`, retain EMS codes ending in a real letter, retain variants such as `YMNGD08-1` vs `YMNGD08-2`.
- [ ] Add failing tests that `EMS番号=棚卸...` is excluded from factual supply.
- [ ] Add failing tests for GoQ snapshot reduction by `受注番号+商品ID`, including latest quantity zero and two different item IDs remaining separate.
- [ ] Run `node tests\project24_full_allocation_rebuild.test.js` and confirm the expected failures.

### Task 2: Implement the pure FIFO rebuild engine

**Files:**
- Create: `Project_24/全件再計算.js`
- Modify: `tests/project24_full_allocation_rebuild.test.js`

- [ ] Implement `全件再計算_SKU正規化_`, `全件再計算_発送最新化_`, `全件再計算_実EMS行_`.
- [ ] Implement quantity-interval FIFO consumption without per-unit expansion.
- [ ] Consume shipped rows only from supplies whose arrival date is on or before the ship date.
- [ ] Account for current immediate rows, reserve Yahoo `a`, then allocate current Korean backorders by order timestamp, order number, and row.
- [ ] Reserve remaining backorders to future EMS rows without setting arrival fields.
- [ ] Return summary, allocation detail, issues, blocked SKU set, P writes, history rows, and current-order writes.
- [ ] Block any SKU with ambiguous matching, unmet shipped consumption, invalid source values, or unexplained EMS excess after Yahoo/current demand.
- [ ] Add regression fixtures for `YMNGD08-1`, Yahoo `b` exclusion, split variants, post-shipment arrivals, partial allocation, and blocked excess.
- [ ] Run the focused test until green.

### Task 3: Add live-source adapters and preview sheets

**Files:**
- Modify: `Project_24/全件再計算.js`
- Modify: `Project_24/棚卸.js`
- Modify: `Project_24/引当.js`
- Modify: `tests/project24_full_allocation_rebuild.test.js`

- [ ] Read the latest CP932 Yahoo CSV through the existing Drive reader and validate required headers, integer quantities, duplicate `code+sub-code`, file ID, update time, and MD5.
- [ ] Read factual EMS rows from the external EMS list with row numbers and all fields needed for P/history; collect legacy synthetic rows separately.
- [ ] Read `発送済み` directly, normalize headers, reduce snapshots, and reject ambiguous latest rows.
- [ ] Read current orders while preserving Taiwan/China routes and exact destination row numbers.
- [ ] Compute deterministic signatures over all relevant source fields.
- [ ] Add `🧪 全件再計算プレビュー` to the main menu.
- [ ] Render `全件再計算_サマリ`, `全件再計算_割当明細`, and `全件再計算_要確認`.
- [ ] Store the serialized plan and signatures in hidden `全件再計算_内部` chunks.
- [ ] Verify preview only writes the four preview/internal sheets.

### Task 4: Implement guarded apply and retire synthetic shelf writes

**Files:**
- Modify: `Project_24/全件再計算.js`
- Modify: `Project_24/引当.js`
- Modify: `Project_24/引当履歴.js`
- Modify: `Project_24/棚卸.js`
- Modify: `tests/project24_full_allocation_rebuild.test.js`

- [ ] Add `✅ 全件再計算を反映` and require typed confirmation.
- [ ] Re-read all sources and reject apply when any signature differs from preview.
- [ ] Create and verify Drive backups for both the active allocation spreadsheet and the external purchasing spreadsheet.
- [ ] Set a persistent in-progress/failure guard that prevents normal allocation while apply is incomplete.
- [ ] Mark exact legacy synthetic EMS rows `再計算除外`; do not alter their quantities, dates, or EMS labels.
- [ ] Replace the external P column from the approved write plan.
- [ ] Replace allocation history through a bulk helper in `引当履歴.js`.
- [ ] Clear and rewrite Korean backorder arrival/EMS fields while retaining Taiwan/China manual fields.
- [ ] Persist blocked SKUs and make normal `② 引き当て実行` treat them as zero stock.
- [ ] Run the existing renderer with ledger/P auto-updates suppressed, then clear the in-progress guard only on full success.
- [ ] Disable `棚卸箱をEMSリストへ追記` so a button assignment cannot recreate ghost stock.

### Task 5: Verify, deploy safely, and generate a live preview

**Files:**
- Test: `tests/project24_full_allocation_rebuild.test.js`
- Test: `tests/project24_zenken_kensan.test.js`
- Test: `tests/project24_arrived_box_color.test.js`

- [ ] Run all three Project_24 Node test files.
- [ ] Run `node --check` on every `Project_24/*.js` file.
- [ ] Inspect `git diff --check`, the scoped diff, and unrelated worktree changes.
- [ ] Commit only Project_24/tests/docs changes.
- [ ] Push with `powershell -ExecutionPolicy Bypass -File tools\gas_safe_push.ps1 Project_24`; never use raw `clasp push`.
- [ ] Generate only the live preview, confirm the known `YMNGD08-1` synthetic row is excluded and the SKU is blocked when its sources do not reconcile.
- [ ] Do not execute live apply until the preview has been reviewed.
