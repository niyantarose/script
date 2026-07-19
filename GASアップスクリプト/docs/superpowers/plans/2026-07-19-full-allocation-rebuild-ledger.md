# Full Allocation Rebuild v3 Ledger Implementation Plan

> Execute with test-driven development. Never apply rebuilt data to live sheets before a preview has been generated and reviewed.

**Goal:** Rebuild allocation from real EMS arrivals, GoQ shipped facts, current orders, and Yahoo `a` free inventory while keeping the current reservation ledger as the only source of truth.

**Architecture:** `全件再計算.js` contains a pure chronological engine and GAS adapters. It produces replacement plans for `取り置き台帳` and `EMS在庫移動台帳`. The existing ④ flow alone derives P assignments, Korean arrival fields, colors, and output sheets.

## Task 1: Failing source-rule tests

- Create `tests/project24_full_allocation_rebuild.test.js`.
- Test source-aware SKU normalization, synthetic EMS exclusion, GoQ latest-snapshot reduction, shipment cutoff, Yahoo floor, oldest current order, excess blocking, and current ledger schemas.
- Run the focused test and confirm RED.

## Task 2: Pure rebuild engine

- Add `Project_24/全件再計算.js`.
- Normalize factual inputs and consume quantity intervals without per-unit expansion. Treat shipped Yahoo `a` as EMS-external, and prioritize valid EMS order-number tags over FIFO.
- Build shipped ledger rows, active Korean-backorder rows, Yahoo/immediate movement rows, future-P preview details, summaries, issues, and blocked SKUs.
- Keep deterministic IDs and ordering.
- Run focused and existing Project_24 tests until GREEN.

## Task 3: Live adapters and preview

- Read and validate the latest CP932 Yahoo CSV using existing helpers.
- Read external EMS rows, direct `発送済み`, current order details, and unresolved ledger work.
- Compute deterministic signatures.
- Add preview menu and render summary/detail/review/internal sheets.
- Verify preview changes only preview/internal sheets.

## Task 4: Guarded replacement and ④ integration

- Revalidate signatures and create two Drive backups.
- Archive old operational ledger sheets and refuse unresolved physical-return work.
- Replace both ledgers once, persist blocked SKUs, and exclude them in normal ④/P planning.
- Clear only old Korean derived arrival/EMS fields, then run existing ④ to derive P, Korean arrival/EMS, colors, and outputs.
- Preserve Taiwan/China manual dates and factual EMS fields.

## Task 5: Verify and deploy safely

- Run all Project_24 tests, syntax checks, and diff checks.
- Commit only scoped changes.
- Push with `tools/gas_safe_push.ps1 Project_24`; never raw `clasp push`.
- Generate only a live preview and verify `YMNGD08-1` before any live apply.
