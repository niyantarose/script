# Project_24 先行引当・現物固定・取り置き画面再設計 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 未着EMSを先行引当として注文分類へ含めながら、棚で数えた現物を受注番号・SKU単位で固定し、全件再計算でも移動させない。取り置き登録は注文単位の要作業画面へ変更し、全件照合は読み取り専用の「引当状況一覧」とGoQ差分で行う。

**Architecture:** 現在数量の正は引き続き `取り置き台帳` とし、各有効行へ `引当段階`（先行／到着済／現物確認済み）を追加する。通常④と全件再計算は同じ純粋計算エンジンを呼び、P列・入荷予定・分類・引当状況・現物出力はその1回の結果から派生させる。先行は論理分類へ含めるが、納品書ピック・⑤箱締め・Yahoo余り出力からは型とガードで除外する。人の入力は洗い替え画面から分離した非表示シートへ永続化し、初回だけ移行シートで旧「開始前在庫」と消失差分を確認する。

**Tech Stack:** Google Apps Script V8、Google Sheets、DriveApp、Node.js `assert`＋`vm` による純粋関数テスト。

**Spec:** `docs/superpowers/specs/2026-07-22-preallocation-physical-lock-design.md`

## Global Constraints

- 作業場所は worktree `GASアップスクリプト/.worktrees/full-allocation-rebuild-v3/GASアップスクリプト`、ブランチ `codex/full-allocation-rebuild-v3`。
- 各編集セッション開始前に `git pull origin codex/full-allocation-rebuild-v3` を実行する。
- **警告(2026-07-22判明): この worktree で `gas_pull_sync.ps1 Project_24` を実行してはならない。** 同期基準が旧版（最終push時点）のため、未pushの三段階実装をオンラインの旧コピーで巻き戻して sync コミットしてしまう（Codex セッションで2回、Claude 引き継ぎで1回発生。復元コミット参照）。オンライン編集の取り込みは Task 9 の安全push直前に、`git diff` で「取り込まれた内容が旧版と一致していないか」を確認しながら手動で行う。誤って実行した場合は直前コミットから `Project_24/取り置き計算.js` と `Project_24/取り置き台帳.js` を checkout で復元し、全テストを回す。
- 素の `clasp push` は禁止。反映は `powershell -ExecutionPolicy Bypass -File tools\gas_safe_push.ps1 Project_24` だけを使う。
- 実データの反映は必ずプレビュー、4入力（GoQ・EMS・Yahoo・台帳）の署名再検証、バックアップ成功後に行う。テスト中に本番シートへ書かない。
- 書き込み系公開入口は `直列_` で排他する。内部関数名の末尾 `_` は `google.script.run` の公開入口に使わない。
- 台湾・中国・ダニエル便の既存別ルート処理、GoQ/Yahoo API更新、手塗り色の保存は変更しない。
- 既存 `取置ID` は変更しない。同じ元EMS・同じ注文・同じSKUの昇格ではIDを維持する。
- 既存9個の `project24*.test.js` をすべて回帰させる。段階値を追加しても旧台帳行は読めるようにする。
- 先行段階は論理的な確保であり、`物理出荷可`、納品書・ピック対象、⑤未ピック、EMS在庫移動、Yahoo追加対象へ含めない。対象EMSに先行行が残る場合は⑤を停止する。
- シート書式・値・入力規則は配列を作って範囲単位で反映する。データ行ループ内の `getRange`、`setValue`、`setBackground`、`setDataValidation` を禁止する。

## Execution Phases

- **Phase 1 — 台帳スキーマ・三段階計算・物理ガード:** Task 1〜4。日常④と全件再計算の共通エンジン、分類、P列、⑤・Yahoo・ピック境界までを先に完成させる。
- **Phase 2 — 初回現物移行:** Task 5。旧開始前在庫と最後の旧④バックアップ差分を、保留を計算へ混ぜず安全に移行する。
- **Phase 3 — 作業画面・GoQ差分・統合:** Task 6〜8。入力永続化、要作業画面、引当状況一覧、反映フローを完成させる。
- **Release verification:** Task 9。全回帰、レビュー、安全push、実データプレビューを行う。

---

### Task 1: 三段階台帳ドメインと不変条件をテスト駆動で追加

**Files:**
- Modify: `Project_24/取り置き計算.js`
- Modify: `Project_24/取り置き台帳.js`
- Modify: `tests/project24_reservation_domain.test.js`
- Create: `tests/project24_three_stage_allocation.test.js`

**Interfaces:**
- `TORIOKI_STAGE = { PLANNED:'先行', ARRIVED:'到着済', PHYSICAL:'現物確認済み' }`
- `取り置き_段階正規化_(row, emsStatusByNo)`：新列があれば採用し、旧EMS行はEMS状態から先行／到着済へ変換する。旧 `開始前在庫` は自動で現物扱いせず `要移行` を返す。
- `取り置き_段階別集計_(rows, movements, emsStatusByNo)`：行キー別に `現物確認済み数量`、`到着済引当数量`、`先行引当数量`、`合計確保数量` と行内訳を返す。
- `取り置き_不変条件検証_(orders, ledgerRows, supplies)`：注文超過、段階重複、EMS供給超過、現物固定注文の消滅・数量減を errors と warnings に分ける。
- `取り置き_現物確認変換計画_(inputRows, ledgerRows, supplies, now)`：入力数量を新規加算せず、同じ注文・SKUの有効行を現物へ変換する。到着済行は元EMS消費を保持し、先行行は未着供給を解放する。元EMS不明現物は到着済供給をSKU FIFOで控除し、推測控除先は `供給控除EMS` に残す。現物確認済みを減らす入力は拒否し、解除フローへ誘導する。

- [x] **Step 1: 三段階のREDテストを書く**

`tests/project24_three_stage_allocation.test.js` に次を追加する。

1. 同じ注文・SKUに現物1、到着済1、先行1があれば段階別集計は各1・合計3。
2. `10117608 / MRBLUE41b / 注文14` は現物6＋先行8＋未引当0で、同じ8個を到着済へも数えない。
3. 既存EMS行2個の現物確認は数量を2個増やさず、段階だけ現物へ変換する。
4. 現物確認済み3個を2個へ減らす入力はエラーにし、台帳を書き換えない。
5. 注文数量超過、同一取置ID重複、EMS数量超過を停止エラーにする。
6. 現物固定の注文が消えた／数量が減った場合は自動解除せず停止エラーにする。
7. 旧 `開始前在庫` は `要移行` となり、通常集計へ無条件に混ぜない。
8. 先行2個を現物2個へ変換すると未着箱の2個は解放され、台帳合計は2のまま。
9. 元EMS不明現物1個は同SKUの最古到着済箱から1個を控除し、箱がなければ要確認。

- [x] **Step 2: REDを確認**

Run: `node tests/project24_three_stage_allocation.test.js`

Expected: 新しい定数・関数が未定義のためFAIL。

- [x] **Step 3: 台帳スキーマと純粋関数を実装**

`TORIOKI_CFG.台帳HDR` の既存列を維持したまま末尾へ `引当段階`、`EMS到着予定日`、`現物確認日時`、`現物確認メモ`、`供給控除EMS` を追加する。既存14列の読み込みを壊さず、保存時だけ新スキーマへ展開する。新規有効行では `開始前在庫` を生成しない。

`取り置き_集計_` の既存戻り値は維持し、段階別結果を追加する。`activeByKey` は三段階合計を返すため既存の残必要数計算を壊さない。

- [x] **Step 4: 現物確認変換を実装**

変換対象は同じ `取り置き_行キー_` の有効行だけとし、元EMS番号・元EMS商品コード・取置ID・登録日時を維持する。複数行にまたがる場合は既存の行順で必要数だけ変換し、行の一部なら同じ証跡を持つ決定的な分割IDを作る。確認日時・確認メモだけを更新する。

- [x] **Step 5: GREENと既存ドメイン回帰を確認**

Run:

```powershell
node tests/project24_three_stage_allocation.test.js
node tests/project24_reservation_domain.test.js
```

Expected: 両方 `ALL PASS` または `ALL TESTS PASSED`。

- [x] **Step 6: コミット**

```powershell
git add Project_24/取り置き計算.js Project_24/取り置き台帳.js tests/project24_reservation_domain.test.js tests/project24_three_stage_allocation.test.js
git commit -m "feat(Project_24): 取り置き台帳を三段階管理へ拡張"
```

> 引き継ぎメモ(2026-07-22 Claude): Codex停止時点の未コミットRED2本（到着予定日の到着済化禁止・空白入力スキップ）を実装してGREEN化。
> 旧台帳互換（種別ありEMS番号なし行=到着済扱い）を追加して既存workflowテストの回帰2件を解消。旧変換関数2つ（死にコード）を削除。全10スイートPASS。

---

### Task 2: 全件再計算を「現物固定→Yahoo保護→到着済→先行」へ変更

**Files:**
- Modify: `Project_24/全件再計算.js`
- Modify: `tests/project24_full_allocation_rebuild.test.js`
- Modify: `tests/project24_three_stage_allocation.test.js`

**Interfaces:**
- `全件再計算_供給段階_(status)`：`到着済` は到着済、`在庫反映済み` は過去締め済み、その他の実EMSは先行。
- `全件再計算_現物固定を予約_(physicalRows, orders, supplies)`：現物を元注文へ固定し、元EMSありはその供給を一度消費、元EMS不明は同SKU到着済供給を古い順に控除する。控除不能分は重要issueにする。
- `全件再計算_再構築_(input)` の戻り値に `stageSummary`、`promotionRows`、`futureReservations`、`invariantErrors` を追加する。

- [x] **Step 1: 実例を含むREDテストを追加**

`tests/project24_full_allocation_rebuild.test.js` へ次を追加する。

1. `10117428 / MRBLUE42-6b / 注文3 / 現物確認済み3` は未着EMSが2個でも現物3個のまま残り、新しい割当を作らない。
2. `10117608 / MRBLUE41b / 注文14 / 現物6 / 未着EMS8` は現物6＋先行8で全数確保し、供給消費は8だけ。
3. Yahoo `a` 数量は取り寄せ供給に使わず、到着済EMSの残量計算から先に保護する。`b` は引き続き無視する。
4. 到着済供給を先にFIFOし、残需要だけ未着供給へFIFOする。
5. 同じEMS・同じ注文の先行行が、EMS状態変更後も同じ取置ID・数量で到着済へ昇格する。
6. 箱コードまたは数量が変わった場合、現物行は維持し、差分だけ `issues` と停止対象へ出す。
7. 到着済の余りだけをYahoo追加候補にし、未着の余りは候補へ出さない。
8. 元EMS不明現物は最古到着済箱から控除され、到着済箱が無ければ要確認で反映停止。
9. 先行→現物変換で元の未着供給が解放され、別の残需要へ再配分できる。

- [x] **Step 2: REDを確認**

Run: `node tests/project24_full_allocation_rebuild.test.js`

Expected: 現行の `futureReservations: []`、到着済だけの割当、開始前在庫の無条件保持に関する新テストがFAIL。

- [x] **Step 3: 供給を段階化し、現物固定を最初に予約**

`全件再計算_実EMS行_` の戻り値へ `stage` と `status` を保持する。SKUごとの処理では、元EMS付き現物行をその供給から先に差し引き、元EMS不明現物は同SKUの到着済供給を到着日・EMS番号順で控除する。控除先は `供給控除EMS` に残し、`元EMS番号` は由来不明のままにする。どの箱からも控除できない数量は推測で外部在庫にせず停止対象へ出す。

- [x] **Step 4: Yahoo保護・到着済・先行の順で割り当て**

現物固定後、Yahoo自由在庫の保護を確定し、その後に現在取寄せへ到着済供給、未着供給の2パスで割り当てる。新規台帳行には段階と到着予定日を保存する。`在庫反映済み` 供給は現役注文へ使わない。

- [x] **Step 5: 停止条件と昇格を実装**

同じ決定IDが既存先行行にあれば置換ではなく段階更新とし、箱内容不一致、注文超過、現物注文の消失は `重要` issue と `invariantErrors` へ入れる。全件反映アダプターは `invariantErrors` が1件でもあればプレビューのみ作り、運用台帳を置換しない。

- [x] **Step 6: GREENと回帰を確認**

Run:

```powershell
node tests/project24_full_allocation_rebuild.test.js
node tests/project24_three_stage_allocation.test.js
node tests/project24_yahoo_stock_export.test.js
```

Expected: 全PASS。

- [x] **Step 7: コミット**

> 実装メモ(2026-07-22 Claude): 実EMS行_が未着を捨てる入口を開放。takeは段階フィルタ(単一or配列)化——
> 歴史説明(発送済み/即納/Yahoo/開始前在庫)は到着済+過去締め済み(旧挙動と同一)、現役取寄せは到着済のみ、
> 先行パスは未着のみ。現物固定はGASアダプター(計画を作る_)から現在台帳の取り置き中行を
> physicalRows/currentPlannedへ接続済み。invariantErrorsはapplyBlocked経由で反映を停止する。

```powershell
git add Project_24/全件再計算.js tests/project24_full_allocation_rebuild.test.js tests/project24_three_stage_allocation.test.js
git commit -m "feat(Project_24): 全件再計算へ現物固定と先行引当を統合"
```

---

### Task 3: 通常④引当と注文分類を三段階数量で統一

**Files:**
- Modify: `Project_24/引当.js`
- Modify: `Project_24/P列自動記入.js`
- Modify: `tests/project24_pre_arrival_plan.test.js`
- Modify: `tests/project24_arrived_box_color.test.js`
- Create: `tests/project24_three_stage_classification.test.js`

**Interfaces:**
- `引当計画_行へ反映_`：各受注明細行へ `現物確認済み数量`、`到着済引当数量`、`先行引当数量`、`未引当数量`、`主段階` を付与する。
- `注文充足集計_(arr)`：注文内全行の注文数と4数量を集計し、不変条件を確認する。
- `注文区分判定_(arr, paid, cod)`：今回新規割当の有無ではなく、先行を含む確保総数で `確保0→一部→希望日→代引→未入金→入金済` の優先順位から必ず1分類へ返す。
- `注文物理出荷可_(arr)`：不足0かつ先行0の注文だけtrue。分類上の出荷可能とは独立に持つ。
- `引当行状態_`：先行=薄青、到着済=現行ラベンダー、現物=緑、不足=白/グレー、要確認=黄/赤。
- `P列計画_確定割当_(plan)`：到着済行だけを通常台帳確定へ渡し、先行行は段階 `先行` の計画行として別に返す。

- [x] **Step 1: 分類のREDテストを書く**（+段階対応の行状態・物理出荷可・旧経路互換もテスト済み）

`tests/project24_three_stage_classification.test.js` に、先行だけ／到着済だけ／現物混在の各注文を作り、次を検証する。

1. 全数先行＋入金済→出荷可能（行表示は先行）。
2. 一部先行→部分在庫。
3. 全数先行＋未来希望日→希望日待ち。
4. 全数先行＋未入金→出荷GO未入金。
5. 全数先行＋代引き→出荷可能かつ代引き表示。
6. 確保0→引当待ち。
7. 混在注文も受注番号単位で1分類だけへ出る。
8. 箱到着で分類先と取置IDは変わらず、表示段階・色だけ先行→到着済へ変わる。
9. 全数先行の出荷可能は `物理出荷可=false`、全数到着済／現物は `true`。

- [x] **Step 2: REDを確認**

Run: `node tests/project24_three_stage_classification.test.js`

Expected: 新しい集計関数が未定義、または現行の「黄色あり」条件によりFAIL。

> 進捗メモ(2026-07-22 Claude): 純粋層は実装済み＝注文充足集計_/注文区分判定_(確保総数ベース・優先順
> 確保0→一部→希望日→代引→未入金→入金済)/注文物理出荷可_/引当行状態_の段階分岐(段階付与フラグで
> 旧経路互換)。④呼び出しは注文区分判定_(byOrder, paid, cod)へ変更済み。Step 3以降(供給統一・
> 分類シート列・P列色昇格)が未着手。

> 進捗メモ2(2026-07-22 Claude, 44ed48f): ④への段階配線まで完了。**重要な契約変更**=旧開始前在庫(要移行)は
> activeByKeyに数える(除外すると④が二重割当する本番回帰だった。Codexの旧契約テスト3本を新契約へ更新)。
> EMS状態は「未着」だけ先行(不明・空・在庫反映済みは到着済へ倒す)。④新規行は引当段階=到着済。
> 引当計画_行へ反映_が現物/到着済/先行/未引当/主段階を行へ付与済み。
> Step 3の残り=P列計画_純計算_の先行対応(未着箱の薄青P列を④へ吸収)、Step 4=分類シートへの段階列出力、
> Step 5=P列色の昇格確認。

- [ ] **Step 3: 通常④の供給と台帳計画を統一**

独立していた `P列計画_純計算_` の割当判断を共通三段階エンジンへ吸収する。`引当実行_本体_` は到着済・未着両方の実EMSを共通入力へ渡し、その結果から台帳行、P列テキスト、薄青背景、入荷予定列を作る。通常④でも `取り置き_不変条件検証_` を通過後だけ保存する。全件再計算は歴史入力を再構築した後、同じ現在需要アロケータを呼ぶ。

- [x] **Step 4: 注文分類と出力列を更新**（112a66d: 5分類シートへ段階6列+状態の理由を出力。
  分類は確保総数ベース済み。残り=Step 3のP列計画_純計算_吸収とStep 5のP列色昇格）

5分類シートの共通列へ `引当段階`、`現物確認済み`、`到着済引当`、`先行引当`、`不足`、`状態の理由` を追加する。既存の注文情報・取り置きメモ・代引き表示は残す。1受注番号につき出力先配列を1つだけ選ぶ。

- [ ] **Step 5: P列と色の昇格を確認**

先行P列は薄青、到着済への状態変更時はP列テキストが同じでも背景を通常色へ戻す。受注明細と分類シートの色は手塗りを引き継がず、毎回段階から再設定する。

- [ ] **Step 6: GREENと関連回帰を確認**

Run:

```powershell
node tests/project24_three_stage_classification.test.js
node tests/project24_pre_arrival_plan.test.js
node tests/project24_arrived_box_color.test.js
node tests/project24_reservation_domain.test.js
```

Expected: 全PASS。

- [ ] **Step 7: コミット**

```powershell
git add Project_24/引当.js Project_24/P列自動記入.js tests/project24_pre_arrival_plan.test.js tests/project24_arrived_box_color.test.js tests/project24_three_stage_classification.test.js
git commit -m "feat(Project_24): 先行を含む注文分類と自動色分けを統一"
```

---

### Task 4: 先行を納品書・⑤箱締め・Yahoo出力から遮断

**Files:**
- Modify: `Project_24/引当.js`
- Modify: `Project_24/引当履歴.js`
- Modify: `Project_24/Yahoo在庫変更出力.js`
- Modify: `Project_24/受注明細個別ボタン.js`
- Modify: `tests/project24_reservation_workflow.test.js`
- Modify: `tests/project24_yahoo_stock_export.test.js`
- Create: `tests/project24_physical_operation_guard.test.js`

**Interfaces:**
- `物理オペ対象行_(row)`：`引当段階` が到着済または現物確認済みの有効行だけtrue。
- `物理出荷対象注文_(orderSummary)`：不足0かつ先行0の注文だけtrue。
- `便締め_先行残検査_(ledgerRows, targetEmsSet)`：対象EMSに有効な先行行が1件でもあれば停止理由を返す。
- `Yahoo変更_対象行_`：到着済余りだけを受け付け、先行・未着の行は除外理由付きで拒否する。

- [ ] **Step 1: 物理境界のREDテストを書く**

1. 先行だけで全数確保した注文は分類上出荷可能でも、納品書・ピック対象へ出ない。
2. 到着済＋現物で全数確保した注文は物理出荷対象になる。
3. ⑤の対象EMSに先行行が残れば、未ピック確認・ステータス変更・移動台帳保存の前に停止する。
4. ⑤の未ピック一覧は到着済／現物だけを表示し、先行を件数へ含めない。
5. `日本在庫` とYahoo出力は到着済余りだけを対象にし、未着供給の余りを出力しない。
6. 個別引当・今回入荷EMS表示は未着供給を現物ピック済みとして表示しない。

- [ ] **Step 2: REDを確認**

Run:

```powershell
node tests/project24_physical_operation_guard.test.js
node tests/project24_yahoo_stock_export.test.js
```

Expected: 段階フィルタと⑤先行残ガードが未実装のためFAIL。

- [ ] **Step 3: 物理対象の共通判定を実装**

物理対象判定は `取り置き計算.js` の段階正規化を使い、各画面で文字列を個別解釈しない。納品書・ピックに渡す注文行へ `物理出荷可` を明示し、分類名だけで現物ありと判断しない。

- [ ] **Step 4: ⑤箱締めをガード**

`到着済を在庫反映済みへ本体_` は対象EMS選択直後に先行残を検査する。残っていれば「④を実行して到着済へ昇格してください」と表示して中止する。未ピック一覧、surplus、EMS在庫移動計画は物理対象行だけから作る。台帳読込失敗を握り潰して締める現行catchは廃止し、安全側に停止する。

- [ ] **Step 5: Yahoo出力と個別表示をガード**

`日本在庫` の各行へ段階を保持し、Yahoo出力は `状態=到着済余り` だけを許す。入力EMSが未着または一覧にない場合は空出力で成功させず中止する。個別引当は未着供給を選択候補・現物余り表示へ入れない。

- [ ] **Step 6: GREENと回帰を確認**

Run:

```powershell
node tests/project24_physical_operation_guard.test.js
node tests/project24_yahoo_stock_export.test.js
node tests/project24_reservation_workflow.test.js
```

Expected: 全PASS。

- [ ] **Step 7: コミット**

```powershell
git add Project_24/引当.js Project_24/引当履歴.js Project_24/Yahoo在庫変更出力.js Project_24/受注明細個別ボタン.js tests/project24_reservation_workflow.test.js tests/project24_yahoo_stock_export.test.js tests/project24_physical_operation_guard.test.js
git commit -m "feat(Project_24): 先行引当を物理出荷と便締めから遮断"
```

---

### Task 5: 一度きりの現物確認移行を安全に実装

**Files:**
- Create: `Project_24/現物確認移行.js`
- Modify: `Project_24/引当.js`
- Create: `tests/project24_physical_migration.test.js`

**Interfaces and sheets:**
- Sheet `現物確認移行`。
- Hidden audit sheet `現物確認移行_差分`。
- `現物確認移行_候補計算_(currentLedger, baselineLedger, currentOrders)`：現在の旧開始前在庫と、指定基準バックアップ－現在数の正差分を候補化する。
- `現物確認移行_反映計画_(candidates, choices, ledger, orders, supplies, now)`：`現物確認済みにする`、`計算上の引当だったので解除`、`保留` を入力キー単位で検証し、適用可能行・未適用行・新台帳・差分を返す。
- Public entries: `現物確認移行を作成()`、`現物確認移行を反映()`。どちらも `直列_` を通す。

- [ ] **Step 1: REDテストを書く**

1. 旧開始前在庫61行相当を候補へ出し、数量を勝手に現物へ変えない。
2. 指定バックアップ3・現在2の `10117428 / MRBLUE42-6b` は差1を復元候補へ出す。
3. 比較基準は `取り置き台帳_全件再計算前_20260721_123350` だけを使い、他の幽霊確保を自動採用しない。
4. `現物確認済みにする` は注文数量までだけ復元し、元の現物・EMS行と二重加算しない。
5. `解除` は旧開始前在庫を手動解除へ変更し、EMS由来の現物確認済みを自動解除しない。
6. `保留` 行は台帳・需要・供給のいずれにも加算されず、全件反映ガードだけを残す。
7. 1入力キーのエラーは他の有効な入力キーを破棄せず、エラー入力と理由を残す。
8. バックアップ失敗・入力署名変化・数量超過では保存処理自体を開始しない。

- [ ] **Step 2: REDを確認**

Run: `node tests/project24_physical_migration.test.js`

Expected: 新規関数未定義でFAIL。

- [ ] **Step 3: 候補生成と注文一括入力を実装**

基準シート `取り置き台帳_全件再計算前_20260721_123350` を明示的に読み、受注番号＋SKU単位で現在台帳と比較する。以前の確保は棚確認の事実ではないため、差分を自動復元しない。シートには注文単位の一括選択列と商品行ごとの移行数量・選択列を設け、一括選択は空欄の商品行だけへ伝播する。

- [ ] **Step 4: バックアップ・署名・差分付き反映を実装**

反映前にスプレッドシート全体のDriveコピー、取り置き台帳、EMS在庫移動台帳のスナップショットを作成する。候補作成時に保存した注文・EMS・Yahoo・台帳署名と再比較し、異なれば中止する。入力キー別検証後、適用可能な変更を1回の台帳保存で反映し、未適用入力は残す。保存成功後に変更前後を `現物確認移行_差分` へ追記する。

- [ ] **Step 5: 通常処理ガードへ接続**

未解決の旧開始前在庫または保留候補がある間、通常全件反映は停止する。通常④は既存運用を止めず、旧行・保留を計算へ混ぜない状態で要確認メッセージを表示する。移行完了後は有効な `開始前在庫` 行を残さず、現物確認済みまたは手動解除へ統合する。

- [ ] **Step 6: GREENを確認してコミット**

Run:

```powershell
node tests/project24_physical_migration.test.js
node tests/project24_full_allocation_rebuild.test.js
```

Expected: 全PASS。

```powershell
git add Project_24/現物確認移行.js Project_24/引当.js tests/project24_physical_migration.test.js
git commit -m "feat(Project_24): 旧取り置きの現物確認移行を追加"
```

---

### Task 6: 手入力の永続保存と注文単位の取り置き登録を実装

**Files:**
- Modify: `Project_24/取り置き台帳.js`
- Modify: `Project_24/引当.js`
- Modify: `tests/project24_reservation_workflow.test.js`
- Create: `tests/project24_reservation_work_screen.test.js`

**Interfaces and sheets:**
- Hidden sheet `取り置き入力保存` with headers: `入力キー,受注番号,SKU,商品コード,棚確認,取り置きメモ,確認メモ,注文作業メモ,未反映現物確認数量,入力エラー,最終表示日時,更新日時`。
- Hidden sheet `取り置き入力履歴`：終了後90日を超えた入力保存を追記式で移し、削除せず監査保持する。
- `取り置き_入力キー_(row)`：`受注番号|正規化SKU`。SKUがない場合だけ正規化商品コードを使う。
- `取り置き_入力保存マージ_(generatedRows, savedRows, sheetRows)`：洗い替え前のシート入力を内部保存へ先にupsertし、生成行へ復元する。
- `取り置き_作業対象判定_(order, viewMode)`：`要作業`、`部分在庫`、`希望日待ち・現物あり`、`先行引当`、`すべて` を純粋判定する。
- `TORIOKI_CFG.初期HDR` はユーザー向け列だけに変更し、内部キーは非表示列または入力保存で保持する。

Visible headers:

`受注番号,氏名,GoQ受注ステータス,お届け希望日,支払い,注文メモ,ひとことメモ,取り置きメモ,GoQ差分,商品コード,SKU,商品名,注文数量,現物確認済み,到着済引当,先行引当,不足,状態の理由,棚確認,今回の現物確認数量,確認メモ`

- [ ] **Step 1: 永続化と表示対象のREDテストを書く**

次を `tests/project24_reservation_work_screen.test.js` へ追加する。

1. シートを再生成しても取り置きメモ、確認メモ、棚確認、注文作業メモが入力キーで復元される。
2. 同一注文の複数商品は連続し、注文境界の太罫線は最終商品行にだけ付く。
3. 既存5値の棚確認プルダウンが維持され、選択だけでは数量が変わらない。
4. 初期 `要作業` は部分在庫を常に表示し、希望日待ちは到着済／現物ありだけ、出荷可能は棚確認必要／GoQ差分ありだけ表示する。
5. 先行だけ、未引当だけ、処理済み出荷可能は初期表示から外れるが `すべて` には出る。
6. `今回の現物確認数量` は既存引当を現物へ変換し、再反映しても加算されない。
7. 旧内部列8個は可視ヘッダーに含まれない。
8. 1入力キーがエラーでも他の有効キーは反映され、エラー行の数量・メモ・理由は再表示される。
9. GoQ取込・画面更新の途中でも未反映数量は非表示シートへ退避され、黙って空にならない。
10. 書式・入力規則処理はデータ行ループ内でRange APIを呼ばず、受注1,000行のモックで呼出回数が行数比例しない。

- [ ] **Step 2: REDを確認**

Run: `node tests/project24_reservation_work_screen.test.js`

Expected: 新規関数・新ヘッダーが未実装のためFAIL。

- [ ] **Step 3: 内部保存を実装**

更新処理の最初に現在の入力値と未反映数量を `取り置き入力保存` へupsertし、生成後に復元する。保存シートは保護対象の内部シートとして非表示にする。終了から90日までは同シートに残し、それより古い行は `取り置き入力履歴` へ一括追記してから現行保存から外す。履歴は明示的な消去操作以外で削除しない。

- [ ] **Step 4: 注文単位の作業画面へ置換**

受注明細と段階別台帳集計から注文ブロックを生成し、内部列を外す。既存の条件付き書式関数を21列へ追従させ、先行薄青、到着済ラベンダー、現物緑、要確認黄/赤、キャンセル灰を毎回再適用する。値は1回の `setValues`、背景は1回の `setBackgrounds`、入力規則は列範囲またはRangeList、罫線は注文境界のRangeListで適用し、行ごとのRange呼出を残さない。

- [ ] **Step 5: 表示切替を実装**

シート上部の固定セルへ表示モードのプルダウンを置き、`onEdit` から再生成する。既定は `要作業`。表示モード変更は入力保存後に行い、非表示になった行のメモを失わない。

- [ ] **Step 6: 反映処理を現物変換へ接続**

`取り置き登録を反映本体_` は棚戻し処理と現物確認変換を入力キー単位で検証する。エラーキーを除いた適用可能な全変更から新台帳を作り、台帳全体は1回だけ保存する。保存成功後は適用済みキーの数量入力だけを空へ戻し、エラーキーの数量・理由と全メモ・棚確認は保持する。台帳保存そのものが失敗した場合は全キーを未適用のまま残す。

- [ ] **Step 7: GREENと既存ワークフロー回帰を確認**

Run:

```powershell
node tests/project24_reservation_work_screen.test.js
node tests/project24_reservation_workflow.test.js
```

Expected: 全PASS。旧17列前提テストは新しい列数・公開仕様へ更新するが、台帳の一括保存・キャンセル戻し・旧ボタン互換は維持する。

- [ ] **Step 8: コミット**

```powershell
git add Project_24/取り置き台帳.js Project_24/引当.js tests/project24_reservation_workflow.test.js tests/project24_reservation_work_screen.test.js
git commit -m "feat(Project_24): 取り置き登録を注文単位の要作業画面へ刷新"
```

---

### Task 7: 全注文の引当状況一覧とGoQ差分を追加

**Files:**
- Create: `Project_24/引当状況一覧.js`
- Modify: `Project_24/引当.js`
- Create: `tests/project24_allocation_status_list.test.js`

**Interfaces:**
- Sheet `引当状況一覧`（読み取り専用運用）。
- `引当状況_推奨GoQ_(orderSummary)`：分類・支払い・希望日・段階から推奨表示を返す。
- `引当状況_GoQ差分_(currentStatus, recommended, summary)`：`一致` または `差異あり` と理由を返す。GoQへ書き戻さない。
- `引当状況_一覧行_(orders, stageSummary, savedInputs)`：全注文を注文単位にまとめ、商品行は段階別数量と理由を保持する。

Visible headers:

`受注番号,氏名,現在のGoQステータス,推奨GoQステータス,GoQ差分,差異理由,分類,お届け希望日,支払い,商品コード,SKU,商品名,注文数量,現物確認済み,到着済引当,先行引当,不足,状態の理由,取り置きメモ`

- [ ] **Step 1: REDテストを書く**

1. 全数到着済だがGoQが取り寄せ中なら差異あり。
2. 全数先行の注文は推奨状態に「先行」を含め、到着済と誤表示しない。
3. 希望日未来、未入金、代引き、部分在庫を別理由で表示する。
4. 同じ注文は1分類、商品明細はすべて一覧に出る。
5. 一覧生成関数はGoQ更新関数や書き込みAPIを呼ばない。

- [ ] **Step 2: REDを確認**

Run: `node tests/project24_allocation_status_list.test.js`

Expected: 新規関数未定義でFAIL。

- [ ] **Step 3: 純粋差分と一覧生成を実装**

GoQの現在値は受注明細の `受注ステータス` から読む。推奨値は作業案内として生成し、API・CSVへは一切書き戻さない。差異理由は「未着先行のみ」「到着済み全数」「現物固定あり」「希望日待ち」「未入金」「不足N個」の組合せから決定する。

- [ ] **Step 4: 一覧シート書き出しを④へ接続**

④の検証成功後、5分類シートと同じ計画データから一覧を生成する。分類シートと一覧で別計算をしない。段階別自動色と注文境界罫線を適用し、手入力列は作らない。

- [ ] **Step 5: GREENを確認してコミット**

Run:

```powershell
node tests/project24_allocation_status_list.test.js
node tests/project24_three_stage_classification.test.js
```

Expected: 全PASS。

```powershell
git add Project_24/引当状況一覧.js Project_24/引当.js tests/project24_allocation_status_list.test.js
git commit -m "feat(Project_24): 引当状況一覧とGoQ差分を追加"
```

---

### Task 8: メニュー、プレビュー、原子的反映を統合

**Files:**
- Modify: `Project_24/引当.js`
- Modify: `Project_24/全件再計算.js`
- Modify: `Project_24/取り置き台帳.js`
- Modify: `tests/project24_reservation_workflow.test.js`
- Modify: `tests/project24_full_allocation_rebuild.test.js`

- [ ] **Step 1: メニューと公開入口のREDテストを追加**

メニューに次を追加し、公開入口が存在して `直列_` を通ることを検証する。

- `📋 取り置き登録を更新（要作業）`
- `📊 引当状況一覧を更新`
- `🔄 現物確認移行を作成`
- `✅ 現物確認移行を反映`
- 既存の `全件再計算プレビュー` と `全件再計算を反映`

- [ ] **Step 2: プレビューへ4入力署名と停止理由を追加**

プレビュー内部データへ GoQ、EMS、Yahoo、取り置き台帳、取り置き入力保存の各署名を保存する。反映時に再読込して1つでも違えば停止する。`invariantErrors`、箱内容差分、未解決移行、物理固定キャンセルをサマリと要確認シートへ出す。

- [ ] **Step 3: 書き込み順とロールバック境界を固定**

書き込み順は、バックアップ成功→台帳保存→入力保存→P列→受注明細派生列→5分類→引当状況一覧→監査差分とする。台帳保存前の失敗は無変更、台帳保存後の失敗はガードを残して通常処理を停止し、バックアップ名と停止段階を通知する。

- [ ] **Step 4: 回帰テストを通してコミット**

Run:

```powershell
node tests/project24_reservation_workflow.test.js
node tests/project24_full_allocation_rebuild.test.js
node tests/project24_physical_migration.test.js
```

Expected: 全PASS。

```powershell
git add Project_24/引当.js Project_24/全件再計算.js Project_24/取り置き台帳.js tests/project24_reservation_workflow.test.js tests/project24_full_allocation_rebuild.test.js
git commit -m "feat(Project_24): 三段階引当の安全な反映フローを統合"
```

---

### Task 9: 全回帰、コードレビュー、安全push、実データ検証

**Files:**
- Verify all modified Project_24 files and tests.

- [ ] **Step 1: 全Project_24テストを実行**

```powershell
$failed=@(); Get-ChildItem tests -Filter 'project24*.test.js' | Sort-Object Name | ForEach-Object { node $_.FullName; if($LASTEXITCODE-ne 0){$failed+=$_.Name} }; if($failed.Count){ throw ('FAILED: '+($failed-join ', ')) }
```

Expected: 既存9ファイル＋新規6ファイルがすべて終了コード0。

- [ ] **Step 2: GAS JavaScript構文を確認**

GASのトップレベル関数をNode構文検査できる一時コピー方式で、変更した全 `.js` を検査する。Expected: syntax error 0件。

- [ ] **Step 3: 差分レビュー**

`git diff --check`、`git status --short`、`git diff --stat 43beffc..HEAD` を確認し、Project_24、tests、docs以外の変更がないこと、デバッグ出力や一時ファイルがないことを確認する。`superpowers:requesting-code-review` で仕様17節の対象外と停止条件を重点レビューする。

- [ ] **Step 4: 安全push**

```powershell
powershell -ExecutionPolicy Bypass -File tools\gas_safe_push.ps1 Project_24
```

Expected: 3方向同期で衝突なし、Project_24だけが反映され、同期基準が更新される。衝突時は中止し、オンライン差分を先に取り込む。

- [ ] **Step 5: push後の独立検証**

空のscratchディレクトリへ Project_24 の `.clasp.json` をコピーして `clasp pull` し、ローカルProject_24と全ファイルのSHA-256を比較する。Expected: 全一致。

- [ ] **Step 6: 本番はプレビューだけ先に実行**

1. `現物確認移行を作成` で旧61行・71個と消失差分を表示する。
2. `10117428 / MRBLUE42-6b` が「以前3・現在2・差1」で候補に出ることを確認する。
3. `10117608 / MRBLUE41b` が現物6＋先行8として表示され、合計14を超えないことを確認する。
4. 先行だけの注文が薄青で各分類へ出ること、到着済・現物ありだけが要作業へ出ることを確認する。
5. メモ、棚確認、色分けが更新後も残ることを確認する。

ユーザーが移行候補を確認するまでは、`現物確認移行を反映` と `全件再計算を反映` は実行しない。

- [ ] **Step 7: 最終コミットと完了報告**

未コミットの安全push同期基準だけがあれば内容を確認してコミットする。完了報告にはテスト件数、push後SHA一致、プレビューで確認した実例、ユーザーが次に押すメニューを記載する。

---

## Self-review checklist

- [ ] `注文数量 = 現物確認済み + 到着済引当 + 先行引当 + 未引当` を、通常④と全件再計算の両方が同じ関数で検証する。
- [ ] 現物への変更は加算ではなく段階変換であり、再実行が冪等。
- [ ] 元EMS付き現物は元供給を1回だけ消費し、元EMS不明現物は同SKUの到着済供給をFIFO控除して `供給控除EMS` に証跡を残す。控除不能時は停止する。
- [ ] Yahoo `a` は自由在庫として保護し、`b` は物理在庫へ数えない。
- [ ] 先行→到着済で取置ID・注文先・数量が変わらない。
- [ ] 箱内容変更、キャンセル、数量減少で現物を自動移動しない。
- [ ] 取り置き登録の可視ヘッダーから旧内部8列が消え、必要情報と商品別4数量が揃う。
- [ ] メモ・棚確認・現物確認は入力保存または台帳に残り、洗い替えで消えない。
- [ ] 5分類の各注文は排他的で、引当状況一覧には全注文が出る。
- [ ] GoQ差分は表示だけで、GoQへ自動書き込みしない。
- [ ] 自動色分けは先行=薄青、到着済=ラベンダー、現物=緑を毎回再適用する。
- [ ] 先行は納品書・ピック・⑤未ピック・Yahoo出力・EMS在庫移動に入らず、対象便に先行残があれば⑤を停止する。
- [ ] 1入力キーのエラーで他キーを破棄せず、未適用入力・理由・メモを非表示シートへ保持する。
- [ ] データ行ループ内にSpreadsheetサービスのRange呼出を残さない。
- [ ] 移行・全件反映はバックアップと署名再検証の後だけ書き込む。
- [ ] 台湾・中国・ダニエル便の処理に回帰がない。
