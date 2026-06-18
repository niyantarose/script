# 大邱データ転送ボタン Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 韓国側手動入力シート（発注リスト大邱データ / EMS大邱作業データ）の新規行を、発注 / EMSリストへボタンで追記し、EMS転送後は購入Noを自動補完する。

**Architecture:** 既存の `発注_チェック行をEMSリストへ送る` と同じパターン（最終データ行の下に追記 / 重複はSetで判定 / 既存データ・数式は不変）。新規ファイル `Project_19/【大邱】データ転送.js` に転送関数を集約し、「手動運用」メニューにボタンを追加。純粋ヘルパーはセルフテスト関数で検証、転送本体はプレビュー確認＋手動検証。

**Tech Stack:** Google Apps Script（clasp管理、Project_19、scriptId `17N3ueJKAdvmsjAohaQzQCHUitW9RxHZpLlAm2zrzDJzLV811d8cBuHqT`）。既存ヘルパー流用：`normCode_` / `codeKeys_` / `normTrack_`（エクセルからデータ取得.js）、`H2E_findNextAppendRow_` / `H2E_getNextEmsNo_`（【発注リスト】Ｂ列…js）、`EMSリスト_購入No自動補完`（【EMSリスト】購入番号を自動取得.js）。すべて同一GASプロジェクト＝グローバル参照可。

**前提（spec）:** `docs/superpowers/specs/2026-06-17-daegu-transfer-buttons-design.md`

**デプロイ/コミット方針:** 各タスクの反映は `cd Project_19 && clasp push --force`。git commit はユーザー承認時のみ（標準方針）。push前に temp へ `clasp pull` 差分確認を推奨。

**列対応リファレンス（0始まり=getValues配列 / 1始まり=getRange列）**

発注リスト大邱データ（getValues 0始まり）: C入荷日=2, F発注NO=5, H業者=7, I商品名=8, Jオプション=9, K商品コード=10, L数量=11, N品目=13, O重さ=14, P価格=15, T決済方法=19, U決済日=20
発注（getRange 1始まり）: D入荷日=4, G購入No=7, I業者=9, J商品名=10, Kオプション=11, L商品コード=12, M発注数量=13, O品目=15, P重さ=16, Q価格=17, U決済方法=21, V決済日=22, Y EMS発送数=25
EMS大邱作業データ（getValues 0始まり）: A入荷日=0, B発送日=1, D EMS番号=3, H商品コード=7, I数量=8, K品目=10
EMSリスト（getRange 1始まり）: A No=1, B入荷日=2, C EMS発送日=3, F購入No=6, Gステータス=7, I商品コード=9, J数量=10, K品目=11, M EMS番号=13

---

### Task 1: 新規ファイル＋純粋ヘルパー＋セルフテスト

**Files:**
- Create: `Project_19/【大邱】データ転送.js`

- [ ] **Step 1: ファイルを作成し、設定・短縮コードヘルパー・セルフテストを書く**

```javascript
// ============================================================
// 大邱（韓国側手動入力シート）→ 発注 / EMSリスト 転送
//   依存: normCode_ / codeKeys_ / normTrack_（エクセルからデータ取得.js）
//         H2E_findNextAppendRow_ / H2E_getNextEmsNo_（【発注リスト】Ｂ列…js）
//         EMSリスト_購入No自動補完（【EMSリスト】購入番号を自動取得.js）
// ============================================================

const DAEGU_CFG = {
  HACHU_SRC: '発注リスト大邱データ',
  HACHU_DST: '発注',
  HACHU_START_ROW: 7,
  EMS_SRC: 'EMS大邱作業データ',
  EMS_DST: 'EMSリスト',
};

/**
 * 商品コードを発注/EMSリストの表記へ正規化する。
 * 例) KRSJCM03-0506_06 -> KRSJCM03-06、KRSJCM03-0506S -> KRSJCM03-0506S
 * codeKeys_ の最後の要素（短縮形があれば短縮形、無ければ正規化フル）を採用。
 */
function 大邱_短縮コード_(raw) {
  const keys = codeKeys_(raw);
  return keys[keys.length - 1];
}

/**
 * 純粋ロジックのセルフテスト（メニュー不要、エディタから実行してLogを確認）。
 */
function 大邱_セルフテスト_() {
  const cases = [
    ['KRSJCM03-0506_06', 'KRSJCM03-06'],
    ['KRSJCM03-0506-05', 'KRSJCM03-05'],
    ['KRSJCM03-0506S', 'KRSJCM03-0506S'],
    ['MRBLUE40_3', 'MRBLUE40-3'],
  ];
  let ok = 0, ng = 0;
  cases.forEach(([inp, exp]) => {
    const got = 大邱_短縮コード_(inp);
    if (got === exp) { ok++; Logger.log('OK  ' + inp + ' -> ' + got); }
    else { ng++; Logger.log('NG  ' + inp + ' -> ' + got + ' (期待 ' + exp + ')'); }
  });
  Logger.log('大邱_セルフテスト_: OK=' + ok + ' NG=' + ng);
  return ng === 0;
}
```

- [ ] **Step 2: push**

Run: `cd Project_19 && clasp push --force`
Expected: `Pushed N files`（exit 0）

- [ ] **Step 3: セルフテストを実行して検証**

Apps Scriptエディタで `大邱_セルフテスト_` を実行 → 実行ログを確認。
Expected: `大邱_セルフテスト_: OK=4 NG=0`（全て OK 行）

---

### Task 2: 機能1 「発注リスト大邱データ → 発注」

**Files:**
- Modify: `Project_19/【大邱】データ転送.js`（末尾に追記）

- [ ] **Step 1: 発注の重複キーセットと追記行ヘルパーを追記**

```javascript
/** 発注の既存行から重複キー（購入No|正規化コード）の Set を作る */
function 発注_既存キーセット_() {
  const dst = SpreadsheetApp.getActive().getSheetByName(DAEGU_CFG.HACHU_DST);
  const startRow = DAEGU_CFG.HACHU_START_ROW;
  const set = new Set();
  if (!dst) return set;
  const lastRow = dst.getLastRow();
  if (lastRow < startRow) return set;
  const n = lastRow - startRow + 1;
  const g = dst.getRange(startRow, 7, n, 1).getDisplayValues();   // G 購入No
  const l = dst.getRange(startRow, 12, n, 1).getDisplayValues();  // L 商品コード
  for (let i = 0; i < n; i++) {
    const no = String(g[i][0] || '').trim();
    const code = String(l[i][0] || '').trim();
    if (no && code) set.add(no + '|' + normCode_(code));
  }
  return set;
}

/** 発注の最終データ行の次の行（G/L/M列の実データで判定。数式スピル列は見ない） */
function 発注_次の追記行_() {
  const dst = SpreadsheetApp.getActive().getSheetByName(DAEGU_CFG.HACHU_DST);
  const startRow = DAEGU_CFG.HACHU_START_ROW;
  const maxRows = dst.getMaxRows();
  const cols = [7, 12, 13]; // G購入No, L商品コード, M発注数量
  let last = startRow - 1;
  cols.forEach(c => {
    const vals = dst.getRange(startRow, c, maxRows - startRow + 1, 1).getDisplayValues();
    for (let i = vals.length - 1; i >= 0; i--) {
      if (String(vals[i][0] || '').trim() !== '') {
        const rn = startRow + i;
        if (rn > last) last = rn;
        break;
      }
    }
  });
  return last + 1;
}
```

- [ ] **Step 2: 転送本体を追記**

```javascript
function 大邱_発注へ転送() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const src = ss.getSheetByName(DAEGU_CFG.HACHU_SRC);
  const dst = ss.getSheetByName(DAEGU_CFG.HACHU_DST);
  if (!src || !dst) { ui.alert('「' + DAEGU_CFG.HACHU_SRC + '」か「' + DAEGU_CFG.HACHU_DST + '」が見つかりません。'); return; }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(2000)) { ss.toast('他の処理が実行中です。'); return; }
  try {
    const sVals = src.getDataRange().getValues();
    const existing = 発注_既存キーセット_();

    const appendRows = [];
    for (let i = 0; i < sVals.length; i++) {
      const r = sVals[i];
      const orderNo = String(r[5] || '').trim();   // F 発注NO
      const code = String(r[10] || '').trim();      // K 商品コード
      if (!orderNo || !code) continue;
      if (/発注NO|OrderNo/i.test(orderNo)) continue; // ヘッダー除外
      const key = orderNo + '|' + normCode_(code);
      if (existing.has(key)) continue;
      existing.add(key);
      appendRows.push(r);
    }

    if (appendRows.length === 0) { ss.toast('追記対象なし（すべて既存）。'); return; }

    const preview = appendRows.slice(0, 10)
      .map(r => `${r[5]} / ${r[10]} ×${r[11]}`).join('\n');
    const res = ui.alert('発注へ追記',
      `${appendRows.length}件を発注の最終行の下へ追記します。\n\n${preview}` +
      (appendRows.length > 10 ? '\n…ほか' : '') + '\n\n実行する？',
      ui.ButtonSet.YES_NO);
    if (res !== ui.Button.YES) { ui.alert('やめました。'); return; }

    const startRow = 発注_次の追記行_();
    const n = appendRows.length;

    // [大邱index, 発注列(1始まり)]
    const MAP = [
      [2, 4],   // C入荷日 -> D入荷日
      [5, 7],   // F発注NO -> G購入No
      [7, 9],   // H業者   -> I業者
      [8, 10],  // I商品名 -> J商品名
      [9, 11],  // Jオプション -> Kオプション
      [10, 12], // K商品コード -> L商品コード
      [11, 13], // L数量   -> M発注数量
      [13, 15], // N品目   -> O品目
      [14, 16], // O重さ   -> P重さ
      [15, 17], // P価格   -> Q価格
      [19, 21], // T決済方法 -> U決済方法
      [20, 22], // U決済日   -> V決済日
    ];
    MAP.forEach(([si, dc]) => {
      const col = appendRows.map(r => [r[si]]);
      dst.getRange(startRow, dc, n, 1).setValues(col);
    });

    // Y列(25) EMS発送数 のSUMIFSを追記行ぶん設定（既存の一括修正と同型）
    const yF = appendRows.map((_, k) => {
      const row = startRow + k;
      return [`=SUMIFS('EMSリスト'!$J$7:$J,'EMSリスト'!$F$7:$F,$G${row},'EMSリスト'!$I$7:$I,$L${row})`];
    });
    dst.getRange(startRow, 25, n, 1).setFormulas(yF);

    ss.toast(`発注へ追記: ${n}件`);
  } finally {
    lock.releaseLock();
  }
}
```

- [ ] **Step 3: push**

Run: `cd Project_19 && clasp push --force`
Expected: `Pushed N files`（exit 0）

- [ ] **Step 4: 手動検証（ブラウザ＋gviz）**

1. 「発注リスト大邱データ」末尾に、発注に無いことが明らかな新規1行（適当なテスト発注NO・商品コード）を手で追加。
2. メニュー「手動運用 → 発注リスト大邱データ → 発注」を実行 → プレビューに「1件」と出る → YES。
3. gvizで「発注」の最終データ行を読み、購入No(G)/商品コード(L)/数量(M)/入荷日(D)/業者(I) が入り、Y列(25)に `=SUMIFS(...)` が入っていることを確認。
4. もう一度同じボタンを実行 → 「追記対象なし（すべて既存）」になることを確認（重複防止）。
5. テスト行は発注・大邱の両方から削除して原状復帰。
Expected: 1件だけ最終行下に正しく追記、Y列に数式、再実行で重複スキップ。

---

### Task 3: 機能2 「EMS大邱作業データ → EMSリスト」＋補完自動実行

**Files:**
- Modify: `Project_19/【大邱】データ転送.js`（末尾に追記）

- [ ] **Step 1: EMSリストの重複キーセットヘルパーを追記**

```javascript
/** EMSリストの既存行から重複キー（normTrack|normCode|数量）の Set を作る */
function EMS_既存キーセット_() {
  const dst = SpreadsheetApp.getActive().getSheetByName(DAEGU_CFG.EMS_DST);
  const set = new Set();
  if (!dst) return set;
  const vals = dst.getDataRange().getValues();
  let h = -1;
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]).trim() === 'No.') { h = i; break; }
  if (h < 0) return set;
  for (let i = h + 1; i < vals.length; i++) {
    const track = normTrack_(vals[i][12]); // M EMS番号
    const code = normCode_(vals[i][8]);    // I 商品コード
    const qty = String(vals[i][9] || '').trim(); // J 数量
    if (track || code) set.add(track + '|' + code + '|' + qty);
  }
  return set;
}
```

- [ ] **Step 2: 転送本体を追記（D列EMS番号入りのみ・追記後に補完自動実行）**

```javascript
function 大邱_EMSリストへ転送() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const src = ss.getSheetByName(DAEGU_CFG.EMS_SRC);
  const dst = ss.getSheetByName(DAEGU_CFG.EMS_DST);
  if (!src || !dst) { ui.alert('「' + DAEGU_CFG.EMS_SRC + '」か「' + DAEGU_CFG.EMS_DST + '」が見つかりません。'); return; }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(2000)) { ss.toast('他の処理が実行中です。'); return; }
  try {
    const sVals = src.getDataRange().getValues();
    const existing = EMS_既存キーセット_();

    const rows = [];
    for (let i = 0; i < sVals.length; i++) {
      const r = sVals[i];
      const track = String(r[3] || '').trim();   // D EMS番号（これが無い行は対象外）
      const codeRaw = String(r[7] || '').trim();  // H 商品コード
      if (!track || !codeRaw) continue;
      if (/Tracking|追跡|tracking #/i.test(track)) continue; // ヘッダー除外
      const code = 大邱_短縮コード_(codeRaw);
      const qty = String(r[8] || '').trim();       // I 数量
      const key = normTrack_(track) + '|' + normCode_(code) + '|' + qty;
      if (existing.has(key)) continue;
      existing.add(key);
      rows.push({ arrival: r[0], ship: r[1], track: track, code: code, qty: r[8], item: r[10] });
    }

    if (rows.length === 0) { ss.toast('追記対象なし（D列EMS番号入りで未登録の行なし）。'); return; }

    const preview = rows.slice(0, 10)
      .map(o => `${o.track} / ${o.code} ×${o.qty}`).join('\n');
    const res = ui.alert('EMSリストへ追記',
      `${rows.length}件をEMSリストの最終行の下へ追記し、購入Noを補完します。\n\n${preview}` +
      (rows.length > 10 ? '\n…ほか' : '') + '\n\n実行する？',
      ui.ButtonSet.YES_NO);
    if (res !== ui.Button.YES) { ui.alert('やめました。'); return; }

    const startRow = H2E_findNextAppendRow_(dst); // EMSリスト用の追記行（既存ヘルパー流用）
    let no = H2E_getNextEmsNo_(dst);
    const n = rows.length;

    // 列ごとに setValues
    dst.getRange(startRow, 1, n, 1).setValues(rows.map(() => [no++]));        // A No.
    dst.getRange(startRow, 2, n, 1).setValues(rows.map(o => [o.arrival]));    // B 入荷日
    dst.getRange(startRow, 3, n, 1).setValues(rows.map(o => [o.ship]));       // C EMS発送日
    dst.getRange(startRow, 7, n, 1).setValues(rows.map(() => ['未着']));      // G ステータス
    dst.getRange(startRow, 9, n, 1).setValues(rows.map(o => [o.code]));       // I 商品コード
    dst.getRange(startRow, 10, n, 1).setValues(rows.map(o => [o.qty]));       // J 数量
    dst.getRange(startRow, 11, n, 1).setValues(rows.map(o => [o.item]));      // K 品目
    dst.getRange(startRow, 13, n, 1).setValues(rows.map(o => [o.track]));     // M EMS番号
    // F 購入No は空のまま（次で補完）

    SpreadsheetApp.flush();
    let filled = 0;
    if (typeof EMSリスト_購入No自動補完 === 'function') {
      filled = EMSリスト_購入No自動補完(true) || 0; // silent
    }
    ss.toast(`EMSリストへ追記: ${n}件 / 購入No補完: ${filled}件`);
  } finally {
    lock.releaseLock();
  }
}
```

- [ ] **Step 3: push**

Run: `cd Project_19 && clasp push --force`
Expected: `Pushed N files`（exit 0）

- [ ] **Step 4: 手動検証（ブラウザ＋gviz）**

1. 「EMS大邱作業データ」末尾に、D列(EMS番号)を入れたテスト新規1行（既存の発注にある商品コード・購入Noで照合できるもの）を手で追加。D列が空の行も1行用意。
2. メニュー「手動運用 → EMS大邱作業データ → EMSリスト」を実行 → プレビューに「1件」（D列空の行は対象外）→ YES。
3. gvizで「EMSリスト」最終行を読み、入荷日(B)/発送日(C)/EMS番号(M)/商品コード(I=短縮形)/数量(J)/品目(K)/ステータス(G='未着') が入り、購入No(F)が補完で埋まっていることを確認。
4. 分割発送の確認：同じ入荷日・商品コード・数量だがEMS番号が違うテスト2行 → 両方が別行として残ることを確認（重複扱いされない）。
5. もう一度ボタン実行 → 「追記対象なし」を確認。
6. テスト行はEMSリスト・大邱の両方から削除して原状復帰。
Expected: D列EMS番号入りのみ追記、コードは短縮形、購入No補完が走る、別EMS番号は別行で残る、再実行で重複スキップ。

---

### Task 4: 「手動運用」メニューにボタン追加

**Files:**
- Modify: `Project_19/エクセルからデータ取得.js:5`（`Excel同期メニューを追加_` の `.createMenu('手動運用')` 直後）

- [ ] **Step 1: メニュー項目を追加**

`エクセルからデータ取得.js` の該当箇所を次のように変更（4〜5行目）。

変更前:
```javascript
    .createMenu('手動運用')
    .addItem('EMSカレンダーシートを更新', 'buildEmsCalendarSheet')
```
変更後:
```javascript
    .createMenu('手動運用')
    .addItem('発注リスト大邱データ → 発注', '大邱_発注へ転送')
    .addItem('EMS大邱作業データ → EMSリスト', '大邱_EMSリストへ転送')
    .addItem('購入No補完（EMSリスト）', 'EMSリスト_購入No自動補完')
    .addItem('EMS発送数の数式を一括修正（発注）', '発注_EMS発送数数式を一括修正')
    .addSeparator()
    .addItem('EMSカレンダーシートを更新', 'buildEmsCalendarSheet')
```

- [ ] **Step 2: push**

Run: `cd Project_19 && clasp push --force`
Expected: `Pushed N files`（exit 0）

- [ ] **Step 3: メニュー表示を検証**

スプレッドシートを再読み込み（onOpen再実行）→「手動運用」メニューを開く。
Expected: 先頭に「発注リスト大邱データ → 発注」「EMS大邱作業データ → EMSリスト」「購入No補完（EMSリスト）」「EMS発送数の数式を一括修正（発注）」の4項目＋区切り線が表示される。各項目を一度クリックして関数が起動する（プレビューやトーストが出る）ことを確認。

---

## 検証まとめ（全タスク後）

- 既存の発注 / EMSリストの**既存行と数式（小計R・支払金額T・照合キーZ・消込判定AA）が一切変化していない**ことを gviz / 目視で確認。
- 推奨運用順：先に「発注リスト大邱データ → 発注」、後に「EMS大邱作業データ → EMSリスト」。
