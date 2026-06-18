// ============================================================
// 大邱（韓国側手動入力シート）の 発注 / EMSリスト転送
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
 * 商品コードを発注/EMSリストの表記へ正規化する
 * 例: KRSJCM03-0506_06 -> KRSJCM03-06、KRSJCM03-0506S -> KRSJCM03-0506S
 * codeKeys_ の最後の要素（短縮形があれば短縮形、無ければ正規化フル）を採用
 */
function 大邱_短縮コード_(raw) {
  const keys = codeKeys_(raw);
  return keys[keys.length - 1];
}

/**
 * 純粋ロジックのセルフテスト（メニュー不要、エディタから実行してLogを確認）
 */
function 大邱_セルフテスト() {
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
    else { ng++; Logger.log('NG  ' + inp + ' -> ' + got + ' (期待: ' + exp + ')'); }
  });
  Logger.log('大邱_セルフテスト: OK=' + ok + ' NG=' + ng);
  return ng === 0;
}

// ============================================================
// 機能1: 発注リスト大邱データ ➡ 発注
// ============================================================

/** 
 * 発注の既存行から重複キー（購入No|正規化コード|商品名|オプション）の Set を作る 
 * ※同じコードで商品名やオプションが異なるもの（MDなど）を正しく別行で登録できるよう、商品名とオプションもキーに含めます。
 */
function 発注_既存キーセット_() {
  const dst = SpreadsheetApp.getActive().getSheetByName(DAEGU_CFG.HACHU_DST);
  const startRow = DAEGU_CFG.HACHU_START_ROW;
  const set = new Set();
  if (!dst) return set;
  const lastRow = dst.getLastRow();
  if (lastRow < startRow) return set;
  const n = lastRow - startRow + 1;
  const g = dst.getRange(startRow, 7, n, 1).getDisplayValues();   // G 購入No
  const j = dst.getRange(startRow, 10, n, 1).getDisplayValues();  // J 商品名
  const k = dst.getRange(startRow, 11, n, 1).getDisplayValues();  // K オプション
  const l = dst.getRange(startRow, 12, n, 1).getDisplayValues();  // L 商品コード
  for (let i = 0; i < n; i++) {
    const no = String(g[i][0] || '').trim();
    const name = String(j[i][0] || '').trim();
    const option = String(k[i][0] || '').trim();
    const code = String(l[i][0] || '').trim();
    if (no && code) set.add(no + '|' + normCode_(code) + '|' + name + '|' + option);
  }
  return set;
}

/** 発注の最終データ行の次の行！G列/L列/M列の実データで判定。数式スピル列は見ない！ */
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
      const name = String(r[8] || '').trim();      // I 商品名
      const option = String(r[9] || '').trim();    // J オプション
      const code = String(r[10] || '').trim();     // K 商品コード
      if (!orderNo || !code) continue;
      if (/発注NO|OrderNo/i.test(orderNo)) continue; // ヘッダー除外
      const key = orderNo + '|' + normCode_(code) + '|' + name + '|' + option;
      if (existing.has(key)) continue;
      existing.add(key);
      appendRows.push(r);
    }

    if (appendRows.length === 0) { ss.toast('追記対象なし（すべて既存）。'); return; }

    const preview = appendRows.slice(0, 10)
      .map(r => `${r[5]} / ${r[10]} / ${r[8]}`).join('\n');
    const res = ui.alert('発注へ追記',
      `${appendRows.length}件を発注の最終行の下へ追記します。\n${preview}` +
      (appendRows.length > 10 ? '\n…ほか' : '') + '\n\n実行する！',
      ui.ButtonSet.YES_NO);
    if (res !== ui.Button.YES) { ui.alert('やめました。'); return; }

    const startRow = 発注_次の追記行_();
    const n = appendRows.length;

    // [大邱index(0始まり), 発注列(1始まり)]
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

    ss.toast(`発注へ追記 ${n}件`);
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 機能2: EMS大邱作業データ ➡ EMSリスト（追記後に購入No補完を自動実行！）
// ============================================================

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
    const debugMismatches = [];
    for (let i = 0; i < sVals.length; i++) {
      const r = sVals[i];
      const track = String(r[3] || '').trim();   // D EMS番号（これが無い行は対象外！）
      const codeRaw = String(r[7] || '').trim();  // H 商品コード
      if (!track || !codeRaw) continue;
      if (/Tracking|追跡|tracking #/i.test(track)) continue; // ヘッダー除外
      const code = 大邱_短縮コード_(codeRaw);
      const qty = String(r[8] || '').trim();       // I 数量
      const key = normTrack_(track) + '|' + normCode_(code) + '|' + qty;

      // 追跡と商品コードが一致するが全体（数量など）が異なるペアを調査
      existing.forEach(k => {
        const pk = k.split('|');
        if (pk[0] === normTrack_(track) && pk[1] === normCode_(code)) {
          if (k !== key) {
            debugMismatches.push(`既存: [${k}] vs 大邱: [${key}] (文字数 ${k.length} vs ${key.length})`);
          }
        }
      });

      if (existing.has(key)) continue;
      existing.add(key);
      rows.push({ arrival: r[0], ship: r[1], track: track, code: code, qty: r[8], item: r[10] });
    }

    if (rows.length === 0) { ss.toast('追記対象なし（D列EMS番号入りで未登録の行なし）。'); return; }

    let debugMsg = "";
    const debugTracks = ['ES396936624KR', 'ES396936638KR', 'EG048695960KR'];
    debugTracks.forEach(t => {
      const matchesInExisting = [];
      existing.forEach(k => {
        if (k.startsWith(t)) matchesInExisting.push(k);
      });
      debugMsg += `\n- ${t}: ` + (matchesInExisting.length ? matchesInExisting.join(', ') : 'なし');
    });

    if (debugMismatches.length > 0) {
      debugMsg += "\n\n【不一致ペアの分析】\n" + debugMismatches.slice(0, 10).join('\n');
    }

    const preview = rows.slice(0, 10)
      .map(o => `${o.track} / ${o.code} / ${o.qty}`).join('\n');
    const res = ui.alert('EMSリストへ追記 (デバッグ情報付き)',
      `${rows.length}件をEMSリストの最終行の下へ追記し、購入Noを補完します。\n\n【プレビュー】\n${preview}` +
      (rows.length > 10 ? '\n…ほか' : '') + `\n\n【デバッグ照合情報】\n(EMSリスト内の登録状況)${debugMsg}\n\n実行する！`,
      ui.ButtonSet.YES_NO);
    if (res !== ui.Button.YES) { ui.alert('やめました。'); return; }

    const startRow = H2E_findNextAppendRow_(dst); // EMSリスト用の追記行（既存ヘルパー流用）
    let no = H2E_getNextEmsNo_(dst);
    const n = rows.length;

    dst.getRange(startRow, 1, n, 1).setValues(rows.map(() => [no++]));        // A No.
    dst.getRange(startRow, 2, n, 1).setValues(rows.map(o => [o.arrival]));    // B 入荷日
    dst.getRange(startRow, 3, n, 1).setValues(rows.map(o => [o.ship]));       // C EMS発送日
    dst.getRange(startRow, 7, n, 1).setValues(rows.map(() => ['未着']));      // G ステータス
    dst.getRange(startRow, 9, n, 1).setValues(rows.map(o => [o.code]));       // I 商品コード
    dst.getRange(startRow, 10, n, 1).setValues(rows.map(o => [o.qty]));       // J 数量
    dst.getRange(startRow, 11, n, 1).setValues(rows.map(o => [o.item]));      // K 品目
    dst.getRange(startRow, 13, n, 1).setValues(rows.map(o => [o.track]));     // M EMS番号
    // F 購入No は空のまま（次で補完！）

    SpreadsheetApp.flush();
    let filled = 0;
    if (typeof EMSリスト_購入No自動補完 === 'function') {
      filled = EMSリスト_購入No自動補完(true) || 0; // silent
    }
    ss.toast(`EMSリストへ追記 ${n}件 / 購入No補完 ${filled}件`);
  } finally {
    lock.releaseLock();
  }
}