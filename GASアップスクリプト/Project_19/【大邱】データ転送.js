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
 * 商品コードを照合用に短縮する（表示用では使わない）
 * 例: KRSJCM03-0506_06 -> KRSJCM03-06、KRSJCM03-0506S -> KRSJCM03-0506S
 */
function 大邱_短縮コード_(raw) {
  const c = normCode_(raw);
  const m = c.match(/^(.+)-(\d{2})(\d{2})-(\d{2})$/);
  if (m && (m[4] === m[2] || m[4] === m[3])) return m[1] + '-' + m[4];
  return c;
}

function 大邱_表示コード_(raw) {
  return String(raw || '')
    .normalize('NFKC')
    .replace(/[‐‑‒–—―−]/g, '-')
    .replace(/[＿]/g, '_')
    .replace(/[\s\u3000]+/g, '')
    .trim();
}

function EMS_表示追跡番号_(value) {
  return (typeof normTrack_ === 'function')
    ? normTrack_(value)
    : String(value || '').normalize('NFKC').toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function EMS_転送数量キー_(value) {
  return String(value || '')
    .normalize('NFKC')
    .replace(/,/g, '')
    .replace(/[\s\u3000]+/g, '')
    .trim();
}

function EMS_表示数量_(value) {
  const q = EMS_転送数量キー_(value);
  return q || '';
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
 * 発注の既存行から「購入No(G列)」の Set を作る。
 * ※購入Noは一意（発注NO＋連番）なので、完全一致だけで重複判定できる。
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
  for (let i = 0; i < n; i++) {
    const no = String(g[i][0] || '').trim();
    if (no) set.add(no);
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
      const orderNo = EMS_転送購入Noキー_(r[5]);    // F 発注NO
      const code = 大邱_表示コード_(r[10]);          // K 商品コード
      if (!orderNo || !code) continue;
      if (/発注NO|OrderNo/i.test(orderNo)) continue; // ヘッダー除外
      const key = orderNo;
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

    dst.getRange(startRow, 7, n, 1).setValues(appendRows.map(r => [EMS_転送購入Noキー_(r[5])]));

    // 商品コード(L列)は大邱データの表記を保つ。照合は codeKeys_ 側で表記ゆれを吸収する。
    dst.getRange(startRow, 12, n, 1).setValues(appendRows.map(r => [大邱_表示コード_(r[10])]));

    // Y列(EMS発送数)は発注Y7の1本のMAP式が自動展開するので、ここでは書かない。
    // （万一Y7にMAP式が無い場合は「発注：EMS発送数の式を一括修正」を一度実行）

    ss.toast(`発注へ追記 ${n}件`);
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 機能2: EMS大邱作業データ ➡ EMSリスト（追記後に購入No補完を自動実行！）
// ============================================================

function EMS_転送購入Noキー_(value) {
  return String(value || '')
    .trim()
    .replace(/[＿]/g, '_')
    .replace(/[\s\u3000]+/g, '')
    .replace(/_+/g, '_');
}

function EMS_転送コードキー一覧_(code) {
  const keys = (typeof codeKeys_ === 'function') ? codeKeys_(code) : [normCode_(code)];
  return keys.filter(k => k);
}

function EMS_転送キー一覧_(purchaseNo, track, code, qty) {
  const no = EMS_転送購入Noキー_(purchaseNo);
  const tr = EMS_表示追跡番号_(track);
  const q = EMS_転送数量キー_(qty);
  return EMS_転送コードキー一覧_(code).map(c => [no, tr, c, q].join('|'));
}

function EMS_転送三点キー一覧_(track, code, qty) {
  const tr = EMS_表示追跡番号_(track);
  const q = EMS_転送数量キー_(qty);
  return EMS_転送コードキー一覧_(code).map(c => [tr, c, q].join('|'));
}

function EMS_復元数量違いキー_(purchaseNo, track, code) {
  const no = EMS_転送購入Noキー_(purchaseNo);
  const tr = EMS_表示追跡番号_(track);
  const c = normCode_(code);
  if (!no || !tr || !c) return '';
  return [no, tr, c].join('|');
}

function EMS_復元EMS違いキー_(purchaseNo, code, qty) {
  const no = EMS_転送購入Noキー_(purchaseNo);
  const c = normCode_(code);
  const q = EMS_転送数量キー_(qty);
  if (!no || !c || !q) return '';
  return [no, c, q].join('|');
}

function EMS_キー一覧を追加_(map, keys, item) {
  keys.forEach(key => 枝番_map追加_(map, key, item));
}

function EMS_キー一覧に含む_(set, keys) {
  return keys.some(key => set.has(key));
}

/** EMSリストの既存行から重複判定情報を作る */
function EMS_既存キー情報_(sourceExactKeys) {
  const dst = SpreadsheetApp.getActive().getSheetByName(DAEGU_CFG.EMS_DST);
  const info = { exact: {}, loose: {} };
  if (!dst) return info;
  const vals = dst.getDataRange().getValues();
  let h = -1;
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]).trim() === 'No.') { h = i; break; }
  if (h < 0) return info;
  for (let i = h + 1; i < vals.length; i++) {
    const no = EMS_転送購入Noキー_(vals[i][5]); // F 購入No
    const track = EMS_表示追跡番号_(vals[i][12]); // M EMS番号
    const code = vals[i][8];               // I 商品コード
    const qty = EMS_転送数量キー_(vals[i][9]); // J 数量
    if (!track || !code) continue;
    const exactKeys = EMS_転送キー一覧_(no, track, code, qty);
    const looseKeys = EMS_転送三点キー一覧_(track, code, qty);
    const item = {
      used: false,
      row: i + 1,
      code: 大邱_表示コード_(code),
      qty: qty,
      track: track,
      purchaseNo: no,
      arrival: vals[i][1], // B 入荷日
      ship: vals[i][2]     // C EMS発送日
    };
    if (no) EMS_キー一覧を追加_(info.exact, exactKeys, item);

    // 旧データや購入No未補完の既存行は、EMS番号+商品コード+数量の件数として消費する。
    // 今回の大邱データに購入Noまで一致する既存行は exact 側で判定する。
    if (!no || !sourceExactKeys || !EMS_キー一覧に含む_(sourceExactKeys, exactKeys)) {
      EMS_キー一覧を追加_(info.loose, looseKeys, item);
    }
  }
  return info;
}

function EMS_日付キー_(value) {
  const d = (typeof _emsDateOnly_ === 'function') ? _emsDateOnly_(value) : null;
  if (d instanceof Date && !isNaN(d.getTime())) {
    return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  return String(value || '').normalize('NFKC').trim();
}

function EMS_値あり_(value) {
  return String(value == null ? '' : value).normalize('NFKC').trim() !== '';
}

function EMS_既存行日付更新候補_(updates, existingRow, srcRow) {
  if (!existingRow || !existingRow.row || !srcRow) return;
  if (EMS_値あり_(srcRow.arrival) && EMS_日付キー_(existingRow.arrival) !== EMS_日付キー_(srcRow.arrival)) {
    updates.push({ row: existingRow.row, col: 2, value: srcRow.arrival, label: '入荷日' });
  }
  if (EMS_値あり_(srcRow.ship) && EMS_日付キー_(existingRow.ship) !== EMS_日付キー_(srcRow.ship)) {
    updates.push({ row: existingRow.row, col: 3, value: srcRow.ship, label: 'EMS発送日', isShip: true });
  }
}

function EMS_日付更新を適用_(dst, updates) {
  if (!updates || updates.length === 0) return 0;
  const shipRows = {};
  updates.forEach(update => {
    dst.getRange(update.row, update.col).setValue(update.value);
    if (update.isShip) {
      dst.getRange(update.row, 4).setValue('⇒'); // D 矢印
      shipRows[update.row] = true;
    }
  });
  if (typeof EMS_fillArrivalEstimateRows_ === 'function') {
    Object.keys(shipRows).forEach(row => EMS_fillArrivalEstimateRows_(dst, Number(row), 1, false));
  }
  return updates.length;
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
    const candidates = [];
    const sourceExactKeys = new Set();
    for (let i = 0; i < sVals.length; i++) {
      const r = sVals[i];
      const track = EMS_表示追跡番号_(r[3]);      // D EMS番号（空白や記号は吸収）
      const codeRaw = 大邱_表示コード_(r[7]);      // H 商品コード
      if (!track || !codeRaw) continue;
      if (/Tracking|追跡|tracking #/i.test(track)) continue; // ヘッダー除外
      const code = codeRaw;
      const qty = EMS_転送数量キー_(r[8]);         // I 数量
      const purchaseNo = EMS_転送購入Noキー_(r[19]); // T 購入No
      const exactKeys = EMS_転送キー一覧_(purchaseNo, track, code, qty);
      const looseKeys = EMS_転送三点キー一覧_(track, code, qty);
      const item = {
        purchaseNo: purchaseNo,
        arrival: r[0],
        ship: r[1],
        track: track,
        code: code,
        qty: EMS_表示数量_(r[8]),
        item: r[10],
        exactKeys: exactKeys,
        looseKeys: looseKeys
      };
      candidates.push(item);
      if (purchaseNo) exactKeys.forEach(key => sourceExactKeys.add(key));
    }

    const existing = EMS_既存キー情報_(sourceExactKeys);
    const rows = [];
    const dateUpdates = [];
    candidates.forEach(o => {
      const exactItem = o.purchaseNo ? 枝番_map取得_([existing.exact], [o.exactKeys]) : null;
      if (exactItem) {
        EMS_既存行日付更新候補_(dateUpdates, exactItem, o);
        return;
      }
      const looseItem = 枝番_map取得_([existing.loose], [o.looseKeys]);
      if (looseItem) {
        EMS_既存行日付更新候補_(dateUpdates, looseItem, o);
        return;
      }

      const newItem = { used: false };
      if (o.purchaseNo) {
        EMS_キー一覧を追加_(existing.exact, o.exactKeys, newItem);
      }
      EMS_キー一覧を追加_(existing.loose, o.looseKeys, newItem);
      rows.push(o);
    });

    if (rows.length === 0 && dateUpdates.length === 0) {
      ss.toast('追記/更新対象なし（D列EMS番号入りで未登録または日付更新が必要な行なし）。');
      return;
    }

    let debugMsg = "";
    const debugTracks = ['ES396936624KR', 'ES396936638KR', 'EG048695960KR'];
    debugTracks.forEach(t => {
      const matches = [];
      Object.keys(existing.loose).forEach(k => {
        const queue = existing.loose[k] || [];
        const remain = queue.filter(item => !item.used).length;
        if (k.indexOf(EMS_表示追跡番号_(t) + '|') === 0 && remain > 0) matches.push(k + ' x' + remain);
      });
      debugMsg += `\n- ${t}: ` + (matches.length ? matches.join(', ') : 'なし');
    });

    const preview = rows.slice(0, 10)
      .map(o => `${o.purchaseNo || '(購入No空)'} / ${o.track} / ${o.code} / ${o.qty}`).join('\n');
    const datePreview = dateUpdates.slice(0, 10)
      .map(o => `EMSリスト${o.row}行 ${o.label} -> ${EMS_日付キー_(o.value)}`)
      .join('\n');
    const res = ui.alert('EMSリストへ追記 (デバッグ情報付き)',
      `${rows.length}件をEMSリストの最終行の下へ追記し、既存行の日付 ${dateUpdates.length}件を更新します。\n\n【追記プレビュー】\n${preview || 'なし'}` +
      (rows.length > 10 ? '\n…ほか' : '') +
      `\n\n【日付更新】\n${datePreview || 'なし'}` +
      (dateUpdates.length > 10 ? '\n…ほか' : '') +
      `\n\n【デバッグ照合情報】\n(EMSリスト内の登録状況)${debugMsg}\n\n実行する！`,
      ui.ButtonSet.YES_NO);
    if (res !== ui.Button.YES) { ui.alert('やめました。'); return; }

    const updatedDates = EMS_日付更新を適用_(dst, dateUpdates);

    if (rows.length > 0) {
      const startRow = H2E_findNextAppendRow_(dst); // EMSリスト用の追記行（既存ヘルパー流用）
      let no = H2E_getNextEmsNo_(dst);
      const n = rows.length;

      dst.getRange(startRow, 1, n, 1).setValues(rows.map(() => [no++]));        // A No.
      dst.getRange(startRow, 2, n, 1).setValues(rows.map(o => [o.arrival]));    // B 入荷日
      dst.getRange(startRow, 3, n, 1).setValues(rows.map(o => [o.ship]));       // C EMS発送日
      if (typeof EMS_fillArrivalEstimateRows_ === 'function') {
        EMS_fillArrivalEstimateRows_(dst, startRow, n, false);                  // E 到着予定日
      }
      dst.getRange(startRow, 6, n, 1).setValues(rows.map(o => [o.purchaseNo])); // F 購入No
      dst.getRange(startRow, 7, n, 1).setValues(rows.map(() => ['未着']));      // G ステータス
      dst.getRange(startRow, 9, n, 1).setValues(rows.map(o => [o.code]));       // I 商品コード
      dst.getRange(startRow, 10, n, 1).setValues(rows.map(o => [o.qty]));       // J 数量
      dst.getRange(startRow, 11, n, 1).setValues(rows.map(o => [o.item]));      // K 品目
      dst.getRange(startRow, 13, n, 1).setValues(rows.map(o => [o.track]));     // M EMS番号
    }

    SpreadsheetApp.flush();
    // 発送日では並べ替えない（EMS大邱作業データと同じ追記順を保つ）
    let filled = 0;
    if (typeof EMSリスト_購入No自動補完 === 'function') {
      filled = EMSリスト_購入No自動補完(true) || 0; // silent
    }
    if (typeof 発注_EMS発送数数式を一括修正 === 'function') {
      発注_EMS発送数数式を一括修正();
    }
    ss.toast(`EMSリストへ追記 ${rows.length}件 / 日付更新 ${updatedDates}件 / 購入No補完 ${filled}件`);
  } finally {
    lock.releaseLock();
  }
}

function EMSリスト_大邱から不足行を確認して復元() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const src = ss.getSheetByName(DAEGU_CFG.EMS_SRC);
  const dst = ss.getSheetByName(DAEGU_CFG.EMS_DST);
  if (!src || !dst) { ui.alert('「' + DAEGU_CFG.EMS_SRC + '」か「' + DAEGU_CFG.EMS_DST + '」が見つかりません。'); return; }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(2000)) { ss.toast('他の処理が実行中です。'); return; }
  try {
    const sVals = src.getDataRange().getValues();
    const dVals = dst.getDataRange().getValues();
    const existing = EMS_既存キー情報_(null);
    const rows = [];
    const qtyUpdates = [];
    const trackUpdates = [];
    const codeUpdates = [];
    const dateUpdates = [];
    const extraRows = [];
    const qtyless = {};
    const trackless = {};
    const sourceTracklessCount = {};
    const sourceQtylessCount = {};
    const sourceRows = { exact: {}, loose: {} };

    for (let i = 0; i < sVals.length; i++) {
      const r = sVals[i];
      const track = EMS_表示追跡番号_(r[3]);
      const codeRaw = 大邱_表示コード_(r[7]);
      const qty = EMS_転送数量キー_(r[8]);
      const purchaseNo = EMS_転送購入Noキー_(r[19]);
      if (!track || !codeRaw || !qty) continue;
      if (/Tracking|Trakking|追跡/i.test(String(r[3] || ''))) continue;
      const key = EMS_復元EMS違いキー_(purchaseNo, codeRaw, qty);
      if (key) sourceTracklessCount[key] = (sourceTracklessCount[key] || 0) + 1;
      const qtylessKey = EMS_復元数量違いキー_(purchaseNo, track, codeRaw);
      if (qtylessKey) sourceQtylessCount[qtylessKey] = (sourceQtylessCount[qtylessKey] || 0) + 1;

      const item = { used: false };
      if (purchaseNo) EMS_キー一覧を追加_(sourceRows.exact, EMS_転送キー一覧_(purchaseNo, track, codeRaw, qty), item);
      EMS_キー一覧を追加_(sourceRows.loose, EMS_転送三点キー一覧_(track, codeRaw, qty), item);
    }

    let h = -1;
    for (let i = 0; i < dVals.length; i++) if (String(dVals[i][0]).trim() === 'No.') { h = i; break; }
    if (h >= 0) {
      for (let i = h + 1; i < dVals.length; i++) {
        const key = EMS_復元数量違いキー_(dVals[i][5], dVals[i][12], dVals[i][8]);
        if (!key) continue;
        (qtyless[key] = qtyless[key] || []).push({
          row: i + 1,
          qty: EMS_転送数量キー_(dVals[i][9]),
          used: false
        });

        const tKey = EMS_復元EMS違いキー_(dVals[i][5], dVals[i][8], dVals[i][9]);
        if (tKey) {
          (trackless[tKey] = trackless[tKey] || []).push({
            row: i + 1,
            track: EMS_表示追跡番号_(dVals[i][12]),
            used: false
          });
        }

        const no = EMS_転送購入Noキー_(dVals[i][5]);
        const track = EMS_表示追跡番号_(dVals[i][12]);
        const codeRaw = 大邱_表示コード_(dVals[i][8]);
        const qty = EMS_転送数量キー_(dVals[i][9]);
        if (track && codeRaw && qty) {
          const exactKeys = EMS_転送キー一覧_(no, track, codeRaw, qty);
          const looseKeys = EMS_転送三点キー一覧_(track, codeRaw, qty);
          const canFixTrack = tKey && sourceTracklessCount[tKey] > 0;
          const qKey = EMS_復元数量違いキー_(no, track, codeRaw);
          const canFixQty = qKey && sourceQtylessCount[qKey] > 0;
          const matched = (no && 枝番_map取得_([sourceRows.exact], [exactKeys])) ||
            枝番_map取得_([sourceRows.loose], [looseKeys]);
          if (!matched && !canFixTrack && !canFixQty) {
            extraRows.push({
              row: i + 1,
              purchaseNo: no,
              track: track,
              code: codeRaw,
              qty: EMS_表示数量_(dVals[i][9])
            });
          }
        }
      }
    }

    for (let i = 0; i < sVals.length; i++) {
      const r = sVals[i];
      const track = EMS_表示追跡番号_(r[3]);       // D EMS番号
      const codeRaw = 大邱_表示コード_(r[7]);       // H 商品コード
      if (!track || !codeRaw) continue;
      if (/Tracking|追跡|tracking #/i.test(track)) continue;

      const qty = EMS_転送数量キー_(r[8]);          // I 数量
      if (!qty) continue;
      const purchaseNo = EMS_転送購入Noキー_(r[19]); // T 購入No
      const exactKeys = EMS_転送キー一覧_(purchaseNo, track, codeRaw, qty);
      const looseKeys = EMS_転送三点キー一覧_(track, codeRaw, qty);

      const exactItem = purchaseNo ? 枝番_map取得_([existing.exact], [exactKeys]) : null;
      if (exactItem) {
        EMS_既存行日付更新候補_(dateUpdates, exactItem, { arrival: r[0], ship: r[1] });
        if (exactItem.code !== codeRaw) {
          codeUpdates.push({
            row: exactItem.row,
            srcRow: i + 1,
            purchaseNo: purchaseNo,
            track: track,
            oldCode: exactItem.code,
            code: codeRaw,
            qty: EMS_表示数量_(r[8])
          });
        }
        continue;
      }

      const tracklessKey = EMS_復元EMS違いキー_(purchaseNo, codeRaw, qty);
      const sameNoCodeQtyRows = tracklessKey ? (trackless[tracklessKey] || []) : [];
      const availableTrackRows = sameNoCodeQtyRows.filter(item => !item.used);
      if (sourceTracklessCount[tracklessKey] === 1 && availableTrackRows.length === 1 && availableTrackRows[0].track !== track) {
        availableTrackRows[0].used = true;
        const item = { used: false };
        EMS_キー一覧を追加_(existing.exact, exactKeys, item);
        EMS_キー一覧を追加_(existing.loose, looseKeys, item);
        trackUpdates.push({
          row: availableTrackRows[0].row,
          srcRow: i + 1,
          purchaseNo: purchaseNo,
          oldTrack: availableTrackRows[0].track,
          track: track,
          code: codeRaw,
          qty: EMS_表示数量_(r[8])
        });
        EMS_既存行日付更新候補_(dateUpdates, { row: availableTrackRows[0].row }, { arrival: r[0], ship: r[1] });
        continue;
      }

      const qtylessKey = EMS_復元数量違いキー_(purchaseNo, track, codeRaw);
      const sameRows = qtylessKey ? (qtyless[qtylessKey] || []) : [];
      const availableSameRows = sameRows.filter(item => !item.used);
      if (availableSameRows.length === 1 && availableSameRows[0].qty !== qty) {
        availableSameRows[0].used = true;
        const item = { used: false };
        EMS_キー一覧を追加_(existing.exact, exactKeys, item);
        EMS_キー一覧を追加_(existing.loose, looseKeys, item);
        qtyUpdates.push({
          row: availableSameRows[0].row,
          srcRow: i + 1,
          purchaseNo: purchaseNo,
          track: track,
          code: codeRaw,
          oldQty: availableSameRows[0].qty,
          qty: EMS_表示数量_(r[8])
        });
        EMS_既存行日付更新候補_(dateUpdates, { row: availableSameRows[0].row }, { arrival: r[0], ship: r[1] });
        continue;
      }

      const looseItem = 枝番_map取得_([existing.loose], [looseKeys]);
      if (looseItem) {
        EMS_既存行日付更新候補_(dateUpdates, looseItem, { arrival: r[0], ship: r[1] });
        if (looseItem.code !== codeRaw) {
          codeUpdates.push({
            row: looseItem.row,
            srcRow: i + 1,
            purchaseNo: purchaseNo,
            track: track,
            oldCode: looseItem.code,
            code: codeRaw,
            qty: EMS_表示数量_(r[8])
          });
        }
        continue;
      }

      const item = { used: false };
      if (purchaseNo) EMS_キー一覧を追加_(existing.exact, exactKeys, item);
      EMS_キー一覧を追加_(existing.loose, looseKeys, item);

      rows.push({
        srcRow: i + 1,
        purchaseNo: purchaseNo,
        arrival: r[0],
        ship: r[1],
        track: track,
        code: codeRaw,
        qty: EMS_表示数量_(r[8]),
        item: r[10]
      });
    }

    if (rows.length === 0 && qtyUpdates.length === 0 && trackUpdates.length === 0 && codeUpdates.length === 0 && dateUpdates.length === 0 && extraRows.length === 0) {
      ui.alert('EMS大邱作業データとEMSリストの行数・数量・EMS番号は一致しています。');
      return;
    }

    const preview = rows.slice(0, 20)
      .map(o => `大邱${o.srcRow}行: ${o.purchaseNo || '(購入No空)'} / ${o.track} / ${o.code} / ${o.qty}`)
      .join('\n');
    const updatePreview = qtyUpdates.slice(0, 20)
      .map(o => `EMSリスト${o.row}行: ${o.purchaseNo} / ${o.track} / ${o.code} / ${o.oldQty} -> ${o.qty}（大邱${o.srcRow}行）`)
      .join('\n');
    const trackPreview = trackUpdates.slice(0, 20)
      .map(o => `EMSリスト${o.row}行: ${o.purchaseNo} / ${o.code} x${o.qty} / ${o.oldTrack || '(空)'} -> ${o.track}（大邱${o.srcRow}行）`)
      .join('\n');
    const codePreview = codeUpdates.slice(0, 20)
      .map(o => `EMSリスト${o.row}行: ${o.purchaseNo || '(購入No空)'} / ${o.track} x${o.qty} / ${o.oldCode || '(空)'} -> ${o.code}（大邱${o.srcRow}行）`)
      .join('\n');
    const datePreview = dateUpdates.slice(0, 20)
      .map(o => `EMSリスト${o.row}行: ${o.label} -> ${EMS_日付キー_(o.value)}`)
      .join('\n');
    const extraPreview = extraRows.slice(0, 20)
      .map(o => `EMSリスト${o.row}行: ${o.purchaseNo || '(購入No空)'} / ${o.track} / ${o.code} / ${o.qty}`)
      .join('\n');
    const res = ui.alert(
      'EMSリスト不足行の復元',
      `EMS大邱作業データを正として、EMSリストに無い ${rows.length}行を下へ追記し、EMS番号違い ${trackUpdates.length}行・商品コード違い ${codeUpdates.length}行・数量違い ${qtyUpdates.length}行を修正します。\nEMSリストだけにある ${extraRows.length}行は削除します。\n\n【追記】\n${preview || 'なし'}` +
      (rows.length > 20 ? '\n…ほか' : '') +
      `\n\n【EMS番号修正】\n${trackPreview || 'なし'}` +
      (trackUpdates.length > 20 ? '\n…ほか' : '') +
      `\n\n【商品コード修正】\n${codePreview || 'なし'}` +
      (codeUpdates.length > 20 ? '\n…ほか' : '') +
      `\n\n【数量修正】\n${updatePreview || 'なし'}` +
      (qtyUpdates.length > 20 ? '\n…ほか' : '') +
      `\n\n【日付更新】\n${datePreview || 'なし'}` +
      (dateUpdates.length > 20 ? '\n…ほか' : '') +
      `\n\n【EMSリストだけにある行（削除）】\n${extraPreview || 'なし'}` +
      (extraRows.length > 20 ? '\n…ほか' : '') + '\n\n実行する？',
      ui.ButtonSet.YES_NO
    );
    if (res !== ui.Button.YES) { ui.alert('やめました。'); return; }

    qtyUpdates.forEach(update => {
      dst.getRange(update.row, 10).setValue(update.qty); // J 数量
    });
    trackUpdates.forEach(update => {
      dst.getRange(update.row, 13).setValue(update.track); // M EMS番号
    });
    codeUpdates.forEach(update => {
      dst.getRange(update.row, 9).setValue(update.code); // I 商品コード
    });
    const updatedDates = EMS_日付更新を適用_(dst, dateUpdates);
    extraRows
      .map(o => o.row)
      .sort((a, b) => b - a)
      .forEach(row => dst.deleteRow(row));

    if (rows.length > 0) {
      const startRow = H2E_findNextAppendRow_(dst);
      let no = H2E_getNextEmsNo_(dst);
      const n = rows.length;

      dst.getRange(startRow, 1, n, 1).setValues(rows.map(() => [no++]));        // A No.
      dst.getRange(startRow, 2, n, 1).setValues(rows.map(o => [o.arrival]));    // B 入荷日
      dst.getRange(startRow, 3, n, 1).setValues(rows.map(o => [o.ship]));       // C EMS発送日
      if (typeof EMS_fillArrivalEstimateRows_ === 'function') {
        EMS_fillArrivalEstimateRows_(dst, startRow, n, false);                  // E 到着予定日
      }
      dst.getRange(startRow, 4, n, 1).setValues(rows.map(() => ['⇒']));         // D 矢印
      dst.getRange(startRow, 6, n, 1).setValues(rows.map(o => [o.purchaseNo])); // F 購入No
      dst.getRange(startRow, 7, n, 1).setValues(rows.map(() => ['未着']));      // G ステータス
      dst.getRange(startRow, 9, n, 1).setValues(rows.map(o => [o.code]));       // I 商品コード
      dst.getRange(startRow, 10, n, 1).setValues(rows.map(o => [o.qty]));       // J 数量
      dst.getRange(startRow, 11, n, 1).setValues(rows.map(o => [o.item]));      // K 品目
      dst.getRange(startRow, 13, n, 1).setValues(rows.map(o => [o.track]));     // M EMS番号
    }

    SpreadsheetApp.flush();
    // 発送日では並べ替えない（EMS大邱作業データと同じ追記順を保つ）
    if (typeof 発注_EMS発送数数式を一括修正 === 'function') {
      発注_EMS発送数数式を一括修正();
    }

    ss.toast(`EMSリスト同期: 追記 ${rows.length}件 / 削除 ${extraRows.length}件 / EMS番号修正 ${trackUpdates.length}件 / 商品コード修正 ${codeUpdates.length}件 / 数量修正 ${qtyUpdates.length}件 / 日付更新 ${updatedDates}件`);
  } finally {
    lock.releaseLock();
  }
}

function EMS_箱照合Mapへ追加_(map, track, qty, rowNo) {
  const key = EMS_表示追跡番号_(track);
  if (!key) return;
  const q = Number(EMS_転送数量キー_(qty)) || 0;
  if (!map[key]) map[key] = { rows: 0, qty: 0, rowNos: [] };
  map[key].rows++;
  map[key].qty += q;
  if (rowNo) map[key].rowNos.push(rowNo);
}

function EMS_大邱箱Map_(src) {
  const vals = src.getDataRange().getValues();
  const map = {};
  for (let i = 0; i < vals.length; i++) {
    const track = EMS_表示追跡番号_(vals[i][3]); // D EMS番号
    if (!track || /TRACKING|TRAKKING|追跡/i.test(String(vals[i][3] || ''))) continue;
    EMS_箱照合Mapへ追加_(map, track, vals[i][8], i + 1); // I 数量
  }
  return map;
}

function EMS_リスト箱Map_(dst) {
  const vals = dst.getDataRange().getValues();
  let h = -1;
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]).trim() === 'No.') { h = i; break; }
  const map = {};
  if (h < 0) return map;
  for (let i = h + 1; i < vals.length; i++) {
    EMS_箱照合Mapへ追加_(map, vals[i][12], vals[i][9], i + 1); // M EMS番号 / J 数量
  }
  return map;
}

function EMS_箱照合行_(track, a, b) {
  const av = a || { rows: 0, qty: 0, rowNos: [] };
  const bv = b || { rows: 0, qty: 0, rowNos: [] };
  return track + '  大邱:' + av.rows + '行/' + av.qty + '個  EMSリスト:' + bv.rows + '行/' + bv.qty + '個';
}

function EMS番号_列を正規化_(sh, col, startRow) {
  const lastRow = sh.getLastRow();
  if (lastRow < startRow) return 0;
  const n = lastRow - startRow + 1;
  const range = sh.getRange(startRow, col, n, 1);
  const vals = range.getDisplayValues();
  let changed = 0;
  const out = vals.map(row => {
    const original = String(row[0] || '');
    if (!original || /Tracking|Trakking|追跡/i.test(original)) return [original];
    const cleaned = EMS_表示追跡番号_(original);
    if (original !== cleaned) changed++;
    return [cleaned || original];
  });
  if (changed > 0) range.setValues(out);
  return changed;
}

function EMS番号_大邱とEMSリストを正規化() {
  const ss = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(DAEGU_CFG.EMS_SRC);
  const dst = ss.getSheetByName(DAEGU_CFG.EMS_DST);
  if (!src || !dst) { SpreadsheetApp.getUi().alert('EMS大邱作業データかEMSリストが見つかりません。'); return; }

  const d = EMS番号_列を正規化_(src, 4, 1);  // EMS大邱 D列
  const e = EMS番号_列を正規化_(dst, 13, 7); // EMSリスト M列
  SpreadsheetApp.getActive().toast('EMS番号を正規化: EMS大邱 ' + d + '件 / EMSリスト ' + e + '件');
}

function EMS大邱_EMS番号ごとに罫線を更新() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (!sh || sh.getName() !== DAEGU_CFG.EMS_SRC) {
    SpreadsheetApp.getUi().alert('EMS大邱作業データで実行してください。');
    return;
  }

  const count = EMS大邱_EMS番号ごとに罫線を更新_(sh);
  SpreadsheetApp.getActive().toast('EMS大邱: EMS番号ごとの罫線を更新しました: ' + count + '箱');
}

function EMS大邱_EMS番号ごとに罫線を更新_(sh) {
  const startRow = 3;
  const emsNoCol = 4; // D列 EMS番号
  const startCol = 1;
  const colCount = Math.min(20, sh.getMaxColumns()); // A:T
  const lastRow = EMS大邱_罫線最終データ行_(sh);
  if (lastRow < startRow) return 0;

  const numRows = lastRow - startRow + 1;
  sh.getRange(startRow, startCol, numRows, colCount)
    .setBorder(false, false, false, false, false, false);

  const emsValues = sh.getRange(startRow, emsNoCol, numRows, 1)
    .getDisplayValues()
    .map(row => EMS_表示追跡番号_(row[0]));

  let groupStartIndex = null;
  let currentEmsNo = '';
  let groupCount = 0;

  for (let i = 0; i <= numRows; i++) {
    const isEnd = i === numRows;
    const emsNo = isEnd ? '' : emsValues[i];

    if (groupStartIndex === null) {
      if (!isEnd && emsNo !== '') {
        groupStartIndex = i;
        currentEmsNo = emsNo;
      }
      continue;
    }

    if (isEnd || emsNo !== currentEmsNo) {
      const groupRange = sh.getRange(startRow + groupStartIndex, startCol, i - groupStartIndex, colCount);
      groupRange.setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);
      groupRange.setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
      groupCount++;

      if (!isEnd && emsNo !== '') {
        groupStartIndex = i;
        currentEmsNo = emsNo;
      } else {
        groupStartIndex = null;
        currentEmsNo = '';
      }
    }
  }

  return groupCount;
}

function EMS大邱_罫線最終データ行_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 3) return 2;
  const numRows = lastRow - 2;
  const values = sh.getRange(3, 1, numRows, 20).getDisplayValues();
  const checkIndexes = [3, 7, 8, 10, 19]; // D/H/I/K/T

  for (let i = values.length - 1; i >= 0; i--) {
    if (checkIndexes.some(idx => String(values[i][idx] || '').trim() !== '')) {
      return 3 + i;
    }
  }
  return 2;
}

function EMSリスト_大邱と箱数を照合() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const src = ss.getSheetByName(DAEGU_CFG.EMS_SRC);
  const dst = ss.getSheetByName(DAEGU_CFG.EMS_DST);
  if (!src || !dst) { ui.alert('「' + DAEGU_CFG.EMS_SRC + '」か「' + DAEGU_CFG.EMS_DST + '」が見つかりません。'); return; }

  const daegu = EMS_大邱箱Map_(src);
  const list = EMS_リスト箱Map_(dst);
  const dKeys = Object.keys(daegu).sort();
  const lKeys = Object.keys(list).sort();
  const lSet = new Set(lKeys);
  const dSet = new Set(dKeys);

  const missing = dKeys.filter(k => !lSet.has(k));
  const extra = lKeys.filter(k => !dSet.has(k));
  const diff = dKeys.filter(k => lSet.has(k) && (daegu[k].rows !== list[k].rows || daegu[k].qty !== list[k].qty));

  const lines = [];
  lines.push('EMS大邱作業データ: ' + dKeys.length + '箱');
  lines.push('EMSリスト: ' + lKeys.length + '箱');
  lines.push('不足箱: ' + missing.length + ' / 数量・行数違い: ' + diff.length + ' / EMSリストだけにある箱: ' + extra.length);

  if (missing.length) {
    lines.push('');
    lines.push('【不足箱（大邱にあるがEMSリストにない）】');
    missing.slice(0, 30).forEach(k => lines.push(EMS_箱照合行_(k, daegu[k], list[k])));
    if (missing.length > 30) lines.push('...ほか ' + (missing.length - 30) + '箱');
  }

  if (diff.length) {
    lines.push('');
    lines.push('【数量・行数違い】');
    diff.slice(0, 30).forEach(k => lines.push(EMS_箱照合行_(k, daegu[k], list[k])));
    if (diff.length > 30) lines.push('...ほか ' + (diff.length - 30) + '箱');
  }

  if (!missing.length && !diff.length) {
    lines.push('');
    lines.push('大邱とEMSリストの箱数・数量は一致しています。');
  } else {
    lines.push('');
    lines.push('合わせるには「EMSリスト：大邱から不足行を確認して復元」を実行してください。');
  }

  ui.alert('EMS大邱とEMSリストの箱数照合', lines.join('\n'), ui.ButtonSet.OK);
}

function 大邱発注_発注No情報_(value) {
  const raw = String(value || '')
    .trim()
    .replace(/[＿]/g, '_')
    .replace(/[\s\u3000]+/g, '')
    .replace(/_+/g, '_');
  if (!raw || /発注NO|OrderNo|DorderDate/i.test(raw)) return null;

  const m = raw.match(/^(.+\d{6,8})_([^_]+)(?:_(\d+)(?:_\d+)*)?$/);
  if (m) {
    const seq = Number(m[3]) || 0;
    const base = m[1] + '_' + m[2];
    return {
      base: base,
      seq: seq,
      numbered: seq > 0,
      normalized: seq > 0 ? base + '_' + seq : base,
    };
  }

  return {
    base: raw.replace(/_+$/, ''),
    seq: 0,
    numbered: false,
    normalized: raw.replace(/_+$/, ''),
  };
}

function 大邱発注_データ発注No情報_(value) {
  const info = 大邱発注_発注No情報_(value);
  if (!info) return null;
  return /\d{6}/.test(info.base) ? info : null;
}

function 大邱発注_連番管理_() {
  const used = {};
  const next = {};

  const has = (base, seq) => !!(used[base] && used[base][seq]);

  const reserve = (base, seq) => {
    if (!base || !seq) return;
    if (!used[base]) used[base] = {};
    used[base][seq] = true;
    if (!next[base] || next[base] <= seq) next[base] = seq + 1;
  };

  const takeNext = base => {
    if (!next[base]) next[base] = 1;
    while (has(base, next[base])) next[base]++;
    const seq = next[base];
    reserve(base, seq);
    return seq;
  };

  return {
    reserve: reserve,
    assign: info => {
      if (info.numbered && !has(info.base, info.seq)) {
        reserve(info.base, info.seq);
        return info.seq;
      }
      return takeNext(info.base);
    },
  };
}

function 大邱発注_チェックボックスを付ける_(sh, startRow, numRows) {
  if (!sh || numRows <= 0) return;

  const fVals = sh.getRange(startRow, 6, numRows, 1).getDisplayValues();
  const validations = sh.getRange(startRow, 1, numRows, 1).getDataValidations();
  for (let i = 0; i < numRows; i++) {
    const info = 大邱発注_データ発注No情報_(fVals[i][0]);
    if (!info) continue;

    const rule = validations[i][0];
    if (rule && rule.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX) continue;

    sh.getRange(startRow + i, 1).insertCheckboxes();
  }
}

function 大邱発注_カートグループキー_(value) {
  const info = 大邱発注_データ発注No情報_(value);
  if (info) return info.base;
  return '';
}

function 大邱発注_データ範囲_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 1) return null;
  const fVals = sh.getRange(1, 6, lastRow, 1).getDisplayValues();
  let first = -1, last = -1;
  for (let i = 0; i < fVals.length; i++) {
    if (大邱発注_データ発注No情報_(fVals[i][0])) {
      if (first < 0) first = i + 1;
      last = i + 1;
    }
  }
  return first < 0 ? null : { first: first, last: last, rows: last - first + 1 };
}

function 大邱発注_罫線列数_(sh) {
  return Math.min(26, sh.getMaxColumns()); // A:Z
}

function 大邱発注_格子罫線_(sh, startRow, numRows) {
  if (!sh || !startRow || !numRows || numRows <= 0) return;
  const maxRows = sh.getMaxRows();
  startRow = Math.max(1, startRow);
  let endRow = Math.min(maxRows, startRow + numRows - 1);
  const dataRange = 大邱発注_データ範囲_(sh);
  if (dataRange) {
    startRow = Math.max(dataRange.first, startRow);
    endRow = Math.min(dataRange.last, endRow);
  }
  numRows = endRow - startRow + 1;
  if (numRows <= 0) return;

  const maxCol = 大邱発注_罫線列数_(sh);
  const fVals = sh.getRange(startRow, 6, numRows, 1).getDisplayValues();
  const checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();

  sh.getRange(startRow, 1, numRows, maxCol)
    .setBorder(false, false, false, false, false, false);

  const chkRange = sh.getRange(startRow, 1, numRows, 1);
  chkRange.clearDataValidations();

  for (let i = 0; i < numRows; i++) {
    const rowNo = startRow + i;
    const hasNo = !!大邱発注_データ発注No情報_(fVals[i][0]);
    const chkCell = sh.getRange(rowNo, 1);

    if (hasNo) {
      chkCell.setDataValidation(checkboxRule);
      sh.getRange(rowNo, 1, 1, maxCol)
        .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
    } else {
      chkCell.clearContent();
    }
  }
}

function 大邱発注_drawGroupBorders_(sh) {
  const range = 大邱発注_データ範囲_(sh);
  if (!range) return;

  const maxCol = 大邱発注_罫線列数_(sh);
  const keys = sh.getRange(range.first, 6, range.rows, 1)
    .getDisplayValues()
    .map(r => 大邱発注_カートグループキー_(r[0]));

  let gStart = null, cur = '';
  for (let i = 0; i <= range.rows; i++) {
    const isEnd = i === range.rows;
    const key = isEnd ? '' : keys[i];

    if (gStart === null) {
      if (!isEnd && key !== '') { gStart = i; cur = key; }
      continue;
    }

    if (isEnd || key !== cur) {
      sh.getRange(range.first + gStart, 1, i - gStart, maxCol)
        .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
      if (!isEnd && key !== '') { gStart = i; cur = key; }
      else { gStart = null; cur = ''; }
    }
  }
}

function 大邱発注_グループ罫線() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== DAEGU_CFG.HACHU_SRC) {
    SpreadsheetApp.getUi().alert('発注リスト大邱データで実行してください。');
    return;
  }

  const range = 大邱発注_データ範囲_(sh);
  if (!range) {
    SpreadsheetApp.getUi().alert('F列の発注NOがあるデータ行が見つかりません。');
    return;
  }

  大邱発注_格子罫線_(sh, range.first, range.rows);
  大邱発注_drawGroupBorders_(sh);
  SpreadsheetApp.getActive().toast('発注リスト大邱データのグループ罫線を更新しました。');
}

function 枝番_コードキー_(code) {
  if (!code) return [];
  if (typeof codeKeys_ === 'function') return codeKeys_(code);
  return [normCode_(code)];
}

function 枝番_数量_(value) {
  const n = Number(String(value || '').replace(/,/g, '').trim());
  return isFinite(n) ? n : 0;
}

function 枝番_注文行を採番_(rows) {
  const tracker = 大邱発注_連番管理_();
  const orderLines = [];
  const updates = [];

  rows.forEach(row => {
    const info = 大邱発注_データ発注No情報_(row.no);
    if (!info || !row.code) return;

    const seq = tracker.assign(info);
    const fixedNo = info.base + '_' + seq;
    if (String(row.no || '').trim() !== fixedNo) {
      updates.push({ row: row.row, value: fixedNo });
    }

    orderLines.push({
      row: row.row,
      base: info.base,
      no: fixedNo,
      code: row.code,
      qty: row.qty,
    });
  });

  return { orderLines: orderLines, updates: updates };
}

function 枝番_注文ラインMap_(orderLines) {
  const map = {};

  orderLines.forEach(line => {
    const item = {
      no: line.no,
      qty: line.qty,
      remaining: line.qty,
    };
    枝番_コードキー_(line.code).forEach(codeKey => {
      const key = line.base + '|' + normCode_(codeKey);
      (map[key] = map[key] || []).push(item);
    });
  });

  return map;
}

function 枝番_EMS購入No更新計画_(emsRows, orderMap) {
  const pointers = {};
  const updates = [];
  let noMatch = 0;

  emsRows.forEach(row => {
    const info = 大邱発注_データ発注No情報_(row.no);
    if (!info || !row.code) return;

    let lines = null;
    let matchedKey = '';
    const keys = 枝番_コードキー_(row.code);
    for (let i = 0; i < keys.length; i++) {
      const key = info.base + '|' + normCode_(keys[i]);
      if (orderMap[key] && orderMap[key].length > 0) {
        lines = orderMap[key];
        matchedKey = key;
        break;
      }
    }

    if (!lines) {
      noMatch++;
      return;
    }

    let p = pointers[matchedKey] || 0;
    while (p < lines.length && lines[p].remaining <= 0) p++;
    if (p >= lines.length) {
      noMatch++;
      pointers[matchedKey] = p;
      return;
    }

    const fixedNo = lines[p].no;
    if (String(row.no || '').trim() !== fixedNo) {
      updates.push({ row: row.row, value: fixedNo });
    }

    lines[p].remaining -= row.qty;
    while (p < lines.length - 1 && lines[p].remaining <= 0) p++;
    pointers[matchedKey] = p;
  });

  return { updates: updates, noMatch: noMatch };
}

function 枝番_列へ書き込み_(sh, col, updates) {
  updates.forEach(update => {
    sh.getRange(update.row, col).setValue(update.value);
  });
}

function 枝番_文字キー_(value) {
  return String(value || '').trim().replace(/\s+/g, ' ');
}

function 枝番_数量キー_(value) {
  const n = 枝番_数量_(value);
  return n ? String(n) : 枝番_文字キー_(value);
}

function 枝番_追跡キー_(value) {
  if (typeof normTrack_ === 'function') return normTrack_(value);
  return String(value || '').replace(/[\s\u3000]/g, '').toUpperCase();
}

function 枝番_map追加_(map, key, item) {
  if (!key) return;
  (map[key] = map[key] || []).push(item);
}

function 枝番_map取得_(maps, keyLists) {
  for (let i = 0; i < maps.length; i++) {
    const map = maps[i];
    const keys = keyLists[i] || [];
    for (let j = 0; j < keys.length; j++) {
      const queue = map[keys[j]];
      if (!queue) continue;
      while (queue.length && queue[0].used) queue.shift();
      if (queue.length) {
        queue[0].used = true;
        return queue[0];
      }
    }
  }
  return null;
}

function 大邱_発注EMS枝番を一括修正(silent) {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const hac = ss.getSheetByName(DAEGU_CFG.HACHU_SRC);
  const ems = ss.getSheetByName(DAEGU_CFG.EMS_SRC);
  if (!hac || !ems) {
    if (!silent) ui.alert('発注リスト大邱データかEMS大邱作業データが見つかりません。');
    return { order: 0, ems: 0, noMatch: 0 };
  }

  const hVals = hac.getDataRange().getValues();
  const orderRows = [];
  for (let i = 0; i < hVals.length; i++) {
    const no = String(hVals[i][5] || '').trim();      // F 発注NO
    const code = String(hVals[i][10] || '').trim();   // K 商品コード
    const info = 大邱発注_データ発注No情報_(no);
    if (!info || !code) continue;
    orderRows.push({ row: i + 1, no: no, code: code, qty: 枝番_数量_(hVals[i][11]) });
  }

  const orderPlan = 枝番_注文行を採番_(orderRows);
  const orderMap = 枝番_注文ラインMap_(orderPlan.orderLines);

  const eVals = ems.getDataRange().getValues();
  const emsRows = [];
  for (let i = 0; i < eVals.length; i++) {
    const no = String(eVals[i][19] || '').trim();     // T 購入No
    const code = String(eVals[i][7] || '').trim();    // H 商品コード
    const info = 大邱発注_データ発注No情報_(no);
    if (!info || !code) continue;
    emsRows.push({ row: i + 1, no: no, code: code, qty: 枝番_数量_(eVals[i][8]) });
  }
  const emsPlan = 枝番_EMS購入No更新計画_(emsRows, orderMap);

  if (!silent) {
    const res = ui.alert(
      '大邱 発注No枝番一括修正',
      '発注リスト大邱データ F列: ' + orderPlan.updates.length + '件\n' +
      'EMS大邱作業データ T列: ' + emsPlan.updates.length + '件\n' +
      'EMS大邱で対応発注なし: ' + emsPlan.noMatch + '件\n\n実行する？',
      ui.ButtonSet.YES_NO
    );
    if (res !== ui.Button.YES) return { order: 0, ems: 0, noMatch: emsPlan.noMatch };
  }

  枝番_列へ書き込み_(hac, 6, orderPlan.updates);
  枝番_列へ書き込み_(ems, 20, emsPlan.updates);
  SpreadsheetApp.flush();

  if (!silent) {
    ss.toast('大邱枝番修正: 発注 ' + orderPlan.updates.length + '件 / EMS大邱 ' + emsPlan.updates.length + '件');
  }
  return { order: orderPlan.updates.length, ems: emsPlan.updates.length, noMatch: emsPlan.noMatch };
}

function 大邱データに発注とEMSを合わせる() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const hacSrc = ss.getSheetByName(DAEGU_CFG.HACHU_SRC);
  const emsSrc = ss.getSheetByName(DAEGU_CFG.EMS_SRC);
  const hacDst = ss.getSheetByName(DAEGU_CFG.HACHU_DST);
  const emsDst = ss.getSheetByName(DAEGU_CFG.EMS_DST);
  if (!hacSrc || !emsSrc || !hacDst || !emsDst) {
    ui.alert('必要なシート（大邱/発注/EMSリスト）のいずれかが見つかりません。');
    return;
  }

  const startRes = ui.alert(
    '大邱データに発注/EMSを合わせる',
    '大邱側の枝番を整えたうえで、発注G列とEMSリストF列を大邱データの購入Noへ合わせます。\n\n実行する？',
    ui.ButtonSet.YES_NO
  );
  if (startRes !== ui.Button.YES) {
    ui.alert('やめました。');
    return;
  }

  const srcFix = 大邱_発注EMS枝番を一括修正(true);

  const exactOrderMap = {}, codeQtyOrderMap = {}, codeOrderMap = {};
  const hSrcVals = hacSrc.getDataRange().getValues();
  for (let i = 0; i < hSrcVals.length; i++) {
    const no = String(hSrcVals[i][5] || '').trim();     // F
    const info = 大邱発注_データ発注No情報_(no);
    const code = String(hSrcVals[i][10] || '').trim();  // K
    if (!info || !code) continue;

    const item = { no: no, used: false };
    const qty = 枝番_数量キー_(hSrcVals[i][11]);        // L
    const name = 枝番_文字キー_(hSrcVals[i][8]);        // I
    const option = 枝番_文字キー_(hSrcVals[i][9]);      // J
    枝番_コードキー_(code).forEach(codeKey => {
      const c = normCode_(codeKey);
      枝番_map追加_(exactOrderMap, [info.base, c, qty, name, option].join('|'), item);
      枝番_map追加_(codeQtyOrderMap, [info.base, c, qty].join('|'), item);
      枝番_map追加_(codeOrderMap, [info.base, c].join('|'), item);
    });
  }

  const hLast = hacDst.getLastRow();
  const hUpdates = [];
  const hUpdateByRow = {};
  if (hLast >= 7) {
    const hVals = hacDst.getRange(7, 1, hLast - 6, Math.max(hacDst.getLastColumn(), 13)).getValues();
    for (let i = 0; i < hVals.length; i++) {
      const rowNo = 7 + i;
      const no = String(hVals[i][6] || '').trim();       // G
      const info = 大邱発注_データ発注No情報_(no);
      const code = String(hVals[i][11] || '').trim();    // L
      if (!info || !code) continue;

      const qty = 枝番_数量キー_(hVals[i][12]);          // M
      const name = 枝番_文字キー_(hVals[i][9]);          // J
      const option = 枝番_文字キー_(hVals[i][10]);       // K
      const exactKeys = [], codeQtyKeys = [], codeKeys = [];
      枝番_コードキー_(code).forEach(codeKey => {
        const c = normCode_(codeKey);
        exactKeys.push([info.base, c, qty, name, option].join('|'));
        codeQtyKeys.push([info.base, c, qty].join('|'));
        codeKeys.push([info.base, c].join('|'));
      });

      const matched = 枝番_map取得_(
        [exactOrderMap, codeQtyOrderMap, codeOrderMap],
        [exactKeys, codeQtyKeys, codeKeys]
      );
      if (matched && no !== matched.no) {
        hUpdates.push({ row: rowNo, value: matched.no });
        hUpdateByRow[rowNo] = matched.no;
      }
    }
  }

  const exactEmsMap = {}, trackCodeQtyEmsMap = {}, trackCodeEmsMap = {}, codeQtyEmsMap = {}, codeEmsMap = {};
  const eSrcVals = emsSrc.getDataRange().getValues();
  for (let i = 0; i < eSrcVals.length; i++) {
    const no = String(eSrcVals[i][19] || '').trim();     // T
    const info = 大邱発注_データ発注No情報_(no);
    const code = String(eSrcVals[i][7] || '').trim();    // H
    if (!info || !code) continue;

    const item = { no: no, used: false };
    const qty = 枝番_数量キー_(eSrcVals[i][8]);          // I
    const track = 枝番_追跡キー_(eSrcVals[i][3]);        // D
    枝番_コードキー_(code).forEach(codeKey => {
      const c = normCode_(codeKey);
      if (track) 枝番_map追加_(exactEmsMap, [info.base, track, c, qty].join('|'), item);
      if (track) 枝番_map追加_(trackCodeQtyEmsMap, [track, c, qty].join('|'), item);
      if (track) 枝番_map追加_(trackCodeEmsMap, [track, c].join('|'), item);
      枝番_map追加_(codeQtyEmsMap, [info.base, c, qty].join('|'), item);
      枝番_map追加_(codeEmsMap, [info.base, c].join('|'), item);
    });
  }

  const eLast = emsDst.getLastRow();
  const eUpdates = [];
  const eUpdateByRow = {};
  if (eLast >= 7) {
    const eVals = emsDst.getRange(7, 1, eLast - 6, Math.max(emsDst.getLastColumn(), 13)).getValues();
    for (let i = 0; i < eVals.length; i++) {
      const rowNo = 7 + i;
      const no = String(eVals[i][5] || '').trim();       // F
      const info = 大邱発注_データ発注No情報_(no);
      const code = String(eVals[i][8] || '').trim();     // I
      if (!code) continue;

      const qty = 枝番_数量キー_(eVals[i][9]);           // J
      const track = 枝番_追跡キー_(eVals[i][12]);        // M
      const exactKeys = [], trackCodeQtyKeys = [], trackCodeKeys = [], codeQtyKeys = [], codeKeys = [];
      枝番_コードキー_(code).forEach(codeKey => {
        const c = normCode_(codeKey);
        if (info && track) exactKeys.push([info.base, track, c, qty].join('|'));
        if (track) trackCodeQtyKeys.push([track, c, qty].join('|'));
        if (track) trackCodeKeys.push([track, c].join('|'));
        if (info) codeQtyKeys.push([info.base, c, qty].join('|'));
        if (info) codeKeys.push([info.base, c].join('|'));
      });

      const matched = 枝番_map取得_(
        [exactEmsMap, trackCodeQtyEmsMap, trackCodeEmsMap, codeQtyEmsMap, codeEmsMap],
        [exactKeys, trackCodeQtyKeys, trackCodeKeys, codeQtyKeys, codeKeys]
      );
      if (matched && no !== matched.no) {
        eUpdates.push({ row: rowNo, value: matched.no });
        eUpdateByRow[rowNo] = matched.no;
      }
    }

    const hOrderLines = [];
    if (hLast >= 7) {
      const hValsForEms = hacDst.getRange(7, 1, hLast - 6, Math.max(hacDst.getLastColumn(), 13)).getValues();
      for (let i = 0; i < hValsForEms.length; i++) {
        const rowNo = 7 + i;
        const no = hUpdateByRow[rowNo] || String(hValsForEms[i][6] || '').trim(); // G
        const info = 大邱発注_データ発注No情報_(no);
        const code = String(hValsForEms[i][11] || '').trim();                     // L
        if (!info || !code) continue;
        hOrderLines.push({
          row: rowNo,
          base: info.base,
          no: no,
          code: code,
          qty: 枝番_数量_(hValsForEms[i][12]),                                    // M
        });
      }
    }

    const emsRowsForHachu = [];
    for (let i = 0; i < eVals.length; i++) {
      const rowNo = 7 + i;
      const no = eUpdateByRow[rowNo] || String(eVals[i][5] || '').trim();          // F
      const info = 大邱発注_データ発注No情報_(no);
      const code = String(eVals[i][8] || '').trim();                              // I
      if (!info || !code) continue;
      emsRowsForHachu.push({
        row: rowNo,
        no: no,
        code: code,
        qty: 枝番_数量_(eVals[i][9]),                                             // J
      });
    }

    const emsByHachuPlan = 枝番_EMS購入No更新計画_(
      emsRowsForHachu,
      枝番_注文ラインMap_(hOrderLines)
    );
    emsByHachuPlan.updates.forEach(update => {
      if (eUpdateByRow[update.row]) return;
      eUpdates.push(update);
      eUpdateByRow[update.row] = update.value;
    });
  }

  枝番_列へ書き込み_(hacDst, 7, hUpdates);
  枝番_列へ書き込み_(emsDst, 6, eUpdates);
  SpreadsheetApp.flush();

  if (typeof 発注_EMS発送数数式を一括修正 === 'function') {
    発注_EMS発送数数式を一括修正();
  }

  ss.toast(
    '大邱データに合わせました: 大邱発注 ' + srcFix.order +
    '件 / 大邱EMS ' + srcFix.ems +
    '件 / 発注 ' + hUpdates.length +
    '件 / EMSリスト ' + eUpdates.length + '件',
    '手動運用',
    8
  );
}

// ============================================================
// 発注リスト大邱データ：F列(発注NO)を一意化する採番
//   同じ発注NO(=カート決済No)ごとに、上から _1,_2,… を付与
//   例) 20260407_01 が4行 → 20260407_01_1, _2, _3, _4
//   既に連番付きでも重複していれば次の空き番号へ振り直す
// ============================================================
function 大邱発注_一意Noを採番() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const sh = ss.getSheetByName(DAEGU_CFG.HACHU_SRC);
  if (!sh) { ui.alert('「' + DAEGU_CFG.HACHU_SRC + '」が見つかりません。'); return; }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(2000)) { ss.toast('他の処理が実行中です。'); return; }
  try {
    const FCOL = 6; // F列 発注NO
    const lastRow = sh.getLastRow();
    if (lastRow < 1) { ss.toast('データがありません。'); return; }

    const vals = sh.getRange(1, FCOL, lastRow, 1).getValues();
    const tracker = 大邱発注_連番管理_();
    const newCol = vals.map(r => [r[0]]);
    let count = 0;
    const samples = [];

    for (let i = 0; i < vals.length; i++) {
      const info = 大邱発注_データ発注No情報_(vals[i][0]);
      if (!info) continue;
      const seq = tracker.assign(info);
      const fixed = info.base + '_' + seq;
      if (String(vals[i][0] || '').trim() !== fixed) {
        newCol[i][0] = fixed;
        count++;
        if (samples.length < 12) samples.push(fixed);
      }
    }

    if (count === 0) { ss.toast('採番対象なし（重複なし）。'); return; }

    const res = ui.alert('一意No採番',
      count + '件のF列に連番を付けます（例）：\n\n' + samples.join('\n') +
      (count > 12 ? '\n…ほか' : '') + '\n\nF列を書き換えます。実行する？',
      ui.ButtonSet.YES_NO);
    if (res !== ui.Button.YES) { ui.alert('やめました。'); return; }

    sh.getRange(1, FCOL, lastRow, 1).setValues(newCol);
    ss.toast('一意No採番: ' + count + '件');
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// EMS大邱作業データ：購入No(T列)を発注リスト大邱データから初回補完（先入先出）
//   商品コード(短縮)ごとに、発注大邱の発注ライン(一意No・数量)を行順(先入先出)に並べ、
//   EMS大邱の発送数量で上から消費しながら一意Noを割り当てる。
//   例) KRSJCM03-0506S: 発注[20260402_01_1=80, 20260614_06_1=1]
//       発送 16+16+16+8+16+8=80 → 全部 20260402_01_1、次の1 → 20260614_06_1
//   発注の容量を超えた発送は空のまま（手動）。
//   ※先に「大邱発注_一意Noを採番」を実行してから使う前提。
// ============================================================
function EMS大邱_購入No初回補完() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const ems = ss.getSheetByName(DAEGU_CFG.EMS_SRC);    // EMS大邱作業データ
  const hac = ss.getSheetByName(DAEGU_CFG.HACHU_SRC);  // 発注リスト大邱データ
  if (!ems || !hac) { ui.alert('シートが見つかりません。'); return; }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(2000)) { ss.toast('他の処理が実行中です。'); return; }
  try {
    const T_COL = 20; // EMS大邱 T列 購入No（1始まり: T=20）

    // 発注大邱: 商品コード別名 -> 行順の発注ライン [{no, qty}]（先入先出）
    const hVals = hac.getDataRange().getValues();
    const orderLines = {};
    hVals.forEach(r => {
      const no = String(r[5] || '').trim();        // F 発注NO(=一意購入No)
      const codeRaw = String(r[10] || '').trim();   // K 商品コード
      if (!no || !codeRaw || /発注NO|OrderNo/i.test(no)) return;
      const line = { no: no, qty: Number(r[11]) || 0, remaining: Number(r[11]) || 0 };
      枝番_コードキー_(codeRaw).forEach(code => {
        (orderLines[normCode_(code)] = orderLines[normCode_(code)] || []).push(line);
      });
    });

    // EMS大邱: 行順の発送行（T列が空の行のみ）
    const eVals = ems.getDataRange().getValues();
    const tcol = ems.getRange(1, T_COL, eVals.length, 1).getValues();
    const emsLines = [];
    for (let i = 0; i < eVals.length; i++) {
      const codeRaw = String(eVals[i][7] || '').trim(); // H 商品コード
      if (!codeRaw || /Trakking|Tracking|商品コード/i.test(codeRaw)) continue;
      if (String(tcol[i][0] || '').trim()) continue;     // 既に購入No入りは触らない
      emsLines.push({ i: i, code: codeRaw, qty: Number(eVals[i][8]) || 0 });
    }

    // 先入先出で割り当て（発注ラインの数量を発送数量で消費）
    const pointers = {};
    let filled = 0, nomatch = 0;
    for (const ship of emsLines) {
      const keys = 枝番_コードキー_(ship.code).map(normCode_);
      let matchedKey = '', ords = null;
      for (let k = 0; k < keys.length; k++) {
        const key = keys[k];
        const list = orderLines[key] || [];
        let p = pointers[key] || 0;
        while (p < list.length && list[p].remaining <= 0) p++;
        if (p < list.length) {
          pointers[key] = p;
          matchedKey = key;
          ords = list;
          break;
        }
      }
      if (!ords) { nomatch++; continue; }   // 発注容量を超えた発送→手動

      let p = pointers[matchedKey] || 0;
      while (p < ords.length && ords[p].remaining <= 0) p++;
      if (p >= ords.length) { pointers[matchedKey] = p; nomatch++; continue; }

      tcol[ship.i][0] = ords[p].no;
      filled++;
      ords[p].remaining -= ship.qty;
      while (p < ords.length - 1 && ords[p].remaining <= 0) p++;
      pointers[matchedKey] = p;
    }

    if (filled === 0) { ss.toast('補完できる行なし（該当なし ' + nomatch + '）'); return; }
    const res = ui.alert('EMS大邱 購入No 初回補完（先入先出）',
      filled + '件のT列(購入No)を発注大邱から先入先出で割り当てます。\n' +
      '該当なし（手動）: ' + nomatch + '件\n\n' +
      '※発注ラインの数量を、発送数量で上から順に消費して割り当てます。\n実行する？',
      ui.ButtonSet.YES_NO);
    if (res !== ui.Button.YES) { ui.alert('やめました。'); return; }

    ems.getRange(1, T_COL, eVals.length, 1).setValues(tcol);
    SpreadsheetApp.flush();
    if (typeof 大邱発注_チェックと残り数量を設置 === 'function') {
      大邱発注_チェックと残り数量を設置();
    }
    ss.toast('EMS大邱 購入No補完(FIFO): ' + filled + '件 / 該当なし ' + nomatch + '件');
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 発注リスト大邱データ：A列チェックボックス ＋ Y/Z列「EMS送信済/残り」計算列を設置
// ============================================================
function 大邱発注_チェックと残り数量を設置() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const sh = ss.getSheetByName(DAEGU_CFG.HACHU_SRC);
  const ems = ss.getSheetByName(DAEGU_CFG.EMS_SRC);
  if (!sh) { ui.alert('「' + DAEGU_CFG.HACHU_SRC + '」が見つかりません。'); return; }
  if (!ems) { ui.alert('「' + DAEGU_CFG.EMS_SRC + '」が見つかりません。'); return; }

  const lastRow = sh.getLastRow();
  const fVals = sh.getRange(1, 6, lastRow, 1).getValues(); // F列
  let firstData = -1, lastData = -1;
  for (let i = 0; i < fVals.length; i++) {
    const f = String(fVals[i][0] || '').trim();
    if (f && !/発注NO|OrderNo|DorderDate/i.test(f)) { if (firstData < 0) firstData = i + 1; lastData = i + 1; }
  }
  if (firstData < 0) { ui.alert('データ行が見つかりません。'); return; }
  const n = lastData - firstData + 1;

  sh.getRange(firstData, 1, n, 1).insertCheckboxes();          // A列 チェックボックス
  sh.getRange(firstData - 1, 25).setValue('EMS送信済');         // Y見出し
  sh.getRange(firstData - 1, 26).setValue('残り');              // Z見出し

  const sentRows = {};
  const eVals = ems.getDataRange().getValues();
  for (let i = 0; i < eVals.length; i++) {
    const track = EMS_表示追跡番号_(eVals[i][3]); // D EMS番号
    const no = EMS_転送購入Noキー_(eVals[i][19]); // T 購入No
    const code = eVals[i][7];                      // H 商品コード
    const qty = Number(EMS_転送数量キー_(eVals[i][8])) || 0; // I 数量
    if (!track || !no || !code || !qty) continue;
    (sentRows[no] = sentRows[no] || []).push({
      keys: 枝番_コードキー_(code).map(normCode_),
      qty: qty
    });
  }

  const hVals = sh.getRange(firstData, 1, n, Math.max(sh.getLastColumn(), 26)).getValues();
  const yVals = [], zVals = [];
  for (let i = 0; i < n; i++) {
    const no = EMS_転送購入Noキー_(hVals[i][5]); // F 発注NO
    const code = hVals[i][10];                    // K 商品コード
    const ordered = Number(EMS_転送数量キー_(hVals[i][11])) || 0; // L 数量
    if (!no || !code) {
      yVals.push(['']);
      zVals.push(['']);
      continue;
    }

    const keys = new Set(枝番_コードキー_(code).map(normCode_));
    const sent = (sentRows[no] || []).reduce((acc, item) => {
      return item.keys.some(k => keys.has(k)) ? acc + item.qty : acc;
    }, 0);
    yVals.push([sent || '']);
    zVals.push([ordered ? ordered - sent : '']);
  }
  sh.getRange(firstData, 25, n, 1).setValues(yVals);           // Y EMS送信済
  sh.getRange(firstData, 26, n, 1).setValues(zVals);           // Z 残り

  // 消込色（条件付き書式）: 入荷数(D)あり & 韓国側の残り(Z)>0 → 黄、Z=0 → オレンジ
  // ※ 残り(Z)が空欄の行は色を付けない。
  const cfRange = sh.getRange(firstData, 1, sh.getMaxRows() - firstData + 1, 26); // A..Z データ範囲
  const D = '$D' + firstData, L = '$L' + firstData, Z = '$Z' + firstData;
  const hasRemaining = 'LEN(TRIM(TO_TEXT(' + Z + ')))>0';
  const FY = '=AND(N(' + D + ')>0,N(' + L + ')>0,' + hasRemaining + ',N(' + Z + ')>0)';
  const FO = '=AND(N(' + D + ')>0,N(' + L + ')>0,' + hasRemaining + ',N(' + Z + ')=0)';
  const yellow = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(FY).setBackground('#ffff00').setRanges([cfRange]).build();
  const orange = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(FO).setBackground('#f4b183').setRanges([cfRange]).build();
  // 既存の消込色ルール（$D/$L/$Y/$Zを使うAND式）を除去して入れ直す（旧式も含めて）
  const rules = sh.getConditionalFormatRules().filter(rule => {
    try {
      const c = rule.getBooleanCondition(); if (!c) return true;
      const s = String((c.getCriteriaValues() || [])[0] || '');
      return !(/^=AND\(/.test(s) && (s.indexOf('$Y') >= 0 || s.indexOf('$Z') >= 0) && (s.indexOf('$D') >= 0 || s.indexOf('$L') >= 0));
    } catch (e) { return true; }
  });
  rules.push(orange, yellow);
  sh.setConditionalFormatRules(rules);

  ss.toast('チェックボックス＋残り数量列＋消込色を設置: ' + n + '行');
}

// EMS大邱の最終データ行の次（H商品コード/T購入No/D EMS番号で判定。ヘッダーは除く）
function 大邱EMS_次の追記行_(dst) {
  const lastRow = dst.getLastRow();
  const cols = [8, 20, 4]; // H, T, D
  let last = 2; // ヘッダー2行想定
  if (lastRow < 1) return last + 1;
  cols.forEach(c => {
    const vals = dst.getRange(1, c, lastRow, 1).getDisplayValues();
    for (let i = vals.length - 1; i >= 0; i--) {
      const v = String(vals[i][0] || '').trim();
      if (v && !/Trakking|Tracking|商品コード|購入No/i.test(v)) { const rn = i + 1; if (rn > last) last = rn; break; }
    }
  });
  return last + 1;
}

// ============================================================
// 発注リスト大邱データ：A列チェック行を EMS大邱作業データ へ送る
//   送る列: 購入No(F)・商品コード(K)・数量(L)・入荷日(C)・業者(H)・商品名(I)・品目(N)・重さ(O)・価格(P)
//   EMS番号・発送日 はEMS担当が記入。送ったらチェックを外す。
//   重複判定はしない（分割発送＝同じ行を複数回送れる）。
// ============================================================
function 大邱発注_チェック行をEMS大邱へ送る() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const src = ss.getSheetByName(DAEGU_CFG.HACHU_SRC); // 発注リスト大邱データ
  const dst = ss.getSheetByName(DAEGU_CFG.EMS_SRC);   // EMS大邱作業データ
  if (!src || !dst) { ui.alert('シートが見つかりません。'); return; }

  const t0 = new Date();
  const L = msg => Logger.log('[EMS大邱へ送る] ' + msg);
  L('開始 ' + Utilities.formatDate(t0, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'));

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(2000)) { L('中断: 他の処理が実行中（ロック取得失敗）'); ss.toast('他の処理が実行中です。'); return; }
  try {
    const sVals = src.getDataRange().getValues();
    const picked = [];          // 送信対象（チェック済み＆購入No/コードあり）
    const checkedRows = [];     // A列がチェック(TRUE)の行番号（全部）
    const skipped = [];         // チェックされたが購入No/商品コードが空で送らない行
    const uncheckedData = [];   // 未チェックだが購入No/コードあり＝送ってはいけない行
    for (let i = 0; i < sVals.length; i++) {
      const r = sVals[i];
      const isChecked = (r[0] === true);            // A列チェック
      const no = EMS_転送購入Noキー_(r[5]);         // F 購入No
      const code = 大邱_表示コード_(r[10]);          // K 商品コード
      const qty = EMS_表示数量_(r[11]);              // L 数量

      if (!isChecked) {
        if (no || code) uncheckedData.push(i + 1);   // 未チェック＝送信対象外
        continue;
      }
      checkedRows.push(i + 1);
      if (!no || !code) {
        skipped.push({ row: i + 1, reason: !no ? '購入No空' : '商品コード空' });
        continue;
      }
      picked.push({
        row: i + 1, no: no, code: code,
        qty: qty, date: r[2], vendor: r[7], name: r[8], item: r[13], weight: r[14], price: r[15]
      });
    }

    // スキャン結果をログに残す（チェック/送信対象/未チェックの内訳）
    L('スキャン結果: 全' + sVals.length + '行 / チェック' + checkedRows.length + '件 / 送信対象' + picked.length +
      '件 / チェック済みスキップ' + skipped.length + '件 / 未チェックで除外' + uncheckedData.length + '件');
    picked.forEach(p => L('  送信予定 行' + p.row + ': ' + p.no + ' / ' + p.code + ' / ' + (p.qty || 0) + '個'));
    skipped.forEach(s => L('  ⚠ チェック済みだが送らない 行' + s.row + ': ' + s.reason));

    // 確認: 送信対象がすべてチェック行か（未チェック行の混入がないか）を検証してログ
    const checkedSet = new Set(checkedRows);
    const intruder = picked.map(p => p.row).filter(rn => !checkedSet.has(rn));
    if (intruder.length) L('🛑 異常: 未チェック行が送信対象に混入: 行' + intruder.join(','));
    else L('✔ 確認OK: 送信対象はすべてチェック行（未チェック行は1件も含まれていない）');

    if (picked.length === 0) { L('終了: 送信対象なし'); ss.toast('チェックされた行がありません。'); return; }

    const preview = picked.slice(0, 12).map(p => `${p.no} / ${p.code} / ${p.qty || 0}個`).join('\n');
    const res = ui.alert('EMS大邱へ送る',
      `${picked.length}件をEMS大邱作業データの最終行の下へ送ります。\n\n${preview}` +
      (picked.length > 12 ? '\n…ほか' : '') +
      '\n\n（EMS番号・発送日はEMS担当が記入）\n実行する？',
      ui.ButtonSet.YES_NO);
    if (res !== ui.Button.YES) { L('終了: ユーザーがキャンセル'); ui.alert('やめました。'); return; }

    const startRow = 大邱EMS_次の追記行_(dst);
    const n = picked.length;
    L('書き込み開始: EMS大邱 ' + startRow + '行目から ' + n + '件');
    dst.getRange(startRow, 1, n, 1).setValues(picked.map(p => [p.date]));    // A 入荷日
    dst.getRange(startRow, 5, n, 1).setValues(picked.map(p => [p.vendor]));  // E 業者
    dst.getRange(startRow, 6, n, 1).setValues(picked.map(p => [p.name]));    // F 商品名
    dst.getRange(startRow, 8, n, 1).setValues(picked.map(p => [p.code]));    // H 商品コード
    dst.getRange(startRow, 9, n, 1).setValues(picked.map(p => [p.qty]));     // I 数量
    dst.getRange(startRow, 11, n, 1).setValues(picked.map(p => [p.item]));   // K 品目
    dst.getRange(startRow, 12, n, 1).setValues(picked.map(p => [p.weight])); // L 重量
    dst.getRange(startRow, 13, n, 1).setValues(picked.map(p => [p.price]));  // M 価格
    dst.getRange(startRow, 20, n, 1).setValues(picked.map(p => [p.no]));     // T 購入No

    if (typeof EMS大邱_QRS行範囲を再計算_ === 'function') {
      EMS大邱_QRS行範囲を再計算_(dst, startRow, n);
    }

    // 書き込み後の検証: 追記ブロックを読み戻し「送られた件数」と内容を照合
    SpreadsheetApp.flush();
    const wroteNo = dst.getRange(startRow, 20, n, 1).getValues();   // T 購入No
    const wroteCode = dst.getRange(startRow, 8, n, 1).getValues();  // H 商品コード
    let received = 0, okCount = 0, ngCount = 0;
    for (let i = 0; i < n; i++) {
      const gotNo = EMS_転送購入Noキー_(wroteNo[i][0]);
      const gotCode = 大邱_表示コード_(wroteCode[i][0]);
      if (gotNo || gotCode) received++;                 // 実際にEMS大邱へ入った行
      if (gotNo === picked[i].no && gotCode === picked[i].code) {
        okCount++;
      } else {
        ngCount++;
        L('  ✗ 不一致 EMS大邱' + (startRow + i) + '行: 期待 ' + picked[i].no + '/' + picked[i].code +
          ' → 実際 ' + gotNo + '/' + gotCode);
      }
    }
    L('書き込み検証: 送られた ' + received + '件 / 内容一致 ' + okCount + '件 / 不一致 ' + ngCount + '件');

    picked.forEach(p => src.getRange(p.row, 1).setValue(false));           // チェックを外す
    L('チェック解除: 送信した' + n + '行のA列をFALSEに戻した');

    // 件数チェック: 送った件数(n) と 送られた件数(received) を照合し、最後に表示
    const allOK = (received === n && ngCount === 0);
    const mark = allOK ? '✅ 件数一致しました。' : '⚠️ 件数不一致あり！実行ログを確認してください。';
    L('完了: 送った ' + n + '件 / 送られた ' + received + '件 / 内容一致 ' + okCount + ' / 不一致 ' + ngCount +
      ' / 未チェック除外 ' + uncheckedData.length + '件・スキップ ' + skipped.length + '件 / 所要 ' + (new Date() - t0) + 'ms');
    L('件数チェック: ' + mark);
    ss.toast(`EMS大邱へ送信: ${n}件`);
    ui.alert('EMS大邱へ送信 完了',
      '送った件数　: ' + n + '件\n' +
      '送られた件数: ' + received + '件\n' +
      '照合　　　　: 件数一致 ' + okCount + ' / 件数不一致 ' + ngCount + '\n\n' +
      mark,
      ui.ButtonSet.OK);
  } catch (err) {
    L('🛑 例外: ' + (err && err.stack ? err.stack : err));
    throw err;
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 発注リスト大邱データ：onEditでF列の発注NOに一意の連番を付ける
//   例) 「Wata260618_09_1」を入力済みでも、同じ番号があれば「Wata260618_09_2」にする
//   発注NOが入った行はA列にチェックボックスも付ける。
//   複数行貼り付けも行ごとに重複しない番号へ振る。
//   ※ onEdit(e) 管理スクリプトから、シート名が一致したときに呼ばれる。
// ============================================================
function 大邱発注_onEdit採番_(e) {
  if (!e || !e.range) return;
  const F = 6; // F列 発注NO
  const sc = e.range.getColumn(), ec = e.range.getLastColumn();
  if (sc > F || ec < F) return; // F列を含まない編集はスルー

  const sh = e.range.getSheet();
  const startRow = e.range.getRow(), numRows = e.range.getNumRows();
  const fCells = sh.getRange(startRow, F, numRows, 1);
  const fVals = fCells.getValues();

  const tracker = 大邱発注_連番管理_();
  const lastRow = sh.getLastRow();
  const allF = sh.getRange(1, F, lastRow, 1).getValues();
  const editEndRow = startRow + numRows - 1;

  for (let i = 0; i < allF.length; i++) {
    const rowNo = i + 1;
    if (startRow <= rowNo && rowNo <= editEndRow) continue;

    const info = 大邱発注_データ発注No情報_(allF[i][0]);
    if (!info) continue;

    if (info.numbered) {
      tracker.reserve(info.base, info.seq);
    } else {
      tracker.reserve(info.base, 1);
    }
  }

  let changed = false;
  for (let i = 0; i < fVals.length; i++) {
    const info = 大邱発注_データ発注No情報_(fVals[i][0]);
    if (!info) continue;

    const seq = tracker.assign(info);
    const fixed = info.base + '_' + seq;
    if (String(fVals[i][0] || '').trim() !== fixed) {
      fVals[i][0] = fixed;
      changed = true;
    }
  }
  if (changed) fCells.setValues(fVals);
  大邱発注_チェックボックスを付ける_(sh, startRow, numRows);
  大邱発注_格子罫線_(sh, Math.max(1, startRow - 2), numRows + 4);
  大邱発注_drawGroupBorders_(sh);
}

// ============================================================
// EMSリスト：K列(品目)の関数(XLOOKUP)を外して生データに固定する
//   既存の生値を優先し、無い行は発注(商品コード→品目)から引いて固定。
//   → ARRAYFORMULAが消えるので #REF! も解消、転送が品目を書いても壊れない。
// ============================================================
function EMSリスト_品目を生データ化() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const ems = ss.getSheetByName(DAEGU_CFG.EMS_DST);    // EMSリスト
  const hac = ss.getSheetByName(DAEGU_CFG.HACHU_DST);  // 発注
  if (!ems || !hac) { ui.alert('シートが見つかりません。'); return; }

  // 発注: 商品コード(L列)→品目(O列)（商品コード別名も登録）
  const hVals = hac.getDataRange().getValues();
  const map = {};
  for (let i = 0; i < hVals.length; i++) {
    const code = hVals[i][11];                                         // L列
    const item = String(hVals[i][14] || '').trim();                    // O列
    if (!code || !item) continue;
    枝番_コードキー_(code).forEach(key => {
      key = normCode_(key);
      if (!(key in map)) map[key] = item;
    });
  }

  const eVals = ems.getDataRange().getValues();
  let h = -1;
  for (let i = 0; i < eVals.length; i++) if (String(eVals[i][0]).trim() === 'No.') { h = i; break; }
  if (h < 0) { ui.alert('EMSリストの見出し(No.)が見つかりません。'); return; }
  const startRow = h + 2, n = eVals.length - (h + 1);
  if (n <= 0) { ss.toast('データがありません。'); return; }

  const out = [];
  for (let i = h + 1; i < eVals.length; i++) {
    const code = eVals[i][8];                                         // I列 商品コード
    const cur = String(eVals[i][10] || '').trim();                    // K列 現在値
    let item = '';
    if (cur && !/#REF|#ERROR|#N\/A|#VALUE|#NAME/i.test(cur)) item = cur;  // 既存の生値を優先
    else if (code) {
      const keys = 枝番_コードキー_(code).map(normCode_);
      for (let k = 0; k < keys.length; k++) {
        if (map[keys[k]]) { item = map[keys[k]]; break; }
      }
    }                                                                     // 無ければ発注から固定
    out.push([item]);
  }
  ems.getRange(startRow, 11, n, 1).setValues(out); // K列を生データで上書き（=関数も消える）
  ss.toast('EMSリストK列(品目)を生データ化: ' + n + '行');
}
