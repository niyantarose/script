// ============================================================
// 発注NO採番v2: 「発注日8桁_カート番号_行番号」で全行ユニークにする
//   仕様: docs/superpowers/specs/2026-07-07-orderno-auto-numbering-v2-design.md
//   ・onEditチェーン先頭で実行（後段がタイムアウトしても番号は確定済み）
//   ・日付のみ→カートも自動採番 / 日付_カート→行番号自動 /
//     別業者がカート使用中→次の空きカートへ繰り上げ
//   ・一度付いた番号は変更しない（EMSリスト照合キーが参照するため）
//   ・「採番v2_」関数は純粋ロジック（SpreadsheetApp非依存、Nodeでテスト可能）
// ============================================================

const 採番V2_CFG = {
  COL_NO: 6,      // F列 発注NO
  COL_VENDOR: 8,  // H列 業者
  COL_CODE: 11,   // K列 商品コード
  COL_QTY: 12,    // L列 数量
  FLAG_BG: '#f4cccc', // チェックで塗る赤背景（クリア対象の判定にも使う）
};

// ---------------- 純粋ロジック ----------------

// 発注NOを解析。日付型（6〜8桁、Wata等の英字接頭辞可）のみ対象。見出し等はnull
function 採番v2_解析_(value) {
  const raw = String(value || '')
    .trim()
    .replace(/[＿]/g, '_')
    .replace(/[\s　]+/g, '')
    .replace(/_+/g, '_');
  if (!raw) return null;
  let m = raw.match(/^([A-Za-z]{0,8}\d{6,8})$/);
  if (m) return { kind: 'date', date: m[1], raw: raw };
  m = raw.match(/^([A-Za-z]{0,8}\d{6,8})_(\d{1,3})$/);
  if (m) {
    return { kind: 'cart', date: m[1], cartNum: Number(m[2]), token: m[2], raw: raw };
  }
  m = raw.match(/^([A-Za-z]{0,8}\d{6,8})_(\d{1,3})_(\d{1,4})$/);
  if (m) {
    return {
      kind: 'full', date: m[1], cartNum: Number(m[2]), token: m[2],
      line: Number(m[3]), raw: raw,
    };
  }
  return null;
}

function 採番v2_カート表記_(n) {
  return n <= 99 ? ('0' + n).slice(-2) : String(n);
}

// rows: [{row, no, vendor}] → { carts, dateMax }
//   carts[date|cartNum] = {date, cartNum, token, vendor, maxLine, lines{}, rows[]}
function 採番v2_状態構築_(rows) {
  const carts = {};
  const dateMax = {};
  rows.forEach(r => {
    const info = 採番v2_解析_(r.no);
    if (!info) return;
    if (info.kind === 'date') return; // 壊れ値。予約なし（修復対象）
    const key = info.date + '|' + info.cartNum;
    if (!carts[key]) {
      carts[key] = {
        date: info.date, cartNum: info.cartNum, token: info.token,
        vendor: '', maxLine: 0, lines: {}, rows: [],
      };
    }
    const cart = carts[key];
    const vendor = String(r.vendor || '').trim();
    if (!cart.vendor && vendor) cart.vendor = vendor;
    cart.rows.push({ row: r.row, info: info, vendor: vendor });
    if (info.kind === 'full') {
      cart.lines[info.line] = true;
      if (info.line > cart.maxLine) cart.maxLine = info.line;
    }
    if (!dateMax[info.date] || dateMax[info.date] < info.cartNum) {
      dateMax[info.date] = info.cartNum;
    }
  });
  return { carts: carts, dateMax: dateMax };
}

function 採番v2_空き行番号_(cart) {
  let n = 1;
  while (cart.lines[n]) n++;
  return n;
}

// onEdit用: 編集行に採番する
//   state: 編集範囲外の行から構築した状態（破壊的に更新する）
//   edited: [{row, no, vendor}] 行順
// → { updates: [{row, value}], notices: [] }
function 採番v2_編集計画_(state, edited) {
  const carts = state.carts;
  const dateMax = state.dateMax;
  const updates = [];
  const notices = [];
  const remap = {};   // 衝突繰り上げ: 元key → 先key（同一バッチ内は同じ先へ）
  let prevDateCart = null; // 直前の日付のみ入力の継続情報

  const allocCart = (date, vendor) => {
    const num = (dateMax[date] || 0) + 1;
    dateMax[date] = num;
    const key = date + '|' + num;
    carts[key] = {
      date: date, cartNum: num, token: 採番v2_カート表記_(num),
      vendor: String(vendor || '').trim(), maxLine: 0, lines: {}, rows: [],
    };
    return carts[key];
  };

  edited.forEach(r => {
    const info = 採番v2_解析_(r.no);
    if (!info) { prevDateCart = null; return; }
    const vendor = String(r.vendor || '').trim();
    let cart = null;
    let line = 0;

    if (info.kind === 'date') {
      // 日付のみ → 連続する同日付・同業者は同じカート
      const canContinue = prevDateCart &&
        prevDateCart.date === info.date &&
        prevDateCart.lastRow === r.row - 1 &&
        (!vendor || !prevDateCart.vendor || prevDateCart.vendor === vendor);
      if (canContinue) {
        cart = carts[prevDateCart.key];
        if (!cart.vendor && vendor) cart.vendor = vendor;
      } else {
        cart = allocCart(info.date, vendor);
      }
      line = 採番v2_空き行番号_(cart);
      prevDateCart = {
        date: info.date, key: cart.date + '|' + cart.cartNum,
        vendor: cart.vendor || vendor, lastRow: r.row,
      };
    } else {
      prevDateCart = null;
      const key = info.date + '|' + info.cartNum;
      cart = carts[key];
      if (!cart) {
        // 未使用カート → そのまま使う
        cart = carts[key] = {
          date: info.date, cartNum: info.cartNum, token: info.token,
          vendor: vendor, maxLine: 0, lines: {}, rows: [],
        };
        if (!dateMax[info.date] || dateMax[info.date] < info.cartNum) {
          dateMax[info.date] = info.cartNum;
        }
      } else if (cart.vendor && vendor && cart.vendor !== vendor) {
        // 別業者が使用中 → 次の空きカートへ繰り上げ
        const from = info.date + '_' + cart.token;
        if (remap[key]) {
          cart = carts[remap[key]];
        } else {
          const moved = allocCart(info.date, vendor);
          remap[key] = moved.date + '|' + moved.cartNum;
          cart = moved;
          notices.push(
            '「' + from + '」は' + carts[key].vendor + 'が使用中 → 「' +
            cart.date + '_' + cart.token + '」を採番');
        }
      } else if (!cart.vendor && vendor) {
        cart.vendor = vendor;
      }
      if (info.kind === 'full' && !cart.lines[info.line]) {
        line = info.line; // 指定行番号が空いていればそのまま
      } else {
        line = 採番v2_空き行番号_(cart);
      }
    }

    cart.lines[line] = true;
    if (line > cart.maxLine) cart.maxLine = line;
    const value = cart.date + '_' + cart.token + '_' + line;
    if (value !== info.raw) updates.push({ row: r.row, value: value });
  });

  return { updates: updates, notices: notices };
}

// 修復用: 全行スキャンして行番号なし行への付番計画を作る
//   rows: [{row, no, vendor, code, qty}] 全データ行（行順）
// → { updates, stale: {旧値→[{newNo, code, qty, remaining}]},
//     mixedKept: [...], dupFull: [...], movedBlocks: n }
function 採番v2_修復計画_(rows) {
  const state = 採番v2_状態構築_(rows);
  const carts = state.carts;
  const dateMax = state.dateMax;
  const byRow = {};
  rows.forEach(r => { byRow[r.row] = r; });

  const updates = [];
  const stale = {};
  const mixedKept = [];
  const dupFull = [];
  let movedBlocks = 0;

  const addStale = (oldRaw, row, newNo) => {
    const src = byRow[row] || {};
    (stale[oldRaw] = stale[oldRaw] || []).push({
      newNo: newNo, code: String(src.code || '').trim(),
      qty: Number(String(src.qty || '').replace(/,/g, '')) || 0,
      remaining: Number(String(src.qty || '').replace(/,/g, '')) || 0,
    });
  };

  // 完全重複の検出（報告のみ）
  const seenFull = {};
  rows.forEach(r => {
    const info = 採番v2_解析_(r.no);
    if (!info || info.kind !== 'full') return;
    if (seenFull[info.raw]) {
      dupFull.push({ row: r.row, no: info.raw, firstRow: seenFull[info.raw] });
    } else {
      seenFull[info.raw] = r.row;
    }
  });

  // カートごとに業者ブロックを分割して修復
  Object.keys(carts).forEach(key => {
    const cart = carts[key];
    // 業者ブロック分割（業者空欄は直前のブロックを継承）
    const blocks = [];
    let cur = null;
    cart.rows.forEach(cr => {
      if (!cur || (cr.vendor && cur.vendor && cr.vendor !== cur.vendor)) {
        cur = { vendor: cr.vendor || (cur ? cur.vendor : ''), rows: [] };
        blocks.push(cur);
      } else if (cr.vendor && !cur.vendor) {
        cur.vendor = cr.vendor;
      }
      cur.rows.push(cr);
    });

    blocks.forEach((block, bi) => {
      const hasFull = block.rows.some(cr => cr.info.kind === 'full');
      if (bi > 0 && hasFull) {
        // 採番済み行を含む別業者ブロック → 触らず報告
        mixedKept.push({
          cart: cart.date + '_' + cart.token, vendor: block.vendor,
          rows: block.rows.map(cr => cr.row),
        });
        return;
      }
      if (bi > 0) {
        // 全行未採番の別業者ブロック → 新カートへ移動
        const num = (dateMax[cart.date] || 0) + 1;
        dateMax[cart.date] = num;
        const token = 採番v2_カート表記_(num);
        movedBlocks++;
        let line = 0;
        block.rows.forEach(cr => {
          line++;
          const value = cart.date + '_' + token + '_' + line;
          updates.push({ row: cr.row, value: value });
          addStale(cr.info.raw, cr.row, value);
        });
        return;
      }
      // 先頭ブロック: 行番号なしの行に空き行番号を付与
      block.rows.forEach(cr => {
        if (cr.info.kind === 'full') return;
        const line = 採番v2_空き行番号_(cart);
        cart.lines[line] = true;
        if (line > cart.maxLine) cart.maxLine = line;
        const value = cart.date + '_' + cart.token + '_' + line;
        updates.push({ row: cr.row, value: value });
        addStale(cr.info.raw, cr.row, value);
      });
    });
  });

  // 日付のみの壊れ値: 連続ブロックごとに新カート
  let dateBlock = null;
  rows.forEach(r => {
    const info = 採番v2_解析_(r.no);
    if (!info || info.kind !== 'date') { dateBlock = null; return; }
    const vendor = String(r.vendor || '').trim();
    const canContinue = dateBlock &&
      dateBlock.date === info.date &&
      dateBlock.lastRow === r.row - 1 &&
      (!vendor || !dateBlock.vendor || dateBlock.vendor === vendor);
    if (!canContinue) {
      const num = (dateMax[info.date] || 0) + 1;
      dateMax[info.date] = num;
      dateBlock = {
        date: info.date, token: 採番v2_カート表記_(num),
        vendor: vendor, line: 0, lastRow: 0,
      };
    }
    dateBlock.line++;
    dateBlock.lastRow = r.row;
    if (vendor && !dateBlock.vendor) dateBlock.vendor = vendor;
    const value = info.date + '_' + dateBlock.token + '_' + dateBlock.line;
    updates.push({ row: r.row, value: value });
    addStale(info.raw, r.row, value);
  });

  return {
    updates: updates, stale: stale, mixedKept: mixedKept,
    dupFull: dupFull, movedBlocks: movedBlocks,
  };
}

// EMS側の古い参照（修復で番号が変わった値）を新番号へ再割当
//   emsRows: [{row, no, code, qty}] / stale: 採番v2_修復計画_の出力
//   codeKeyFn: 商品コード→照合キー配列（GASでは codeKeys_）
// → { updates: [{row, value}], unmatched: [{row, no, code}] }
function 採番v2_EMS再割当計画_(emsRows, stale, codeKeyFn) {
  const updates = [];
  const unmatched = [];
  const keysOf = c => {
    const s = String(c || '').trim();
    if (!s) return [];
    return codeKeyFn(s).map(k => String(k).toUpperCase());
  };
  // 行ごとのコードキャッシュ
  const staleKeys = {};
  Object.keys(stale).forEach(oldRaw => {
    staleKeys[oldRaw] = stale[oldRaw].map(line => ({
      line: line, keys: keysOf(line.code),
    }));
  });

  emsRows.forEach(r => {
    const info = 採番v2_解析_(r.no);
    const raw = info ? info.raw : String(r.no || '').trim();
    const lines = staleKeys[raw];
    if (!lines || !lines.length) return;
    const rowKeys = keysOf(r.code);
    if (!rowKeys.length) {
      unmatched.push({ row: r.row, no: raw, code: String(r.code || '').trim() });
      return;
    }
    const hit = k => k.some(x => rowKeys.indexOf(x) >= 0);
    const cands = lines.filter(l => hit(l.keys));
    if (!cands.length) {
      unmatched.push({ row: r.row, no: raw, code: String(r.code || '').trim() });
      return;
    }
    const qty = Number(String(r.qty || '').replace(/,/g, '')) || 0;
    let picked = cands.find(l => l.line.remaining > 0) || cands[cands.length - 1];
    picked.line.remaining -= qty;
    updates.push({ row: r.row, value: picked.line.newNo });
  });

  return { updates: updates, unmatched: unmatched };
}

// 重複チェック（書き換えなし）
//   orderRows: [{row, no, vendor, code, qty}] / emsRows: [{row, no, code, qty}]
// → { noLine, dupFull, mixedCarts, orphans, overShip, noCode }
function 採番v2_チェック集計_(orderRows, emsRows, codeKeyFn) {
  const keysOf = c => {
    const s = String(c || '').trim();
    if (!s) return [];
    return codeKeyFn(s).map(k => String(k).toUpperCase());
  };
  const toQty = v => Number(String(v || '').replace(/,/g, '')) || 0;

  const noLine = [];
  const dupFull = [];
  const seenFull = {};
  const fullSet = {};
  const lineMap = {}; // 'no|codeKey' → line（同一lineは同オブジェクト共有）

  orderRows.forEach(r => {
    const info = 採番v2_解析_(r.no);
    if (!info) return;
    if (info.kind !== 'full') {
      noLine.push({ row: r.row, no: info.raw });
      return;
    }
    if (seenFull[info.raw]) {
      dupFull.push({ row: r.row, no: info.raw, firstRow: seenFull[info.raw] });
    } else {
      seenFull[info.raw] = r.row;
    }
    fullSet[info.raw] = true;
    const line = { row: r.row, no: info.raw, code: String(r.code || '').trim(), qty: toQty(r.qty), shipped: 0 };
    keysOf(r.code).forEach(k => {
      const key = info.raw + '|' + k;
      if (!lineMap[key]) lineMap[key] = line;
    });
  });

  // 業者混在カート
  const state = 採番v2_状態構築_(orderRows);
  const mixedCarts = [];
  Object.keys(state.carts).forEach(key => {
    const cart = state.carts[key];
    const vendors = [];
    cart.rows.forEach(cr => {
      if (cr.vendor && vendors.indexOf(cr.vendor) < 0) vendors.push(cr.vendor);
    });
    if (vendors.length > 1) {
      mixedCarts.push({
        cart: cart.date + '_' + cart.token, vendors: vendors,
        rows: cart.rows.map(cr => cr.row),
      });
    }
  });

  const orphans = [];
  const noCode = [];
  emsRows.forEach(r => {
    const rawNo = String(r.no || '').trim();
    if (!rawNo) return;
    const code = String(r.code || '').trim();
    if (!code || code.charAt(0) === '★') {
      noCode.push({ row: r.row, no: rawNo });
      return;
    }
    const info = 採番v2_解析_(rawNo);
    if (!info || info.kind !== 'full' || !fullSet[info.raw]) {
      orphans.push({ row: r.row, no: rawNo, code: code });
      return;
    }
    const rowKeys = keysOf(code);
    let line = null;
    for (let i = 0; i < rowKeys.length; i++) {
      if (lineMap[info.raw + '|' + rowKeys[i]]) { line = lineMap[info.raw + '|' + rowKeys[i]]; break; }
    }
    if (!line) {
      orphans.push({ row: r.row, no: rawNo, code: code });
      return;
    }
    line.shipped += toQty(r.qty);
  });

  const overShip = [];
  const seenLine = {};
  Object.keys(lineMap).forEach(key => {
    const line = lineMap[key];
    const id = line.no + '|' + line.row;
    if (seenLine[id]) return;
    seenLine[id] = true;
    if (line.qty > 0 && line.shipped > line.qty) {
      overShip.push({ no: line.no, code: line.code, qty: line.qty, shipped: line.shipped });
    }
  });

  return {
    noLine: noLine, dupFull: dupFull, mixedCarts: mixedCarts,
    orphans: orphans, overShip: overShip, noCode: noCode,
  };
}

// 転送前の購入No検証（送る行の番号が「ちゃんとした状態」かを確認する）
//   picked: [{row, no}] 送ろうとしている行 / allNos: 発注リスト大邱F列の全値
//   opts.checkDup: 大邱シート内での同番号重複もNGにする（大邱→EMS大邱ゲート用）
//   opts.requireExists: 発注リスト大邱に存在する番号のみ許可（EMS大邱→EMSリストゲート用）
// → { bad: [{row, no, reason}] }
function 採番v2_送信前検証_(picked, allNos, opts) {
  opts = opts || {};
  const counts = {};
  const fullSet = {};
  (allNos || []).forEach(v => {
    const info = 採番v2_解析_(v);
    if (info && info.kind === 'full') {
      counts[info.raw] = (counts[info.raw] || 0) + 1;
      fullSet[info.raw] = true;
    }
  });

  const bad = [];
  picked.forEach(p => {
    const s = String(p.no || '').trim();
    if (!s) { bad.push({ row: p.row, no: '(空欄)', reason: '購入Noが空欄' }); return; }
    const info = 採番v2_解析_(s);
    if (!info) { bad.push({ row: p.row, no: s, reason: '形式不正' }); return; }
    if (info.kind !== 'full') {
      bad.push({ row: p.row, no: s, reason: '行番号なし（採番未完了）' });
      return;
    }
    if (opts.checkDup && counts[info.raw] > 1) {
      bad.push({ row: p.row, no: s, reason: 'シート内で同じ番号が' + counts[info.raw] + '行ある' });
      return;
    }
    if (opts.requireExists && !fullSet[info.raw]) {
      bad.push({ row: p.row, no: s, reason: '発注リスト大邱に存在しない' });
    }
  });
  return { bad: bad };
}

function 採番v2_検証結果文面_(bad, limit) {
  const n = limit || 15;
  const lines = bad.slice(0, n).map(b => b.row + '行目 「' + b.no + '」 … ' + b.reason);
  if (bad.length > n) lines.push('…ほか ' + (bad.length - n) + '件');
  return lines.join('\n');
}

// ---------------- GAS結線 ----------------

// 発注リスト大邱データのデータ行範囲のF/H(/K/L)を読み込む
function 採番v2_大邱行読込_(sh, withCodeQty) {
  const start = DAEGU_CFG.HACHU_SRC_DATA_START;
  const lastRow = sh.getLastRow();
  if (lastRow < start) return [];
  const n = lastRow - start + 1;
  const width = withCodeQty ? (採番V2_CFG.COL_QTY - 採番V2_CFG.COL_NO + 1) : (採番V2_CFG.COL_VENDOR - 採番V2_CFG.COL_NO + 1);
  const vals = sh.getRange(start, 採番V2_CFG.COL_NO, n, width).getValues();
  const rows = [];
  for (let i = 0; i < n; i++) {
    rows.push({
      row: start + i,
      no: vals[i][0],
      vendor: vals[i][採番V2_CFG.COL_VENDOR - 採番V2_CFG.COL_NO],
      code: withCodeQty ? vals[i][採番V2_CFG.COL_CODE - 採番V2_CFG.COL_NO] : '',
      qty: withCodeQty ? vals[i][採番V2_CFG.COL_QTY - 採番V2_CFG.COL_NO] : '',
    });
  }
  return rows;
}

// 大邱→EMS大邱 送信ゲート: 送る行の発注NOが全部確定していればnull、NGなら中止文面
function 大邱発注_送信前購入No検証_(src, picked) {
  const all = 採番v2_大邱行読込_(src, false).map(r => r.no);
  const res = 採番v2_送信前検証_(
    picked.map(p => ({ row: p.row, no: p.no })), all, { checkDup: true });
  if (!res.bad.length) return null;
  return '購入No（発注NO）が未確定の行があるため送信を中止しました。\n\n' +
    採番v2_検証結果文面_(res.bad) +
    '\n\n「発注NO：修復」を実行するか、F列を修正してから再実行してください。';
}

// EMS大邱→EMSリスト 送信ゲート: 追記行の購入Noが確定済み＆発注リスト大邱に実在すればnull
function 大邱発注_EMSリスト転送前購入No検証_(ss, rows) {
  const hac = ss.getSheetByName(DAEGU_CFG.HACHU_SRC);
  const all = hac ? 採番v2_大邱行読込_(hac, false).map(r => r.no) : [];
  const res = 採番v2_送信前検証_(
    rows.map(o => ({ row: o.row, no: o.purchaseNo })), all, { requireExists: true });
  if (!res.bad.length) return null;
  return '購入Noが不正な行があるため転送を中止しました（行番号はEMS大邱作業データ）。\n\n' +
    採番v2_検証結果文面_(res.bad) +
    '\n\n「EMS大邱：購入Noを初回補完」または「発注NO：修復」で購入Noを整えてから再実行してください。';
}

// onEdit本体（onEditチェーンの先頭で呼ぶこと）
function 大邱発注_onEdit採番v2_(e) {
  if (!e || !e.range) return;
  const F = 採番V2_CFG.COL_NO;
  const sc = e.range.getColumn(), ec = e.range.getLastColumn();
  if (sc > F || ec < F) return;

  const sh = e.range.getSheet();
  const dataStart = DAEGU_CFG.HACHU_SRC_DATA_START;
  const editStart = Math.max(dataStart, e.range.getRow());
  const editEnd = e.range.getLastRow();
  if (editEnd < dataStart) return;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) {
    SpreadsheetApp.getActive().toast('他の処理が実行中のため採番をスキップしました。「発注NO：修復」で補完できます。');
    return;
  }
  try {
    const all = 採番v2_大邱行読込_(sh, false);
    const outside = all.filter(r => r.row < editStart || r.row > editEnd);
    const edited = all.filter(r => r.row >= editStart && r.row <= editEnd);
    if (!edited.length) return;

    const state = 採番v2_状態構築_(outside);
    const plan = 採番v2_編集計画_(state, edited);

    if (plan.updates.length) {
      const numRows = editEnd - editStart + 1;
      const cells = sh.getRange(editStart, F, numRows, 1);
      const vals = cells.getValues();
      plan.updates.forEach(u => { vals[u.row - editStart][0] = u.value; });
      cells.setValues(vals);
    }
    if (plan.notices.length) {
      SpreadsheetApp.getActive().toast(plan.notices.join('\n'), '発注NO採番', 8);
    }
  } finally {
    lock.releaseLock();
  }

  // 装飾は編集近傍のみ（全域再描画はしない）
  const decoStart = Math.max(dataStart, editStart - 2);
  const decoRows = editEnd - decoStart + 5;
  大邱発注_チェックボックスを付ける_(sh, decoStart, decoRows);
  大邱発注_格子罫線_(sh, decoStart, decoRows);
  大邱発注_グループ罫線_近傍_(sh, editStart, editEnd - editStart + 1);
}

// 編集近傍のカートグループだけ太枠を引き直す
function 大邱発注_グループ罫線_近傍_(sh, startRow, numRows) {
  const dataStart = DAEGU_CFG.HACHU_SRC_DATA_START;
  const lastRow = sh.getLastRow();
  if (lastRow < dataStart) return;

  const PAD = 120; // 編集位置の前後をこの行数まで見てグループ境界を探す
  let from = Math.max(dataStart, startRow - PAD);
  let to = Math.min(lastRow, startRow + numRows - 1 + PAD);
  const vals = sh.getRange(from, 採番V2_CFG.COL_NO, to - from + 1, 1).getDisplayValues();

  const baseOf = v => {
    const info = 採番v2_解析_(v);
    if (!info || info.kind === 'date') return '';
    return info.date + '|' + info.cartNum;
  };

  // 編集範囲が属するグループの先頭・末尾までスキャン範囲を縮める
  const editFromIdx = Math.max(0, startRow - from);
  const editToIdx = Math.min(vals.length - 1, startRow + numRows - 1 - from);
  let a = editFromIdx;
  const startBase = baseOf(vals[editFromIdx][0]);
  while (a > 0 && startBase && baseOf(vals[a - 1][0]) === startBase) a--;
  let b = editToIdx;
  const endBase = baseOf(vals[editToIdx][0]);
  while (b < vals.length - 1 && endBase && baseOf(vals[b + 1][0]) === endBase) b++;

  const maxCol = 大邱発注_罫線列数_(sh);
  let gStart = null, cur = '';
  for (let i = a; i <= b + 1; i++) {
    const isEnd = i > b;
    const key = isEnd ? '' : baseOf(vals[i][0]);
    if (gStart === null) {
      if (!isEnd && key !== '') { gStart = i; cur = key; }
      continue;
    }
    if (isEnd || key !== cur) {
      sh.getRange(from + gStart, 1, i - gStart, maxCol)
        .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
      if (!isEnd && key !== '') { gStart = i; cur = key; }
      else { gStart = null; cur = ''; }
    }
  }
}

// EMSリストのヘッダー行（A列='No.'）を探す。見つからなければ-1
function 採番v2_EMSリストヘッダー_(vals) {
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === 'No.') return i;
  }
  return -1;
}

// メニュー: 発注NO修復（行番号付与→EMS側へ反映）
function 大邱発注_発注NO修復() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const sh = ss.getSheetByName(DAEGU_CFG.HACHU_SRC);
  const emsDaegu = ss.getSheetByName(DAEGU_CFG.EMS_SRC);
  const emsList = ss.getSheetByName(DAEGU_CFG.EMS_DST);
  if (!sh) { ui.alert('「' + DAEGU_CFG.HACHU_SRC + '」が見つかりません。'); return; }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) { ss.toast('他の処理が実行中です。少し待って再実行してください。'); return; }
  try {
    const rows = 採番v2_大邱行読込_(sh, true);
    const plan = 採番v2_修復計画_(rows);

    if (!plan.updates.length && !plan.mixedKept.length && !plan.dupFull.length) {
      ui.alert('発注NO修復', '修復対象はありません（全行採番済み）。', ui.ButtonSet.OK);
      return;
    }

    const samples = plan.updates.slice(0, 10)
      .map(u => u.row + '行目 → ' + u.value);
    const lines = [
      '行番号を付与: ' + plan.updates.length + '件' +
        (plan.movedBlocks ? '（うち別業者かぶりで新カートへ移動: ' + plan.movedBlocks + 'ブロック）' : ''),
      '採番済みで触らない業者混在: ' + plan.mixedKept.length + '件',
      '完全重複（手動確認）: ' + plan.dupFull.length + '件',
      '',
    ].concat(samples, plan.updates.length > 10 ? ['…ほか'] : []);
    lines.push('', 'F列を書き換え、EMS大邱・EMSリストの古い参照も追随させます。実行する？');
    const res = ui.alert('発注NO修復', lines.join('\n'), ui.ButtonSet.YES_NO);
    if (res !== ui.Button.YES) { ss.toast('やめました。'); return; }

    // 1) 大邱F列へ書き込み
    if (plan.updates.length) {
      const start = DAEGU_CFG.HACHU_SRC_DATA_START;
      const n = sh.getLastRow() - start + 1;
      const cells = sh.getRange(start, 採番V2_CFG.COL_NO, n, 1);
      const vals = cells.getValues();
      plan.updates.forEach(u => { vals[u.row - start][0] = u.value; });
      cells.setValues(vals);
    }

    // 2) EMS大邱 T列(購入No) の古い参照を再割当（コード=H列, 数量=I列）
    let emsDaeguCount = 0;
    const unmatched = [];
    if (emsDaegu && Object.keys(plan.stale).length) {
      const eVals = emsDaegu.getDataRange().getValues();
      const emsRows = [];
      for (let i = 0; i < eVals.length; i++) {
        emsRows.push({ row: i + 1, no: eVals[i][19], code: eVals[i][7], qty: eVals[i][8] });
      }
      const st = JSON.parse(JSON.stringify(plan.stale)); // remaining消費を独立させる
      const emsPlan = 採番v2_EMS再割当計画_(emsRows, st, codeKeys_);
      emsPlan.updates.forEach(u => {
        emsDaegu.getRange(u.row, 20).setValue(u.value);
      });
      emsDaeguCount = emsPlan.updates.length;
      emsPlan.unmatched.forEach(u => unmatched.push('EMS大邱 ' + u.row + '行目 ' + u.no + ' / コード「' + (u.code || '空欄') + '」'));
    }

    // 3) EMSリスト F列(購入No) の古い参照を再割当（コード=I列, 数量=J列）
    let emsListCount = 0;
    if (emsList && Object.keys(plan.stale).length) {
      const lVals = emsList.getDataRange().getValues();
      const h = 採番v2_EMSリストヘッダー_(lVals);
      if (h >= 0) {
        const emsRows = [];
        for (let i = h + 1; i < lVals.length; i++) {
          emsRows.push({ row: i + 1, no: lVals[i][5], code: lVals[i][8], qty: lVals[i][9] });
        }
        const st = JSON.parse(JSON.stringify(plan.stale));
        const listPlan = 採番v2_EMS再割当計画_(emsRows, st, codeKeys_);
        listPlan.updates.forEach(u => {
          emsList.getRange(u.row, 6).setValue(u.value);
        });
        emsListCount = listPlan.updates.length;
        listPlan.unmatched.forEach(u => unmatched.push('EMSリスト ' + u.row + '行目 ' + u.no + ' / コード「' + (u.code || '空欄') + '」'));
      }
    }
    SpreadsheetApp.flush();

    // 4) 既存の一括修正・補完で仕上げ
    let tail = '';
    if (typeof 大邱_発注EMS枝番を一括修正 === 'function') {
      const r = 大邱_発注EMS枝番を一括修正(true);
      tail += '\n枝番一括修正: 発注' + r.order + '件 / EMS大邱' + r.ems + '件';
    }
    if (typeof EMSリスト_購入No自動補完 === 'function') {
      const filled = EMSリスト_購入No自動補完(true);
      tail += '\nEMSリスト購入No補完: ' + filled + '件';
    }

    // 修復で解消した箇所の赤背景を自動リセット（残った課題だけ赤が残る）
    大邱発注_発注NOフラグ更新_();

    const report = [
      '発注NO修復が完了しました。',
      '大邱F列: ' + plan.updates.length + '件 / EMS大邱T列: ' + emsDaeguCount + '件 / EMSリストF列: ' + emsListCount + '件' + tail,
      '赤背景を最新状態に更新しました（残っている赤はまだ要対応の箇所です）。',
    ];
    if (unmatched.length) {
      report.push('', '【要手動】コードが照合できず旧番号のまま（コードを入れて再実行）:');
      report.push(unmatched.slice(0, 15).join('\n'));
      if (unmatched.length > 15) report.push('…ほか ' + (unmatched.length - 15) + '件');
    }
    if (plan.mixedKept.length) {
      report.push('', '【要手動】採番済みの業者混在カート:');
      plan.mixedKept.slice(0, 10).forEach(m => {
        report.push(m.cart + ' の ' + m.vendor + '（' + m.rows.join(',') + '行目）');
      });
    }
    if (plan.dupFull.length) {
      report.push('', '【要手動】完全に同じ番号:');
      plan.dupFull.slice(0, 10).forEach(d => {
        report.push(d.no + '（' + d.firstRow + '行目と' + d.row + '行目）');
      });
    }
    ui.alert('発注NO修復', report.join('\n'), ui.ButtonSet.OK);
  } finally {
    lock.releaseLock();
  }
}

// 大邱F列・EMSリストF列の赤フラグを実データから塗り直す
//   解消済みの赤（このチェックが塗ったFLAG_BGのみ）は自動でリセットされる。
//   手動で付けた他の色には触らない。→ 集計結果を返す
function 大邱発注_発注NOフラグ更新_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(DAEGU_CFG.HACHU_SRC);
  const emsList = ss.getSheetByName(DAEGU_CFG.EMS_DST);
  if (!sh) return null;

  const orderRows = 採番v2_大邱行読込_(sh, true);
  const emsRows = [];
  let emsHeaderRow = -1;
  if (emsList) {
    const lVals = emsList.getDataRange().getValues();
    const h = 採番v2_EMSリストヘッダー_(lVals);
    emsHeaderRow = h;
    if (h >= 0) {
      for (let i = h + 1; i < lVals.length; i++) {
        emsRows.push({ row: i + 1, no: lVals[i][5], code: lVals[i][8], qty: lVals[i][9] });
      }
    }
  }

  const rep = 採番v2_チェック集計_(orderRows, emsRows, codeKeys_);

  const paint = (sheet, col, startRow, lastRow, flagRows) => {
    if (!sheet || lastRow < startRow) return;
    const range = sheet.getRange(startRow, col, lastRow - startRow + 1, 1);
    const bgs = range.getBackgrounds();
    let dirty = false;
    for (let i = 0; i < bgs.length; i++) {
      const want = flagRows[startRow + i] ? 採番V2_CFG.FLAG_BG : null;
      const curIsFlag = bgs[i][0] === 採番V2_CFG.FLAG_BG;
      if (want && !curIsFlag) { bgs[i][0] = 採番V2_CFG.FLAG_BG; dirty = true; }
      else if (!want && curIsFlag) { bgs[i][0] = null; dirty = true; }
    }
    if (dirty) range.setBackgrounds(bgs);
  };

  const orderFlags = {};
  rep.noLine.forEach(x => { orderFlags[x.row] = true; });
  rep.dupFull.forEach(x => { orderFlags[x.row] = true; orderFlags[x.firstRow] = true; });
  rep.mixedCarts.forEach(m => m.rows.forEach(r => { orderFlags[r] = true; }));
  paint(sh, 採番V2_CFG.COL_NO, DAEGU_CFG.HACHU_SRC_DATA_START, sh.getLastRow(), orderFlags);

  const emsFlags = {};
  rep.orphans.forEach(x => { emsFlags[x.row] = true; });
  if (emsList && emsHeaderRow >= 0) {
    paint(emsList, 6, emsHeaderRow + 2, emsList.getLastRow(), emsFlags);
  }
  return rep;
}

// メニュー: 発注NO重複チェック（書き換えなし・赤背景は最新状態に更新）
function 大邱発注_発注NO重複チェック() {
  const ui = SpreadsheetApp.getUi();
  const rep = 大邱発注_発注NOフラグ更新_();
  if (!rep) { ui.alert('「' + DAEGU_CFG.HACHU_SRC + '」が見つかりません。'); return; }

  const lines = ['【発注リスト大邱データ】'];
  lines.push('行番号なし: ' + rep.noLine.length + '件' +
    (rep.noLine.length ? '（例: ' + rep.noLine.slice(0, 5).map(x => x.row + '行 ' + x.no).join(' / ') + '）' : ''));
  lines.push('完全重複: ' + rep.dupFull.length + '件');
  lines.push('業者混在カート: ' + rep.mixedCarts.length + '件' +
    (rep.mixedCarts.length ? '（' + rep.mixedCarts.slice(0, 5).map(m => m.cart + ': ' + m.vendors.join('と')).join(' / ') + '）' : ''));
  lines.push('', '【EMSリスト】');
  lines.push('迷子キー（発注に無い購入No）: ' + rep.orphans.length + '件' +
    (rep.orphans.length ? '（例: ' + rep.orphans.slice(0, 5).map(x => x.row + '行 ' + x.no).join(' / ') + '）' : ''));
  lines.push('商品コード未入力（★コピペ等）: ' + rep.noCode.length + '件' +
    (rep.noCode.length ? '（例: ' + rep.noCode.slice(0, 5).map(x => x.row + '行 ' + x.no).join(' / ') + '）' : ''));
  lines.push('送りすぎ（発送数合計＞発注数）: ' + rep.overShip.length + '件');
  rep.overShip.slice(0, 8).forEach(o => {
    lines.push('  ' + o.no + ' ' + o.code + ' 発注' + o.qty + ' に対し発送' + o.shipped);
  });
  lines.push('', '※分割発送（同じ照合キーが複数回・合計が発注数以内）は正常扱いです。');
  lines.push('※赤背景は該当セルに付けました（再実行で最新化）。');

  ui.alert('発注NO重複チェック', lines.join('\n'), ui.ButtonSet.OK);
}
