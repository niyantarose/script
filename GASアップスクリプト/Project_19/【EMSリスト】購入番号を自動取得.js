// ====== EMSリスト：入荷日＋商品コードから購入No.を自動補完 ======
function EMSリスト_購入No自動補完(silent) {
  const ss = SpreadsheetApp.getActive();
  const ems = ss.getSheetByName('EMSリスト');
  const hachu = ss.getSheetByName('発注');
  if (!ems || !hachu) return 0;

  // 日付を yyyy-MM-dd に正規化（"26/05/20(水)" みたいな文字列にも対応）
  const fmtDate = v => {
    if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
    const s = String(v || '').trim().replace(/\(.+?\)/, '');
    const m = s.match(/^(\d{2,4})\/(\d{1,2})\/(\d{1,2})$/);
    if (m) {
      let y = Number(m[1]); if (y < 100) y += 2000;
      return y + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
    }
    return s;
  };

  // ---- 発注側：入荷日＋商品コード → 購入No. の対応表 ----
  const hVals = hachu.getDataRange().getValues();
  let hHeader = -1;
  for (let i = 0; i < hVals.length; i++) {
    if (String(hVals[i][2]).trim() === 'DorderDate') { hHeader = i; break; }
  }
  if (hHeader < 0) return 0;

  const addOrderNo = (map, key, orderNo) => {
    (map[key] = map[key] || new Set()).add(orderNo);
  };
  const mergeSets = (map, keys) => {
    const merged = new Set();
    keys.forEach(key => {
      const set = map[key];
      if (set) set.forEach(v => merged.add(v));
    });
    return merged.size ? merged : null;
  };

  const orderMap = {};        // 'yyyy-MM-dd|CODE' -> Set(購入No.)
  const orderMapByCodeQty = {}; // 'CODE|数量' -> Set(購入No.)
  const orderMapByCode = {};    // 'CODE' -> Set(購入No.)
  for (let i = hHeader + 1; i < hVals.length; i++) {
    const code = normCode_(hVals[i][11]);            // L列 商品コード
    const arrival = hVals[i][3];                     // D列 入荷日
    const orderNo = String(hVals[i][6] || '').trim();// G列 購入No.
    const qty = String(hVals[i][12] || '').trim();    // M列 数量
    if (!code || !orderNo) continue;
    codeKeys_(code).forEach(k => {
      if (arrival !== '' && arrival !== null) addOrderNo(orderMap, fmtDate(arrival) + '|' + k, orderNo);
      if (qty) addOrderNo(orderMapByCodeQty, k + '|' + qty, orderNo);
      addOrderNo(orderMapByCode, k, orderNo);
    });
  }

  // ---- EMSリスト側：購入No.が空欄の行だけ埋める ----
  const eVals = ems.getDataRange().getValues();
  let eHeader = -1;
  for (let i = 0; i < eVals.length; i++) {
    if (String(eVals[i][0]).trim() === 'No.') { eHeader = i; break; }
  }
  if (eHeader < 0) return 0;

  let filled = 0;
  const ambiguous = [];
  for (let i = eHeader + 1; i < eVals.length; i++) {
    if (String(eVals[i][5] || '').trim() !== '') continue;  // F列に既に値があれば触らない
    const code = normCode_(eVals[i][8]);                    // I列 商品コード
    const qty = String(eVals[i][9] || '').trim();            // J列 数量
    const arrival = eVals[i][1];                            // B列 入荷日
    if (!code) continue;

    const keys = codeKeys_(code);
    let set = null;
    if (arrival !== '' && arrival !== null) {
      set = mergeSets(orderMap, keys.map(k => fmtDate(arrival) + '|' + k));
    }
    if ((!set || set.size === 0) && qty) {
      set = mergeSets(orderMapByCodeQty, keys.map(k => k + '|' + qty));
    }
    if (!set || set.size === 0) {
      set = mergeSets(orderMapByCode, keys);
    }
    if (!set || set.size === 0) continue;                   // 発注側に該当なし
    if (set.size > 1) {                                     // 候補が複数 → 安全のためスキップ
      ambiguous.push(`${i + 1}行目 ${code}（候補: ${[...set].join(', ')}）`);
      continue;
    }
    ems.getRange(i + 1, 6).setValue([...set][0]);           // F列に購入No.
    filled++;
  }

  const msg = `購入No.補完: ${filled}件` +
    (ambiguous.length ? `\n⚠ 候補が複数で保留: \n${ambiguous.join('\n')}` : '');
  if (silent) Logger.log(msg);
  else SpreadsheetApp.getActive().toast(msg);
  return filled;
}
