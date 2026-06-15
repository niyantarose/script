function 発注_商品マスタから一括データ補完() {
  const ss = SpreadsheetApp.getActive();
  const hachu = ss.getSheetByName(CFG.HACHU_SHEET);
  const ui = SpreadsheetApp.getUi();
  if (!hachu) { ui.alert('発注シートが見つからんで'); return; }

  const start = CFG.HACHU_HEADER_ROW + 1;
  const last  = hachu.getLastRow();
  if (last < start) { ui.alert('発注データが無いで'); return; }
  const n = last - CFG.HACHU_HEADER_ROW;

  const { info, byKey, byKeyW, byCode } = _masterMap();

  const get = c => hachu.getRange(start, c, n, 1).getValues();
  const codeCol   = get(CFG.HACHU_CODE);
  const vendorCol = get(CFG.HACHU_VENDOR);
  const nameCol   = get(CFG.HACHU_NAME);
  const itemCol   = get(CFG.HACHU_ITEM);
  const priceCol  = get(CFG.HACHU_PRICE);
  const weightCol = get(CFG.HACHU_WEIGHT);

  const blank = v => String(v ?? '').trim() === '';
  let filled = 0, noMatch = 0;

  for (let i = 0; i < n; i++) {
    const code = String(codeCol[i][0] || '').trim();
    if (!code) continue;
    if (!info[code] && !byCode[code]) { noMatch++; continue; }  // マスタに無いコード

    let vendor = String(vendorCol[i][0] || '').trim();
    if (!vendor && byCode[code] && byCode[code].length === 1) {
      vendor = String(byCode[code][0].vendor || '').trim();
      if (vendor) { vendorCol[i][0] = vendor; filled++; }
    }

    if (info[code]) {
      if (blank(nameCol[i][0]) && info[code].name !== '') { nameCol[i][0] = info[code].name; filled++; }
      if (blank(itemCol[i][0]) && info[code].item !== '') { itemCol[i][0] = info[code].item; filled++; }
    }

    let price = null, weight = null;
    if (vendor && byKey[code + CFG.SEP + vendor] !== undefined) {
      price  = byKey[code + CFG.SEP + vendor];
      weight = byKeyW[code + CFG.SEP + vendor];
    } else if (byCode[code] && byCode[code].length === 1) {
      price  = byCode[code][0].price;
      weight = byCode[code][0].weight;
    }
    if (price  !== null && price  !== '' && blank(priceCol[i][0]))  { priceCol[i][0]  = price;  filled++; }
    if (weight !== null && weight !== '' && typeof weight !== 'undefined' && blank(weightCol[i][0])) {
      weightCol[i][0] = weight; filled++;
    }
  }

  // 触った列だけ書き戻す(関数列は触らへん)
  hachu.getRange(start, CFG.HACHU_VENDOR, n, 1).setValues(vendorCol);
  hachu.getRange(start, CFG.HACHU_NAME,   n, 1).setValues(nameCol);
  hachu.getRange(start, CFG.HACHU_ITEM,   n, 1).setValues(itemCol);
  hachu.getRange(start, CFG.HACHU_PRICE,  n, 1).setValues(priceCol);
  hachu.getRange(start, CFG.HACHU_WEIGHT, n, 1).setValues(weightCol);

  ui.alert(`発注を商品マスタから補完したで\n補完セル：${filled}\nマスタに無いコード行：${noMatch}`);
}