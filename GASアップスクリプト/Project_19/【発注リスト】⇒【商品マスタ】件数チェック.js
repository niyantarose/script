function 件数チェック() {
  const ss = SpreadsheetApp.getActive();
  const hachu = ss.getSheetByName(CFG.HACHU_SHEET);
  const master = ss.getSheetByName(CFG.MASTER_SHEET);
  const ui = SpreadsheetApp.getUi();

  // ---- 発注シート読み込み ----
  const hLast = hachu.getLastRow();
  const hStart = CFG.HACHU_HEADER_ROW + 1;
  const hRows = hLast >= hStart
    ? hachu.getRange(hStart, 1, hLast - hStart + 1, hachu.getLastColumn()).getValues()
    : [];

  // 発注の「商品コード」ユニーク集合（行番号も覚えとく）
  const hachuKeys = new Map();   // key -> {code,vendor,name,row}
  let hachuLineCount = 0;        // コードが入ってる実データ行数
  hRows.forEach((r, i) => {
    const code = String(r[CFG.HACHU_CODE - 1] || '').trim();
    if (!code) return;
    hachuLineCount++;
    const vendor = String(r[CFG.HACHU_VENDOR - 1] || '').trim();
    const name = String(r[CFG.HACHU_NAME - 1] || '').trim();
    const key = _masterPrimaryCodeKey_(code);
    if (!key) return;
    if (!hachuKeys.has(key)) {
      hachuKeys.set(key, { code, vendor, name, row: hStart + i });
    }
  });

  // ---- マスタ読み込み ----
  const mLast = master.getLastRow();
  const mStart = CFG.MASTER_HEADER_ROW + 1;
  const mRows = mLast >= mStart
    ? master.getRange(mStart, 1, mLast - mStart + 1, master.getLastColumn()).getValues()
    : [];

  const masterKeys = new Set();
  let masterLineCount = 0;
  mRows.forEach(r => {
    const code = String(r[CFG.M_CODE - 1] || '').trim();
    if (!code) return;
    masterLineCount++;
    const key = _masterPrimaryCodeKey_(code);
    if (key) masterKeys.add(key);
  });

  // ---- 突き合わせ ----
  const onlyInHachu = [];   // 発注にあってマスタに無い（取りこぼし）
  hachuKeys.forEach((v, key) => {
    if (!masterKeys.has(key)) onlyInHachu.push(v);
  });

  const onlyInMaster = [];  // マスタにあって発注に無い（幽霊行）
  mRows.forEach(r => {
    const code = String(r[CFG.M_CODE - 1] || '').trim();
    if (!code) return;
    const vendor = String(r[CFG.M_VENDOR - 1] || '').trim();
    const key = _masterPrimaryCodeKey_(code);
    if (!hachuKeys.has(key)) {
      onlyInMaster.push({ code, vendor, name: String(r[CFG.M_NAME - 1] || '').trim() });
    }
  });

  // ---- 結果まとめ ----
  let msg = '';
  msg += `発注 実データ行   : ${hachuLineCount}\n`;
  msg += `発注 ユニーク数   : ${hachuKeys.size}（商品コード）\n`;
  msg += `マスタ 件数       : ${masterLineCount}\n`;
  msg += `重複でスキップ    : ${hachuLineCount - hachuKeys.size} 行\n`;
  msg += `─────────────\n`;
  msg += `★発注にあってマスタに無い（取りこぼし）: ${onlyInHachu.length} 件\n`;
  onlyInHachu.forEach(v => {
    msg += `  ・${v.code} / ${v.vendor} / ${v.name}（発注${v.row}行目）\n`;
  });
  msg += `\n△マスタにあって発注に無い: ${onlyInMaster.length} 件\n`;
  onlyInMaster.forEach(v => {
    msg += `  ・${v.code} / ${v.vendor} / ${v.name}\n`;
  });

  Logger.log(msg);
  ui.alert('件数チェック結果', msg, ui.ButtonSet.OK);
}
