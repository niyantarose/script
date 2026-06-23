function チェックを入れた行を商品マスタへ登録(){
  const ss = SpreadsheetApp.getActive();
  const hachu = ss.getSheetByName(CFG.HACHU_SHEET);
  const master = ss.getSheetByName(CFG.MASTER_SHEET);
  const ui = SpreadsheetApp.getUi();

  const last = hachu.getLastRow();
  if (last <= CFG.HACHU_HEADER_ROW){ ui.alert('発注データが無いで'); return; }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)){ ui.alert('今ほかの処理が動いてるみたい。少し待ってもう一回。'); return; }

  try {
    const base = CFG.HACHU_HEADER_ROW + 1;
    const n = last - CFG.HACHU_HEADER_ROW;
    const chkCol = (typeof HATCHU_CFG!=='undefined' && HATCHU_CFG.COL_CHK) ? HATCHU_CFG.COL_CHK : 1;

    const rows = hachu.getRange(base, 1, n, hachu.getLastColumn()).getValues();
    const checkVals = hachu.getRange(base, chkCol, n, 1).getValues();

    const isChecked = v => v===true || v==='TRUE' || v==='true' || v===1;

    const candidates = [];   // [code, vendor, name, item, price, weight] ← No.は書かへん
    const newChecks = checkVals.map(r => r.slice());
    const incomplete = [];

    for (let i=0;i<n;i++){
      if (!isChecked(checkVals[i][0])) continue;

      const code   = String(rows[i][CFG.HACHU_CODE   - 1] || '').trim();
      const vendor = String(rows[i][CFG.HACHU_VENDOR - 1] || '').trim();
      const name   = String(rows[i][CFG.HACHU_NAME   - 1] || '').trim();
      const item   = String(rows[i][CFG.HACHU_ITEM   - 1] || '').trim();
      const price  = rows[i][CFG.HACHU_PRICE  - 1];
      const weight = rows[i][CFG.HACHU_WEIGHT - 1];   // ★重さ

      if (!code){ incomplete.push(base + i); continue; }

      candidates.push([code, vendor, name, item, price, weight]);   // ★6列
      newChecks[i][0] = false;
    }

    const result = 商品マスタ_候補行を商品コードで反映_(master, candidates);

    hachu.getRange(base, chkCol, n, 1).setValues(newChecks);

    let msg = `チェック行を登録\n追加 ${result.added} / 更新行 ${result.updatedRows} / 更新セル ${result.updatedCells} / 統合 ${result.uniqueGroups}`;
    if (incomplete.length){
      msg += `\n\n⚠ 商品コードが空で登録できんかった行(チェックは残してある):\n  ${incomplete.join(', ')} 行目`;
    }
    ui.alert(msg);

  } finally {
    lock.releaseLock();
  }
}

// 指定フィールドが違うときだけ更新。更新したら true（価格・重さ共用）
function _updateMasterFieldIfDiff_(master, code, vendor, fieldCol, value){
  const base = CFG.MASTER_HEADER_ROW + 1;
  const last = master.getLastRow();
  const n = last - CFG.MASTER_HEADER_ROW;
  if (n <= 0) return false;
  const v = master.getRange(base, 1, n, master.getLastColumn()).getValues();
  const targetKey = _masterPrimaryCodeKey_(code);
  for (let i=0;i<n;i++){
    if (_masterPrimaryCodeKey_(v[i][CFG.M_CODE-1]) === targetKey){
      if (v[i][fieldCol-1] !== value){
        master.getRange(base+i, fieldCol).setValue(value);
        return true;
      }
      return false;
    }
  }
  return false;
}

// 5分おきに取りこぼし自動回収。1回だけ実行して仕込む。
function setupマスタ補完トリガー(){
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'マスタ補完_定期') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('マスタ補完_定期').timeBased().everyMinutes(5).create();
  SpreadsheetApp.getActive().toast('5分おきのマスタ補完トリガーを仕込んだで');
}
function マスタ補完_定期(){ マスタ自動補完_(true); }   // silent
