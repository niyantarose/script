function 診断_EMS照合キー(code) {
  code = normCode_(code || 'ZENCHIGD54');
  const ss = SpreadsheetApp.getActive();
  const ems = ss.getSheetByName('EMSリスト');
  const src = ss.getSheetByName('EMS同期データ');
  const lines = ['【照合対象】' + code, ''];

  // EMSリスト
  lines.push('=== EMSリスト ===');
  const ev = ems.getDataRange().getValues();
  let eh = -1;
  for (let i = 0; i < ev.length; i++) if (String(ev[i][0]).trim() === 'No.') { eh = i; break; }
  for (let i = eh + 1; i < ev.length; i++) {
    if (normCode_(ev[i][8]) !== code) continue;
    const arr = _emsFmtDate_(ev[i][1]), qty = String(ev[i][9]);
    lines.push(`行${i+1}  入荷[${arr}] 数量[${qty}] EMS番号[${normTrack_(ev[i][12])||'空'}] 発送日[${ev[i][2]||'空'}]  → key=「${arr}|${code}|${qty}」`);
  }

  // EMS同期データ
  lines.push('', '=== EMS同期データ ===');
  const sv = src.getDataRange().getValues();
  let sh = -1;
  for (let i = 0; i < sv.length; i++) if (String(sv[i][0]).trim() === '入荷') { sh = i; break; }
  for (let i = sh + 1; i < sv.length; i++) {
    if (normCode_(sv[i][7]) !== code) continue;
    const arr = _emsFmtDate_(sv[i][0]), qty = String(sv[i][8]);
    lines.push(`行${i+1}  入荷[${arr}] 数量[${qty}] track[${normTrack_(sv[i][3])||'空'}] 出発[${sv[i][1]||'空'}]  → key=「${arr}|${code}|${qty}」`);
  }

  const msg = lines.join('\n');
  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}

// ============================================================
// 🩺 商品マスタ 健康診断（エディタから実行して実行ログを見る）
//   ・マスタの行数/重複コード/欠損(重さ・品目・価格・商品名)
//   ・発注リスト大邱データとの突合(マスタ未登録コード)
//   ・重量がどこにも無い商品(大邱行・マスタ・EMS大邱実績の全部で空)
// ============================================================
function 商品マスタ_健康診断() {
  const ss = SpreadsheetApp.getActive();
  const L = [];
  const master = ss.getSheetByName(CFG.MASTER_SHEET);
  const daegu = ss.getSheetByName(DAEGU_CFG.HACHU_SRC);
  const ems = ss.getSheetByName(DAEGU_CFG.EMS_SRC);
  if (!master) { Logger.log('商品マスタが見つからない'); return; }
  const has = v => v !== '' && v != null && String(v).trim() !== '';

  // ---- マスタ本体 ----
  const mStart = CFG.MASTER_HEADER_ROW + 1, mLast = master.getLastRow();
  const mRows = mLast >= mStart
    ? master.getRange(mStart, 1, mLast - mStart + 1, Math.max(master.getLastColumn(), 8)).getValues()
    : [];
  let total = 0, noWeight = 0, noItem = 0, noPrice = 0, noName = 0;
  const keyCount = {}, masterHasW = {};
  mRows.forEach(r => {
    const code = String(r[CFG.M_CODE - 1] || '').trim();
    if (!code) return;
    total++;
    if (!has(r[CFG.M_WEIGHT - 1])) noWeight++;
    if (!has(r[CFG.M_ITEM - 1])) noItem++;
    if (!has(r[CFG.M_PRICE - 1])) noPrice++;
    if (!has(r[CFG.M_NAME - 1])) noName++;
    const k = _masterPrimaryCodeKey_(code);
    keyCount[k] = (keyCount[k] || 0) + 1;
    if (has(r[CFG.M_WEIGHT - 1])) masterHasW[k] = true;
  });
  const dup = Object.keys(keyCount).filter(k => keyCount[k] > 1);

  L.push('===== 商品マスタ 健康診断 =====');
  L.push('マスタ行数(コードあり): ' + total + ' / ユニークコード: ' + Object.keys(keyCount).length);
  L.push('重複コード: ' + dup.length + '種' + (dup.length ? '  例: ' + dup.slice(0, 10).join(', ') : ''));
  L.push('欠損: 重さ空 ' + noWeight + ' / 品目空 ' + noItem + ' / 価格空 ' + noPrice + ' / 商品名空 ' + noName);

  // ---- EMS大邱の発送実績(重量あり)コード ----
  const emsHasW = {};
  if (ems && ems.getLastRow() >= 3) {
    const en = ems.getLastRow() - 2;
    const ev = ems.getRange(3, 8, en, 5).getValues(); // H=コード .. L=weight
    for (let i = 0; i < en; i++) {
      const c = normCode_(ev[i][0]);
      if (c && has(ev[i][4])) codeKeys_(c).forEach(k => { emsHasW[k] = true; });
    }
  }

  // ---- 発注リスト大邱データとの突合 ----
  if (daegu) {
    const dStart = 6, dLast = daegu.getLastRow();
    const dRows = dLast >= dStart ? daegu.getRange(dStart, 1, dLast - dStart + 1, 15).getValues() : [];
    const dKeys = {};
    dRows.forEach(r => {
      const code = String(r[10] || '').trim(); // K 商品コード
      if (!code) return;
      const k = _masterPrimaryCodeKey_(code);
      if (!dKeys[k]) dKeys[k] = { code: code, cnt: 0, wRows: 0 };
      dKeys[k].cnt++;
      if (has(r[14])) dKeys[k].wRows++; // O weight
    });
    const dAll = Object.keys(dKeys);
    const notInMaster = dAll.filter(k => !(k in keyCount));
    const noWeightAnywhere = dAll.filter(k =>
      dKeys[k].wRows === 0 && !masterHasW[k] && !emsHasW[k]);

    L.push('');
    L.push('===== 発注リスト大邱データとの突合 =====');
    L.push('大邱のユニークコード: ' + dAll.length);
    L.push('マスタ未登録: ' + notInMaster.length + '種' +
      (notInMaster.length ? '  例: ' + notInMaster.slice(0, 10).map(k => dKeys[k].code).join(', ') : ''));
    L.push('重量がどこにも無い(大邱行・マスタ・EMS大邱実績すべて空): ' + noWeightAnywhere.length + '種');
    if (noWeightAnywhere.length) {
      L.push('  → ' + noWeightAnywhere.slice(0, 20).map(k => dKeys[k].code).join(', ') +
        (noWeightAnywhere.length > 20 ? ' …ほか' : ''));
    }
  }

  const msg = L.join('\n');
  Logger.log(msg);
  ss.toast('商品マスタ健康診断 完了（実行ログ参照）', '🩺 診断', 5);
}

// ==== 一時デバッグ: 小杉山行がマスタ更新されない原因調査(調査後削除OK) ====
function DEBUG_小杉山チェック() {
  const sourceCfg = DAEGU_HACHU_MASTER_CFG;
  const ss = SpreadsheetApp.getActive();
  const source = ss.getSheetByName(sourceCfg.HACHU_SHEET);
  const master = ss.getSheetByName(CFG.MASTER_SHEET);
  const sourceStart = sourceCfg.HACHU_HEADER_ROW + 1;
  const sourceLast = source.getLastRow();
  const sourceWidth = Math.max(sourceCfg.HACHU_CODE, sourceCfg.HACHU_VENDOR, sourceCfg.HACHU_NAME,
    sourceCfg.HACHU_ITEM, sourceCfg.HACHU_PRICE, sourceCfg.HACHU_WEIGHT || 0);
  const sv = source.getRange(sourceStart, 1, sourceLast - sourceStart + 1, sourceWidth).getValues();
  let found = 0;
  for (let i = 0; i < sv.length; i++) {
    const code = String(sv[i][sourceCfg.HACHU_CODE - 1] || '').trim();
    if (code.indexOf('小杉山') < 0) continue;
    found++;
    Logger.log('src row ' + (sourceStart + i) + ' code=[' + code + '] key=[' + _masterPrimaryCodeKey_(code) +
      '] name=[' + sv[i][sourceCfg.HACHU_NAME - 1] + '] excluded=' + _masterIsExcludedCode_(code));
  }
  Logger.log('source hits: ' + found);
  const mStart = CFG.MASTER_HEADER_ROW + 1;
  const mRows = master.getLastRow() - CFG.MASTER_HEADER_ROW;
  const mv = master.getRange(mStart, CFG.M_CODE, mRows, 6).getValues();
  let mfound = 0;
  for (let i = 0; i < mv.length; i++) {
    const code = String(mv[i][0] || '').trim();
    if (code.indexOf('小杉山') < 0) continue;
    mfound++;
    Logger.log('master row ' + (mStart + i) + ' code=[' + code + '] key=[' + _masterPrimaryCodeKey_(code) +
      '] name=[' + mv[i][2] + ']');
  }
  Logger.log('master hits: ' + mfound + ' / masterヘッダー行=' + CFG.MASTER_HEADER_ROW + ' / master最終行=' + master.getLastRow());
}
