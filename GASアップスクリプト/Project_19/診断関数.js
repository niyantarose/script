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
