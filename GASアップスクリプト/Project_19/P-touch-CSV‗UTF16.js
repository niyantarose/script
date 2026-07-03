function Pタッチ印刷用CSV_UTF16を保存() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('EMS大邱作業データ');
  if (!sheet) throw new Error('シート「EMS大邱作業データ」がありません。');

  const selection = ss.getActiveRange();
  const startRow = selection.getRow();
  const endRow = selection.getLastRow();

  const values = sheet.getDataRange().getValues();
  if (values.length < 3) { SpreadsheetApp.getUi().alert('データがありません。'); return; }

  const header = values[1].map(v => String(v).trim());
  const codeCol = findHeaderColumn_(header, ['ItemCode', 'Code', '商品コード']);
  const nameCol = findHeaderColumn_(header, ['Description / Title', 'Product Name', '商品名']);
  const qtyCol  = findHeaderColumn_(header, ['Qty', 'Units']);

  if (codeCol === -1) throw new Error('「ItemCode」列がありません。');
  if (nameCol === -1) throw new Error('「Description / Title」列がありません。');

  const merged = {};
  let promoCount = 0;

  for (let i = startRow - 1; i < endRow; i++) {
    const code = String(values[i][codeCol] || '').trim();
    if (!code) continue;
    const name = String(values[i][nameCol] || '').trim();
    const qty  = parseInt(values[i][qtyCol], 10) || 1;

    if (code === 'Promotional Item') {
      const key = 'PROMO_' + promoCount++;
      merged[key] = { code: code, name: name, qty: qty };
    } else {
      if (merged[code]) {
        merged[code].qty += qty;
      } else {
        merged[code] = { code: code, name: name, qty: qty };
      }
    }
  }

  if (Object.keys(merged).length === 0) {
    SpreadsheetApp.getUi().alert('対象データがありません。');
    return;
  }

  const rows = [['商品コード', '商品名']];
  for (const key in merged) {
    const { code, name, qty } = merged[key];
    for (let i = 0; i < qty; i++) {
      rows.push([code, name]);
    }
  }

  const csv = rows.map(r => r.map(c => {
    const s = String(c);
    return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }).join(',')).join('\r\n');

  const FIXED_NAME = 'P-touch_today.csv';
  const blob = Utilities.newBlob('', 'text/csv')
    .setDataFromString(csv, 'UTF-16LE')
    .setName(FIXED_NAME);

  const folder = getOutputFolder_();
  const existing = folder.getFilesByName(FIXED_NAME);
  while (existing.hasNext()) existing.next().setTrashed(true);
  folder.createFile(blob);

  SpreadsheetApp.getUi().alert(
    'P-touch用CSVを更新しました！（UTF-16版）\n\n' +
    'ファイル名：' + FIXED_NAME + '\n' +
    'ラベル枚数：' + (rows.length - 1) + '枚\n\n' +
    'P-touch Editorで「更新」を押してください。'
  );
}