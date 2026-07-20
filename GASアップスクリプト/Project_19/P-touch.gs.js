const PTUCH_OUTPUT_FOLDER_ID = '1RpfZYLGs6SMz9Bfd9rnCJW4OIEQ3K2HN';

function Pタッチ印刷用Excelを保存() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('EMS大邱作業データ');
  if (!sheet) throw new Error('シート「EMS大邱作業データ」がありません。');

  // 選択範囲の行番号を取得
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

  // ステップ1: 選択行を抽出しつつ重複をまとめる
  // Promotional Item と 注文番号（数字のみコード）は合体させず商品名ごとに別管理
  const merged = {};
  let promoCount = 0;
  const printedRows = []; // シール発行済みにする行番号(U列)

  for (let i = startRow - 1; i < endRow; i++) {
    const code = String(values[i][codeCol] || '').trim();
    if (!code) continue;
    printedRows.push(i + 1);
    let name = String(values[i][nameCol] || '').trim();
    const qty  = parseInt(values[i][qtyCol], 10) || 1;

    // 商品名の個別置換（特定の韓国語文言を英語表記に変換）
    name = applyNameReplacements_(name);

    // スペース有無・大文字小文字を無視してPromotional Item判定
    const codeNormalized = code.replace(/\s/g, '').toLowerCase();
    // コードが数字だけ（注文番号）かどうか判定
    const isOrderNumber = /^\d+$/.test(code);

    if (codeNormalized === 'promotionalitem' || isOrderNumber) {
      // Promotional Item と 注文番号コードは商品名ごとに別キーで管理（合体させない）
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

  // ステップ2: 数量分だけ行を展開
  const rows = [['商品コード', '商品名']];
  for (const key in merged) {
    const { code, name, qty } = merged[key];

    // Promotional Item系だけ表示用に置き換え。注文番号（数字コード）はそのまま表示
    const codeNormalized = code.replace(/\s/g, '').toLowerCase();
    const displayCode = (codeNormalized === 'promotionalitem')
      ? 'ささやかなおまけ(謝恩品)'
      : code;

    for (let i = 0; i < qty; i++) {
      rows.push([displayCode, name]);
    }
  }

  // xlsxに変換して保存
  const FIXED_NAME = 'P-touch_today';
  const tempSs = SpreadsheetApp.create(FIXED_NAME);
  const tempSheet = tempSs.getSheets()[0];
  tempSheet.setName('印刷用');
  tempSheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  SpreadsheetApp.flush();

  const blob = UrlFetchApp.fetch(
    'https://docs.google.com/spreadsheets/d/' + tempSs.getId() + '/export?format=xlsx',
    { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }
  ).getBlob().setName(FIXED_NAME + '.xlsx');

  const folder = getOutputFolder_();
  const existing = folder.getFilesByName(FIXED_NAME + '.xlsx');
  while (existing.hasNext()) existing.next().setTrashed(true);
  folder.createFile(blob);

  DriveApp.getFileById(tempSs.getId()).setTrashed(true);

  // 印刷した行をシール発行済みに(U列に発行日・未発行の薄い色を解除)
  if (typeof EMS大邱_シール発行済みにする_ === 'function') {
    EMS大邱_シール発行済みにする_(sheet, printedRows);
  }

  SpreadsheetApp.getUi().alert(
    'P-touch用Excelを更新しました！\n\n' +
    'ファイル名：' + FIXED_NAME + '.xlsx\n' +
    'ラベル枚数：' + (rows.length - 1) + '枚\n' +
    'シール発行済み：' + printedRows.length + '行(U列に記入)\n\n' +
    'P-touch Editorで「更新」を押してください。'
  );
}

// 商品名の個別置換ルール（韓国語の特定文言→英語表記など）
// 増やしたい場合はこの配列に { from, to } を追加していくだけでOK
function applyNameReplacements_(name) {
  const replacements = [
    { from: '케이크받침', to: 'Paperboard for display' }
  ];
  let result = name;
  for (const r of replacements) {
    result = result.split(r.from).join(r.to);
  }
  return result;
}

function findHeaderColumn_(header, candidates) {
  for (let i = 0; i < header.length; i++) {
    if (candidates.includes(header[i])) return i;
  }
  return -1;
}

function getOutputFolder_() {
  return DriveApp.getFolderById(PTUCH_OUTPUT_FOLDER_ID);
}