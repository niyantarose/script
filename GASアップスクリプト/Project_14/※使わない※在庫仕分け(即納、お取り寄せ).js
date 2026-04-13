function 在庫を更新() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const yahooSheet = ss.getSheetByName('Yahoo全在庫');
  const zaikoSheet = ss.getSheetByName('在庫');

  if (!yahooSheet || !zaikoSheet) {
    throw new Error('「Yahoo全在庫」または「在庫」シートが見つかりません。');
  }

  // ★★ ここで例外を定義 ★★
  // キー: sub-code、値: '即納' または '取寄'
  const exceptionMap = {
    'BEAUTY2005aA': '即納',
    'BEAUTY2005bB': '即納',
    'BEAUTY2005Ab': '取寄',
    'BEAUTY2005Ba': '取寄',
  };

  const yahooLastRow = yahooSheet.getLastRow();
  if (yahooLastRow < 2) return;

  const yahooValues = yahooSheet.getRange(2, 1, yahooLastRow - 1, 4).getValues();
  const stockMap = {}; // { code: { 即納: number, 取寄: number } }

  yahooValues.forEach(row => {
    const code    = row[0];       // A列 code
    const subCode = row[2] || ''; // C列 sub-code
    const qty     = Number(row[3]) || 0;

    if (!code) return;

    if (!stockMap[code]) stockMap[code] = { 即納: 0, 取寄: 0 };

    let kind; // '即納' or '取寄'

    // ① まず例外テーブルを優先
    if (subCode && exceptionMap[subCode]) {
      kind = exceptionMap[subCode]; // 即納or取寄をそのまま採用
    } else {
      // ② 例外でなければ、通常ルール
      //    sub-code が code + "b" と完全一致 → お取り寄せ
      //    それ以外 → 即納
      if (subCode === code + 'b') {
        kind = '取寄';
      } else {
        kind = '即納';
      }
    }

    if (kind === '取寄') {
      stockMap[code].取寄 += qty;
    } else {
      stockMap[code].即納 += qty;
    }
  });

  // 在庫シート側に書き込み
  const zaikoLastRow = zaikoSheet.getLastRow();
  if (zaikoLastRow < 2) return;

  const idValues = zaikoSheet.getRange(2, 1, zaikoLastRow - 1, 1).getValues();
  const output = [];

  idValues.forEach(row => {
    const id = row[0];
    if (!id) {
      output.push([0, 0]);
      return;
    }
    const rec = stockMap[id] || { 即納: 0, 取寄: 0 };
    output.push([rec.即納, rec.取寄]);
  });

  zaikoSheet.getRange(2, 2, output.length, 2).setValues(output);
}
