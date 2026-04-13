function Qoo10在庫数量更新_全行() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const yahooSheet = ss.getSheetByName('Yahoo全在庫');
  const qooSheet   = ss.getSheetByName('Qoo10');

  if (!yahooSheet || !qooSheet) {
    throw new Error('Yahoo全在庫 または Qoo10 シートが見つかりません。');
  }

  Logger.log('=== Qoo10在庫数量更新_全行 開始 ===');

  // 1. Yahoo全在庫 → baseMap & subStockMap 作成
  const yahooLastRow = yahooSheet.getLastRow();
  if (yahooLastRow < 2) {
    Browser.msgBox('Yahoo全在庫 にデータがありません。');
    return;
  }

  const yahooValues = yahooSheet.getRange(2, 1, yahooLastRow - 1, 4).getValues();

  const baseMap = {};      // base -> { stockA, stockB, qtyA, qtyB }
  const subStockMap = {};  // subCode -> stock

  function setStockToBase(base, kind, stock) {
    if (!baseMap[base]) {
      baseMap[base] = { stockA: 0, stockB: 0, qtyA: 0, qtyB: 0 };
    }
    if (kind === 'a') {
      baseMap[base].stockA = stock;
    } else if (kind === 'b') {
      baseMap[base].stockB = stock;
    }
  }

  yahooValues.forEach(row => {
    const subCode = String(row[2] || '').trim();  // C列: 子コード
    const stock   = Number(row[3]) || 0;          // D列: 在庫数
    if (!subCode) return;

    // 子コードそのものの在庫
    subStockMap[subCode] = stock;

    // ★ 末尾「小文字 a / b」だけ baseMap に登録（大文字は無視）
    const m = subCode.match(/^(.+)(a|b)$/);   // ← i を付けない
    if (m) {
      const base = m[1];
      const kind = m[2];                      // 'a' or 'b'
      setStockToBase(base, kind, stock);
    }
  });

  // baseごとに数量決定（a/bルール）
  Object.keys(baseMap).forEach(base => {
    const info = baseMap[base];
    const hasA = info.stockA > 0;
    const hasB = info.stockB > 0;

    if (!hasA && !hasB) {
      info.qtyA = 0;
      info.qtyB = 0;
    } else if (hasA) {
      info.qtyA = info.stockA; // 即納は実在庫数
      info.qtyB = 0;
    } else {
      info.qtyA = 0;
      info.qtyB = 10;          // 即納なし・取寄だけ → 10
    }
  });

  Logger.log('base件数: ' + Object.keys(baseMap).length);

  // 2. Qoo10 N列更新
  const qooLastRow = qooSheet.getLastRow();
  if (qooLastRow < 2) {
    Browser.msgBox('Qoo10 シートにデータ行がありません。');
    return;
  }

  const nValues = qooSheet.getRange(2, 14, qooLastRow - 1, 1).getValues();
  const newNValues = [];
  let updatedRowCount = 0;

  for (let i = 0; i < nValues.length; i++) {
    const original = String(nValues[i][0] || '');
    if (!original) {
      newNValues.push([original]);
      continue;
    }

    let blocks  = original.split('$$');
    let changed = false;

    for (let j = 0; j < blocks.length; j++) {
      let block = (blocks[j] || '').trim();
      if (!block) continue;

      let parts = block.split('||');
      if (parts.length < 3) continue;

      const codePart = parts[parts.length - 1].replace(/^\*/, '').trim();
      if (!codePart) continue;

      const optionType = parts[1] || '';
      const is即納 = optionType.indexOf('即納')      !== -1;
      const is取寄 = optionType.indexOf('お取り寄せ') !== -1;

      let newQty = 0;

      // ★ 末尾小文字 a/b の通常コードなら baseMap を優先
      let m = codePart.match(/^(.+)(a|b)$/);  // ← ここも i なし
      if (m && baseMap[m[1]]) {
        const base = m[1];
        const kind = m[2];                   // 'a' or 'b'
        const info = baseMap[base];
        newQty = (kind === 'a') ? info.qtyA : info.qtyB;
      } else {
        // ★ それ以外（HCUT259aA / aB / bA / bB 等）は子コード在庫をそのまま使う
        const stockDirect = Number(subStockMap[codePart] || 0);

        if (stockDirect <= 0) {
          newQty = 0;
        } else if (is即納) {
          newQty = stockDirect;  // 即納 → 実在庫数
        } else if (is取寄) {
          newQty = 10;           // 取寄 → 在庫があれば10
        } else {
          newQty = 0;
        }
      }

      const qtyIndex = parts.length - 2;
      const before   = parts[qtyIndex];
      const newPart  = '*' + newQty;

      if (before !== newPart) {
        parts[qtyIndex] = newPart;
        blocks[j]       = parts.join('||');
        changed         = true;
      }
    }

    if (changed) {
      const updated = blocks.join('$$');
      newNValues.push([updated]);
      updatedRowCount++;
    } else {
      newNValues.push([original]);
    }
  }

  qooSheet.getRange(2, 14, newNValues.length, 1).setValues(newNValues);

  Logger.log('更新行数: ' + updatedRowCount + ' 行');
  Browser.msgBox('Qoo10在庫数量更新完了: ' + updatedRowCount + ' 行を更新しました。');

  Logger.log('=== Qoo10在庫数量更新_全行 終了 ===');
}
