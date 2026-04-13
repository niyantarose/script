function onEdit(e) {
  // 手動実行で落ちないように
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  if (sh.getName() !== "シート1") return;    // ←タブ名
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row <= 1) return;                      // ヘッダー除外

  // B列（商品ID）を編集したときだけ反応
  if (col === 2) {
    const hasValue = String(e.range.getValue()).trim() !== "";

    // H列：在庫ステータス
    const hCell = sh.getRange(row, 8);
    if (hasValue) {
      const ruleH = SpreadsheetApp.newDataValidation()
        .requireValueInList(["即納在庫あり","お取り寄せ在庫あり","在庫なし"], true)
        .setAllowInvalid(false)
        .build();
      hCell.setDataValidation(ruleH);
    } else {
      hCell.clearDataValidations();
      hCell.clearContent();
    }

    // I列：配送ステータス
    const iCell = sh.getRange(row, 9);
    if (hasValue) {
      const ruleI = SpreadsheetApp.newDataValidation()
        .requireValueInList(["予約待ち","発注中","EMS入庫待ち","入庫済み"], true)
        .setAllowInvalid(false)
        .build();
      iCell.setDataValidation(ruleI);
    } else {
      iCell.clearDataValidations();
      iCell.clearContent();
    }
  }
}
