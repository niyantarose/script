function シート縮小_実データ基準() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetNames = [
    '差異一覧_Amazonのみ',
    '差異一覧_Qoo10のみ',
    '未登録一覧_Amazon（要登録）',
    '未登録一覧_Qoo10（要登録）',
  ]; // ←縮めたいシート名だけ入れる

  for (const name of targetNames) {
    const sh = ss.getSheetByName(name);
    if (!sh) continue;

    const lastRow = Math.max(2, sh.getLastRow());
    const lastCol = Math.max(4, sh.getLastColumn());

    const needRows = lastRow + 20; // 余白
    const needCols = lastCol + 2;  // 余白

    const maxRows = sh.getMaxRows();
    const maxCols = sh.getMaxColumns();

    if (maxRows > needRows) sh.deleteRows(needRows + 1, maxRows - needRows);
    if (maxCols > needCols) sh.deleteColumns(needCols + 1, maxCols - needCols);

    Logger.log(`縮小: ${name} max=${maxRows}x${maxCols} → ${sh.getMaxRows()}x${sh.getMaxColumns()}`);
  }

  ss.toast('指定シートを縮小しました', '完了', 6);
}
