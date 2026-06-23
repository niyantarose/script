const LAST_DATA_ROW_NAV_CFG = {
  DEFAULT_HEADER_ROW: 0,
  DEFAULT_FOCUS_COL: 1,
  SHEETS: {
    '発注リスト大邱データ': { headerRow: 5, keyCols: [6, 9, 11, 12], focusCol: 6 },
    'EMS大邱作業データ': { headerRow: 2, keyCols: [4, 8, 9, 20], focusCol: 4 },
    '発注': { headerRow: 6, keyCols: [7, 10, 12, 13], focusCol: 7 },
    'EMSリスト': { headerRow: 6, keyCols: [1, 2, 3, 6, 9, 10, 13], focusCol: 1 },
    '商品マスタ': { headerRow: 6, keyCols: [2, 4, 5], focusCol: 2 },
    '発注リスト太郎データ': { headerRow: 5, keyCols: [6, 9, 11, 12], focusCol: 6 },
    'EMS太郎作業データ': { headerRow: 2, keyCols: [4, 8, 9, 20], focusCol: 4 }
  }
};

function 最終データ行へ移動() {
  return 最終データ行へ移動_(false);
}

function 最終データ行へ移動_自動() {
  return 最終データ行へ移動_(true);
}

function 最終データ行へ移動_(silent) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  if (!sh) return 0;

  SpreadsheetApp.flush();

  const row = 最終データ行を取得_(sh);
  const col = 最終データ行_移動先列_(sh);

  ss.setActiveSheet(sh);
  sh.getRange(row, col).activate();

  if (!silent) {
    ss.toast(`${sh.getName()} の ${row} 行目へ移動しました`, '最終データ行へ移動', 4);
  }
  return row;
}

function 最終データ行を取得_(sh) {
  const cfg = LAST_DATA_ROW_NAV_CFG.SHEETS[sh.getName()] || {};
  const headerRow = Number(cfg.headerRow || LAST_DATA_ROW_NAV_CFG.DEFAULT_HEADER_ROW || 0);
  const startRow = Math.max(1, headerRow + 1);
  const lastRow = Math.max(sh.getLastRow(), startRow);
  if (lastRow < startRow) return startRow;

  const keyCols = Array.isArray(cfg.keyCols) && cfg.keyCols.length
    ? cfg.keyCols
    : 最終データ行_スキャン列_(sh);

  const numRows = lastRow - startRow + 1;
  const columns = keyCols
    .map(Number)
    .filter(col => col >= 1 && col <= sh.getMaxColumns());

  if (columns.length === 0) return lastRow;

  const colValues = columns.map(col => {
    return sh.getRange(startRow, col, numRows, 1).getDisplayValues();
  });

  for (let i = numRows - 1; i >= 0; i--) {
    for (let c = 0; c < colValues.length; c++) {
      if (String(colValues[c][i][0] || '').trim() !== '') {
        return startRow + i;
      }
    }
  }

  return startRow;
}

function 最終データ行_移動先列_(sh) {
  const cfg = LAST_DATA_ROW_NAV_CFG.SHEETS[sh.getName()] || {};
  const active = sh.getActiveRange();
  const activeCol = active ? active.getColumn() : 0;
  const col = Number(cfg.focusCol || activeCol || LAST_DATA_ROW_NAV_CFG.DEFAULT_FOCUS_COL || 1);
  return Math.max(1, Math.min(col, sh.getMaxColumns()));
}

function 最終データ行_スキャン列_(sh) {
  const lastCol = Math.max(1, sh.getLastColumn());
  const maxScanCol = Math.min(lastCol, 30);
  const cols = [];
  for (let col = 1; col <= maxScanCol; col++) cols.push(col);
  return cols;
}
