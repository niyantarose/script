const 商品マスタ重複削除_CFG = {
  SHEET_NAME: '商品マスタ',
  HEADER_ROW: 6,
  CODE_COL: 2,
  KEY_COLUMNS: [
    { label: '商品コード', col: 2 },
    { label: '業者', col: 3 },
    { label: '商品名', col: 4 },
    { label: '品目', col: 5 }
  ]
};

function 商品マスタ_重複行を確認して削除() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(商品マスタ重複削除_CFG.SHEET_NAME);

  if (!sheet) {
    ui.alert('商品マスタシートが見つかりません。');
    return;
  }

  const result = 商品マスタ_重複行を抽出_(sheet);
  if (result.deleteRows.length === 0) {
    ui.alert(
      '商品マスタ重複チェック',
      '削除対象の重複行はありません。\n判定列: 商品コード + 業者 + 商品名 + 品目',
      ui.ButtonSet.OK
    );
    return;
  }

  const preview = result.groups.slice(0, 12).map(function(group) {
    return '・' + group.keyText + '\n  残す行: ' + group.keepRow + ' / 削除行: ' + group.deleteRows.join(', ');
  }).join('\n');
  const omitted = result.groups.length > 12
    ? '\nほか ' + (result.groups.length - 12) + ' グループ'
    : '';

  const response = ui.alert(
    '商品マスタの重複行を削除しますか？',
    '判定列: 商品コード + 業者 + 商品名 + 品目\n' +
      '上にある1行を残して、下側の重複行を削除します。\n\n' +
      '重複グループ: ' + result.groups.length + '件\n' +
      '削除対象: ' + result.deleteRows.length + '行\n\n' +
      preview + omitted,
    ui.ButtonSet.OK_CANCEL
  );
  if (response !== ui.Button.OK) return;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) {
    ui.alert('今ほかの処理が動いているみたいです。少し待ってからもう一度実行してください。');
    return;
  }

  try {
    const latest = 商品マスタ_重複行を抽出_(sheet);
    if (latest.deleteRows.length === 0) {
      ui.alert('削除直前に再確認したところ、削除対象の重複行はありませんでした。');
      return;
    }

    商品マスタ_削除行ブロック_(latest.deleteRows).forEach(function(block) {
      sheet.deleteRows(block.start, block.count);
    });

    ui.alert(
      '商品マスタ重複削除完了',
      '削除した行: ' + latest.deleteRows.length + '行\n' +
        '残した重複グループ: ' + latest.groups.length + '件',
      ui.ButtonSet.OK
    );
  } finally {
    lock.releaseLock();
  }
}

function 商品マスタ_重複行を抽出_(sheet) {
  const cfg = 商品マスタ重複削除_CFG;
  const lastRow = sheet.getLastRow();
  const startRow = cfg.HEADER_ROW + 1;
  if (lastRow < startRow) {
    return { groups: [], deleteRows: [] };
  }

  const values = sheet.getRange(startRow, 1, lastRow - cfg.HEADER_ROW, sheet.getLastColumn()).getValues();
  const seen = new Map();

  values.forEach(function(row, index) {
    const rowNumber = startRow + index;
    const code = 商品マスタ_重複判定値_(row[cfg.CODE_COL - 1]);
    if (!code) return;

    const keyParts = cfg.KEY_COLUMNS.map(function(col) {
      return 商品マスタ_重複判定値_(row[col.col - 1]);
    });
    const key = keyParts.join('\u241F');

    if (!seen.has(key)) {
      seen.set(key, {
        keepRow: rowNumber,
        deleteRows: [],
        keyText: 商品マスタ_重複キー表示_(row)
      });
      return;
    }

    seen.get(key).deleteRows.push(rowNumber);
  });

  const groups = Array.from(seen.values()).filter(function(group) {
    return group.deleteRows.length > 0;
  });
  const deleteRows = [];
  groups.forEach(function(group) {
    Array.prototype.push.apply(deleteRows, group.deleteRows);
  });

  return { groups: groups, deleteRows: deleteRows };
}

function 商品マスタ_重複判定値_(value) {
  let text = String(value == null ? '' : value).trim();
  if (text.normalize) text = text.normalize('NFKC');
  return text
    .replace(/[\s\u00A0\u3000]+/g, ' ')
    .toUpperCase();
}

function 商品マスタ_重複キー表示_(row) {
  return 商品マスタ重複削除_CFG.KEY_COLUMNS.map(function(col) {
    const value = String(row[col.col - 1] == null ? '' : row[col.col - 1]).trim();
    return col.label + ':' + (value || '(空)');
  }).join(' / ');
}

function 商品マスタ_削除行ブロック_(rows) {
  const sorted = rows.slice().sort(function(a, b) { return b - a; });
  if (sorted.length === 0) return [];

  const blocks = [];
  let high = sorted[0];
  let low = sorted[0];

  for (let i = 1; i < sorted.length; i++) {
    const row = sorted[i];
    if (row === low - 1) {
      low = row;
      continue;
    }

    blocks.push({ start: low, count: high - low + 1 });
    high = row;
    low = row;
  }

  blocks.push({ start: low, count: high - low + 1 });
  return blocks;
}
