const 商品マスタ一意化_CFG = {
  SHEET_NAME: '商品マスタ',
  HEADER_ROW: 6,
  FIRST_COL: 2, // B 商品コード
  WIDTH: 6,     // B:G 商品コード/業者/商品名/品目/価格/重さ
  FAST_THRESHOLD: 250
};

function 商品マスタ_商品コードを一意化() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(商品マスタ一意化_CFG.SHEET_NAME);

  if (!sheet) {
    ui.alert('商品マスタシートが見つかりません。');
    return;
  }

  const plan = 商品マスタ_商品コード一意化_計画_(sheet);
  if (plan.groups.length === 0 && plan.invalidRows.length === 0) {
    ui.alert(
      '商品マスタ 商品コード一意化',
      '商品コードが重複している行、または商品コードが「-」だけの行はありません。',
      ui.ButtonSet.OK
    );
    return;
  }

  const preview = plan.groups.slice(0, 12).map(group => {
    return '・' + group.codeText +
      '\n  残す行: ' + group.keepRow +
      ' / 削除行: ' + group.deleteRows.join(', ') +
      '\n  採用: ' + 商品マスタ_一意化行表示_(group.merged);
  }).join('\n');
  const omitted = plan.groups.length > 12
    ? '\nほか ' + (plan.groups.length - 12) + ' グループ'
    : '';

  const response = ui.alert(
    '商品マスタを商品コードごとに1行へ統合しますか？',
    '同じ商品コードは一番下の行を残し、上の行から足りない値だけ補完してから削除します。\n' +
      '商品コードが空、または「-」だけの行は商品コードなしとして整理対象にします。\n' +
      '通常の商品コード内の「-」は消さずに残します。\n' +
      '価格・重さ・品目などは、下にある新しい値を優先します。\n\n' +
      '重複コード: ' + plan.groups.length + '件\n' +
      '商品コードなし行: ' + plan.invalidRows.length + '行\n' +
      '削除対象: ' + plan.deleteRows.length + '行\n\n' +
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
    const result = 商品マスタ_商品コード一意化_実行_(sheet);
    const modeText = result.mode === 'rewrite'
      ? '\n大量データだったので、タイムアウト防止の高速整理で処理しました。'
      : '';
    ui.alert(
      '商品マスタ一意化完了',
      '統合した商品コード: ' + result.groups + '件\n' +
        '整理した行: ' + result.deleted + '行' + modeText,
      ui.ButtonSet.OK
    );
  } finally {
    lock.releaseLock();
  }
}

function 商品マスタ_重複行を確認して削除() {
  return 商品マスタ_商品コードを一意化();
}

function 商品マスタ_商品コード一意化_実行_(sheet) {
  const plan = 商品マスタ_商品コード一意化_計画_(sheet);
  if (plan.groups.length === 0 && plan.invalidRows.length === 0) {
    return { groups: 0, deleted: 0, invalid: 0, mode: 'delete' };
  }

  if (plan.deleteRows.length > 商品マスタ一意化_CFG.FAST_THRESHOLD) {
    return 商品マスタ_商品コード一意化_高速整理_(sheet);
  }

  const cfg = 商品マスタ_一意化設定_();
  plan.groups.forEach(group => {
    sheet.getRange(group.keepRow, cfg.firstCol, 1, cfg.width).setValues([group.merged]);
  });

  商品マスタ_削除行ブロック_(plan.deleteRows).forEach(block => {
    sheet.deleteRows(block.start, block.count);
  });

  return { groups: plan.groups.length, deleted: plan.deleteRows.length, invalid: plan.invalidRows.length, mode: 'delete' };
}

function 商品マスタ_商品コード一意化_計画_(sheet) {
  const cfg = 商品マスタ_一意化設定_();
  const lastRow = sheet.getLastRow();
  const startRow = cfg.headerRow + 1;
  if (lastRow < startRow) return { groups: [], deleteRows: [] };

  const numRows = lastRow - cfg.headerRow;
  const fullWidth = 商品マスタ_一意化対象幅_(sheet, cfg);
  const values = sheet.getRange(startRow, cfg.firstCol, numRows, fullWidth).getValues();
  const groupsByCode = {};
  const invalidRows = [];

  values.forEach((row, index) => {
    const rowNumber = startRow + index;
    if (商品マスタ_削除対象商品コード_(row[0]) || 商品マスタ_商品コード空データ行_(row)) {
      invalidRows.push(rowNumber);
      return;
    }

    const key = 商品マスタ_商品コードキー_(row[0]);
    if (!key) return;
    if (!groupsByCode[key]) {
      groupsByCode[key] = {
        codeText: String(row[0] || '').trim(),
        rows: []
      };
    }
    groupsByCode[key].rows.push({
      row: rowNumber,
      values: row.slice(0, cfg.width)
    });
  });

  const groups = [];
  const deleteRows = [];

  Object.keys(groupsByCode).forEach(key => {
    const group = groupsByCode[key];
    if (group.rows.length < 2) return;

    const merged = 商品マスタ_一意化マージ_(group.rows);
    const keep = group.rows[group.rows.length - 1];
    const dels = group.rows
      .filter(row => row.row !== keep.row)
      .map(row => row.row);

    groups.push({
      codeKey: key,
      codeText: group.codeText,
      keepRow: keep.row,
      deleteRows: dels,
      merged: merged
    });
    Array.prototype.push.apply(deleteRows, dels);
  });
  Array.prototype.push.apply(deleteRows, invalidRows);

  return {
    groups: groups,
    invalidRows: invalidRows,
    deleteRows: deleteRows.sort((a, b) => b - a)
  };
}

function 商品マスタ_一意化設定_() {
  const cfg = (typeof CFG !== 'undefined') ? CFG : {};
  return {
    headerRow: Number(cfg.MASTER_HEADER_ROW || 商品マスタ一意化_CFG.HEADER_ROW),
    firstCol: Number(cfg.M_CODE || 商品マスタ一意化_CFG.FIRST_COL),
    width: 商品マスタ一意化_CFG.WIDTH
  };
}

function 商品マスタ_一意化対象幅_(sheet, cfg) {
  const lastCol = Math.max(sheet.getLastColumn(), cfg.firstCol + cfg.width - 1);
  return Math.max(cfg.width, lastCol - cfg.firstCol + 1);
}

function 商品マスタ_商品コードキー_(value) {
  if (typeof _masterPrimaryCodeKey_ === 'function') return _masterPrimaryCodeKey_(value);
  let text = String(value == null ? '' : value).trim();
  if (text.normalize) text = text.normalize('NFKC');
  return text
    .replace(/[‐‑‒–—―−]/g, '-')
    .replace(/[＿]/g, '_')
    .replace(/[\s\u00A0\u3000]+/g, '')
    .toUpperCase()
    .replace(/_/g, '-')
    .replace(/-+/g, '-');
}

function 商品マスタ_削除対象商品コード_(value) {
  const raw = String(value == null ? '' : value).trim();
  if (!raw) return false;
  let text = raw;
  if (text.normalize) text = text.normalize('NFKC');
  return text
    .replace(/[‐‑‒–—―−]/g, '-')
    .replace(/[＿]/g, '_')
    .replace(/[\s\u00A0\u3000]+/g, '')
    .replace(/_/g, '-')
    .replace(/-+/g, '-') === '-';
}

function 商品マスタ_商品コード空データ行_(row) {
  if (String(row[0] == null ? '' : row[0]).trim() !== '') return false;
  return row.slice(1).some(商品マスタ_値あり_);
}

function 商品マスタ_値あり_(value) {
  if (typeof _masterHasValue_ === 'function') return _masterHasValue_(value);
  return value !== '' && value !== null && typeof value !== 'undefined' && String(value).trim() !== '';
}

function 商品マスタ_一意化マージ_(rows) {
  const width = 商品マスタ一意化_CFG.WIDTH;
  const merged = Array(width).fill('');

  rows.forEach(item => {
    for (let c = 0; c < width; c++) {
      if (商品マスタ_値あり_(item.values[c])) {
        merged[c] = item.values[c];
      }
    }
  });

  return merged;
}

function 商品マスタ_商品コード一意化_高速整理_(sheet) {
  const cfg = 商品マスタ_一意化設定_();
  const lastRow = sheet.getLastRow();
  const startRow = cfg.headerRow + 1;
  if (lastRow < startRow) return { groups: 0, deleted: 0, invalid: 0, mode: 'rewrite' };

  const numRows = lastRow - cfg.headerRow;
  const fullWidth = 商品マスタ_一意化対象幅_(sheet, cfg);
  const values = sheet.getRange(startRow, cfg.firstCol, numRows, fullWidth).getValues();
  const groupsByCode = {};
  let invalid = 0;
  let dataRows = 0;

  values.forEach((row, index) => {
    const hasAnyValue = row.some(商品マスタ_値あり_);
    if (hasAnyValue) dataRows++;

    if (商品マスタ_削除対象商品コード_(row[0])) {
      invalid++;
      return;
    }
    if (商品マスタ_商品コード空データ行_(row)) {
      invalid++;
      return;
    }

    const key = 商品マスタ_商品コードキー_(row[0]);
    if (!key) return;

    if (!groupsByCode[key]) {
      groupsByCode[key] = {
        codeKey: key,
        codeText: String(row[0] || '').trim(),
        latestIndex: index,
        sourceCount: 0,
        merged: Array(fullWidth).fill('')
      };
    }

    const group = groupsByCode[key];
    group.latestIndex = index;
    group.sourceCount++;
    for (let c = 0; c < fullWidth; c++) {
      if (商品マスタ_値あり_(row[c])) group.merged[c] = row[c];
    }
  });

  const groups = Object.keys(groupsByCode)
    .map(key => groupsByCode[key])
    .sort((a, b) => a.latestIndex - b.latestIndex);
  const output = groups.map(group => group.merged);

  if (output.length > 0) {
    sheet.getRange(startRow, cfg.firstCol, output.length, fullWidth).setValues(output);
  }

  if (numRows > output.length) {
    sheet.getRange(startRow + output.length, cfg.firstCol, numRows - output.length, fullWidth).clearContent();
  }

  const duplicateGroups = groups.filter(group => group.sourceCount > 1).length;
  const deleted = Math.max(0, dataRows - output.length);
  return { groups: duplicateGroups, deleted: deleted, invalid: invalid, mode: 'rewrite' };
}

function 商品マスタ_一意化行表示_(row) {
  return [
    '商品コード:' + (row[0] || '(空)'),
    '業者:' + (row[1] || '(空)'),
    '商品名:' + (row[2] || '(空)'),
    '品目:' + (row[3] || '(空)'),
    '価格:' + (row[4] || '(空)'),
    '重さ:' + (row[5] || '(空)')
  ].join(' / ');
}

function 商品マスタ_削除行ブロック_(rows) {
  const sorted = rows.slice().sort((a, b) => b - a);
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
