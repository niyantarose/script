const YAHOO_PRODUCT_NAME_CFG = {
  MASTER_SHEET: '商品マスタ',
  HEADER_ROW: 6,
  CODE_COL: 2,
  OUTPUT_START_COL: 8,
  OUTPUT_HEADERS: ['Yahoo商品名', 'Yahoo参照コード', 'Yahoo取得日時', 'Yahoo取得結果'],
  CSV_FOLDER_ID_PROP: 'YAHOO_QUANTITY_NAME_CSV_FOLDER_ID',
  LEGACY_CSV_FILE_ID_PROP: 'YAHOO_QUANTITY_NAME_CSV_FILE_ID',
  DEFAULT_CSV_FOLDER_ID: '1Gs1VuTw91dF7zHaHu5iJll9_MCxE4p3D'
};

function YahooProductName_addMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('Yahoo商品名')
    .addItem('CSV/Excelフォルダを設定', 'YahooProductName_setCsvFolder')
    .addItem('使用する最新ファイルを確認', 'YahooProductName_showLatestCsv')
    .addSeparator()
    .addItem('確認してOKなら空欄だけ反映', 'YahooProductName_confirmLatestAndFillEmptyFromCsv')
    .addItem('空欄だけ商品マスタへ反映', 'YahooProductName_fillEmptyFromCsv')
    .addItem('上書きで商品マスタへ反映', 'YahooProductName_overwriteFromCsv')
    .addToUi();
}

function YahooProductName_setCsvFolder() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    'Yahoo商品名CSV/Excelフォルダを設定',
    '最新のCSV/Excel/Googleスプレッドシートを入れるDriveフォルダのURL、またはフォルダIDを入力してください。',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const folderId = YahooProductName_extractDriveId_(res.getResponseText());
  if (!folderId) {
    ui.alert('フォルダIDを読み取れませんでした。DriveのフォルダURLかフォルダIDを入力してください。');
    return;
  }

  let latest = null;
  try {
    latest = YahooProductName_findLatestSourceFileInFolder_(folderId);
  } catch (err) {
    ui.alert('フォルダを読み取れませんでした。\n\n' + err.message);
    return;
  }

  const props = PropertiesService.getDocumentProperties();
  props.setProperty(YAHOO_PRODUCT_NAME_CFG.CSV_FOLDER_ID_PROP, folderId);
  props.deleteProperty(YAHOO_PRODUCT_NAME_CFG.LEGACY_CSV_FILE_ID_PROP);

  ui.alert(
    'CSV/Excelフォルダを設定しました。\n\n' +
    'フォルダID: ' + folderId + '\n' +
    '最新ファイル: ' + (latest ? latest.getName() : '見つかりません')
  );
}

function YahooProductName_showLatestCsv() {
  const ui = SpreadsheetApp.getUi();
  const latest = YahooProductName_getLatestSourceFile_();
  if (!latest) return;

  ui.alert(
    '使用する最新ファイル',
    'フォルダ内の最新CSV/Excel/Googleスプレッドシートを使用します。\n\n' +
      'ファイル名: ' + latest.getName() + '\n' +
      '種類: ' + latest.getMimeType() + '\n' +
      '更新日時: ' + Utilities.formatDate(
        latest.getLastUpdated(),
        Session.getScriptTimeZone(),
        'yyyy-MM-dd HH:mm'
      ) + '\n' +
      'ファイルID: ' + latest.getId(),
    ui.ButtonSet.OK
  );
}

function YahooProductName_confirmLatestAndFillEmptyFromCsv() {
  const ui = SpreadsheetApp.getUi();
  const latest = YahooProductName_getLatestSourceFile_();
  if (!latest) return;

  const res = ui.alert(
    '最新ファイルを確認して反映',
    'このファイルで、商品マスタの空欄だけ反映しますか？\n\n' +
      'ファイル名: ' + latest.getName() + '\n' +
      '種類: ' + latest.getMimeType() + '\n' +
      '更新日時: ' + Utilities.formatDate(
        latest.getLastUpdated(),
        Session.getScriptTimeZone(),
        'yyyy-MM-dd HH:mm'
      ) + '\n' +
      'ファイルID: ' + latest.getId(),
    ui.ButtonSet.OK_CANCEL
  );
  if (res !== ui.Button.OK) return;

  YahooProductName_importFromCsv_(false, latest);
}

function YahooProductName_fillEmptyFromCsv() {
  YahooProductName_importFromCsv_(false);
}

function YahooProductName_overwriteFromCsv() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert(
    'Yahoo商品名を上書き',
    '既存のYahoo商品名も上書きします。よろしいですか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (res !== ui.Button.OK) return;
  YahooProductName_importFromCsv_(true);
}

function YahooProductName_importFromCsv_(overwrite, sourceFileOverride) {
  const cfg = YAHOO_PRODUCT_NAME_CFG;
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const master = ss.getSheetByName(cfg.MASTER_SHEET);
  if (!master) {
    ui.alert('商品マスタシートが見つかりません。');
    return;
  }

  const sourceFile = sourceFileOverride || YahooProductName_getLatestSourceFile_();
  if (!sourceFile) return;

  const csvData = YahooProductName_readSourceIndexFromFile_(sourceFile);
  const index = csvData.index;

  YahooProductName_prepareMasterColumns_(master);

  const startRow = cfg.HEADER_ROW + 1;
  const lastRow = master.getLastRow();
  if (lastRow < startRow) {
    ui.alert('商品マスタにデータ行がありません。');
    return;
  }

  const rowCount = lastRow - cfg.HEADER_ROW;
  const codes = master.getRange(startRow, cfg.CODE_COL, rowCount, 1).getDisplayValues();
  const outputRange = master.getRange(
    startRow,
    cfg.OUTPUT_START_COL,
    rowCount,
    cfg.OUTPUT_HEADERS.length
  );
  const outputValues = outputRange.getValues();
  const now = new Date();

  let updated = 0;
  let alreadyFilled = 0;
  let notFound = 0;
  let ambiguous = 0;
  let blankCode = 0;

  for (let i = 0; i < rowCount; i++) {
    const masterCode = String(codes[i][0] || '').trim();
    if (!masterCode) {
      blankCode++;
      continue;
    }

    const currentName = String(outputValues[i][0] || '').trim();
    if (!overwrite && currentName) {
      alreadyFilled++;
      continue;
    }

    const rec = index[YahooProductName_normCode_(masterCode)];
    if (!rec) {
      outputValues[i][3] = '未検出';
      notFound++;
      continue;
    }

    if (rec.names.length > 1) {
      outputValues[i][0] = overwrite ? '' : outputValues[i][0];
      outputValues[i][1] = rec.refs.slice(0, 5).join(', ');
      outputValues[i][2] = now;
      outputValues[i][3] = '候補複数: ' + rec.names.slice(0, 3).join(' / ');
      ambiguous++;
      continue;
    }

    outputValues[i][0] = rec.names[0];
    outputValues[i][1] = rec.refs.slice(0, 5).join(', ');
    outputValues[i][2] = now;
    outputValues[i][3] = 'OK';
    updated++;
  }

  outputRange.setValues(outputValues);
  master.getRange(startRow, cfg.OUTPUT_START_COL + 2, rowCount, 1)
    .setNumberFormat('yyyy-mm-dd hh:mm');

  ui.alert(
    'Yahoo商品名の反映が完了しました。\n\n' +
    '更新: ' + updated + '\n' +
    '既存値ありスキップ: ' + alreadyFilled + '\n' +
    '未検出: ' + notFound + '\n' +
    '候補複数: ' + ambiguous + '\n' +
    '商品コード空欄: ' + blankCode + '\n\n' +
    '読込ファイル: ' + (csvData.name || sourceFile.getName())
  );
}

function YahooProductName_prepareMasterColumns_(sheet) {
  const cfg = YAHOO_PRODUCT_NAME_CFG;
  const range = sheet.getRange(
    cfg.HEADER_ROW,
    cfg.OUTPUT_START_COL,
    1,
    cfg.OUTPUT_HEADERS.length
  );
  range.setValues([cfg.OUTPUT_HEADERS]);
  range.setFontWeight('bold');
}

function YahooProductName_getLatestSourceFile_() {
  const folderId = YahooProductName_getCsvFolderId_();
  if (!folderId) return null;

  const ui = SpreadsheetApp.getUi();
  let latest = null;
  try {
    latest = YahooProductName_findLatestSourceFileInFolder_(folderId);
  } catch (err) {
    ui.alert('CSV/Excelフォルダを読み取れませんでした。\n\n' + err.message);
    return null;
  }

  if (!latest) {
    ui.alert(
      'CSV/Excelファイルが見つかりません。',
      '設定フォルダ内に .csv / .xlsx / .xls / .xlsm / Googleスプレッドシート のいずれかを置いてください。\n\n' +
        'フォルダID: ' + folderId,
      ui.ButtonSet.OK
    );
    return null;
  }

  return latest;
}

function YahooProductName_getCsvFolderId_() {
  const props = PropertiesService.getDocumentProperties();
  let folderId = props.getProperty(YAHOO_PRODUCT_NAME_CFG.CSV_FOLDER_ID_PROP) ||
    YAHOO_PRODUCT_NAME_CFG.DEFAULT_CSV_FOLDER_ID;
  if (folderId) return folderId;

  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    'Yahoo商品名CSV/Excelフォルダ',
    '最新ファイルを入れるDriveフォルダのURL、またはフォルダIDを入力してください。',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return '';

  folderId = YahooProductName_extractDriveId_(res.getResponseText());
  if (!folderId) {
    ui.alert('フォルダIDを読み取れませんでした。');
    return '';
  }

  props.setProperty(YAHOO_PRODUCT_NAME_CFG.CSV_FOLDER_ID_PROP, folderId);
  props.deleteProperty(YAHOO_PRODUCT_NAME_CFG.LEGACY_CSV_FILE_ID_PROP);
  return folderId;
}

function YahooProductName_findLatestSourceFileInFolder_(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  let latest = null;

  while (files.hasNext()) {
    const file = files.next();
    if (!YahooProductName_isSupportedSourceFile_(file)) continue;

    if (!latest || file.getLastUpdated().getTime() > latest.getLastUpdated().getTime()) {
      latest = file;
    }
  }

  return latest;
}

function YahooProductName_isSupportedSourceFile_(file) {
  const name = String(file.getName() || '').toLowerCase();
  const mime = String(file.getMimeType() || '').toLowerCase();

  if (mime === 'application/vnd.google-apps.spreadsheet') return true;
  if (name.match(/\.(csv|xlsx|xls|xlsm)$/)) return true;
  if (mime === 'text/csv') return true;
  if (mime === 'application/vnd.ms-excel') return true;
  if (mime === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') return true;
  if (mime === 'application/vnd.ms-excel.sheet.macroenabled.12') return true;
  return false;
}

function YahooProductName_extractDriveId_(value) {
  const text = String(value || '').trim();
  if (!text) return '';

  let match = text.match(/\/folders\/([A-Za-z0-9_-]+)/);
  if (match) return match[1];

  match = text.match(/\/d\/([A-Za-z0-9_-]+)/);
  if (match) return match[1];

  match = text.match(/[?&]id=([A-Za-z0-9_-]+)/);
  if (match) return match[1];

  match = text.match(/[-\w]{20,}/);
  return match ? match[0] : '';
}

function YahooProductName_extractDriveFileId_(value) {
  return YahooProductName_extractDriveId_(value);
}

function YahooProductName_readSourceIndexFromFile_(file) {
  const mime = String(file.getMimeType() || '').toLowerCase();
  const name = String(file.getName() || '');
  const lowerName = name.toLowerCase();

  if (mime === 'application/vnd.google-apps.spreadsheet') {
    return YahooProductName_readSpreadsheetIndex_(file.getId(), name);
  }

  if (lowerName.match(/\.(xlsx|xls|xlsm)$/) ||
      mime === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
      mime === 'application/vnd.ms-excel' ||
      mime === 'application/vnd.ms-excel.sheet.macroenabled.12') {
    return YahooProductName_readExcelIndex_(file);
  }

  return YahooProductName_readCsvIndexFromFileId_(file.getId());
}

function YahooProductName_readSpreadsheetIndex_(spreadsheetId, name) {
  const spreadsheet = YahooProductName_openSpreadsheetWithRetry_(spreadsheetId);
  const sheet = spreadsheet.getSheets()[0];
  const rows = sheet.getDataRange().getDisplayValues();
  return {
    name: name,
    index: YahooProductName_buildCsvIndex_(rows)
  };
}

function YahooProductName_openSpreadsheetWithRetry_(spreadsheetId) {
  let lastError = null;
  for (let i = 0; i < 6; i++) {
    try {
      return SpreadsheetApp.openById(spreadsheetId);
    } catch (err) {
      lastError = err;
      Utilities.sleep(1000);
    }
  }
  throw lastError;
}

function YahooProductName_readExcelIndex_(file) {
  const tempId = YahooProductName_convertExcelToSpreadsheet_(file);
  try {
    return YahooProductName_readSpreadsheetIndex_(tempId, file.getName() + '（一時変換）');
  } finally {
    try {
      DriveApp.getFileById(tempId).setTrashed(true);
    } catch (err) {
      Logger.log('一時変換ファイルを削除できませんでした: ' + err.message);
    }
  }
}

function YahooProductName_convertExcelToSpreadsheet_(file) {
  const token = ScriptApp.getOAuthToken();
  const resp = UrlFetchApp.fetch(
    'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(file.getId()) +
      '/copy?supportsAllDrives=true',
    {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        name: 'tmp_yahoo_product_name_' + Utilities.formatDate(
          new Date(),
          Session.getScriptTimeZone(),
          'yyyyMMddHHmmss'
        ),
        mimeType: 'application/vnd.google-apps.spreadsheet'
      }),
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    }
  );

  if (resp.getResponseCode() >= 400) {
    throw new Error('ExcelファイルをGoogleスプレッドシートへ一時変換できませんでした: ' + resp.getContentText());
  }

  const meta = JSON.parse(resp.getContentText());
  if (!meta.id) {
    throw new Error('Excelファイルの一時変換IDを取得できませんでした。');
  }

  return meta.id;
}

function YahooProductName_readCsvIndexFromFileId_(fileId) {
  const csv = YahooProductName_fetchCsvText_(fileId);
  return {
    name: csv.name,
    index: YahooProductName_buildCsvIndexFromText_(csv.text)
  };
}

function YahooProductName_fetchCsvText_(fileId) {
  const charsets = ['Shift_JIS', 'Windows-31J', 'UTF-8'];
  let lastError = null;

  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    for (let i = 0; i < charsets.length; i++) {
      const text = blob.getDataAsString(charsets[i]);
      if (YahooProductName_textHasRequiredHeaders_(text)) {
        return { name: file.getName(), text: text };
      }
    }
  } catch (err) {
    lastError = err;
  }

  try {
    const token = ScriptApp.getOAuthToken();
    const metaResp = UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fileId) +
        '?fields=name,mimeType,size&supportsAllDrives=true',
      {
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      }
    );
    const meta = metaResp.getResponseCode() < 400
      ? JSON.parse(metaResp.getContentText())
      : {};

    const mediaResp = UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fileId) +
        '?alt=media&supportsAllDrives=true',
      {
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      }
    );

    if (mediaResp.getResponseCode() >= 400) {
      throw new Error('Drive API ' + mediaResp.getResponseCode() + ': ' + mediaResp.getContentText());
    }

    for (let i = 0; i < charsets.length; i++) {
      const text = mediaResp.getContentText(charsets[i]);
      if (YahooProductName_textHasRequiredHeaders_(text)) {
        return { name: meta.name || fileId, text: text };
      }
    }
  } catch (err2) {
    lastError = err2;
  }

  throw new Error('CSVを読み取れませんでした: ' + (lastError ? lastError.message : 'unknown error'));
}

function YahooProductName_textHasRequiredHeaders_(text) {
  let firstRow = null;
  YahooProductName_forEachCsvRow_(String(text || ''), function(row) {
    firstRow = row;
    return false;
  });
  return YahooProductName_hasRequiredHeaders_([firstRow || []]);
}

function YahooProductName_readCsvRows_(blob) {
  const charsets = ['Shift_JIS', 'Windows-31J', 'UTF-8'];
  let lastError = null;

  for (let i = 0; i < charsets.length; i++) {
    try {
      const text = blob.getDataAsString(charsets[i]);
      const rows = Utilities.parseCsv(text);
      if (YahooProductName_hasRequiredHeaders_(rows)) return rows;
    } catch (err) {
      lastError = err;
    }
  }

  throw new Error('CSVを読み取れませんでした: ' + (lastError ? lastError.message : 'unknown error'));
}

function YahooProductName_hasRequiredHeaders_(rows) {
  if (!rows || rows.length === 0) return false;
  const headers = rows[0].map(function(v) {
    return YahooProductName_normHeader_(v);
  });
  return headers.indexOf('code') !== -1 &&
    headers.indexOf('name') !== -1 &&
    headers.indexOf('sub-code') !== -1;
}

function YahooProductName_buildCsvIndex_(rows) {
  if (!YahooProductName_hasRequiredHeaders_(rows)) {
    throw new Error('CSVに code / name / sub-code の見出しが見つかりません。');
  }

  const headers = rows[0].map(function(v) {
    return String(v || '').trim().toLowerCase();
  });
  const codeCol = headers.indexOf('code');
  const nameCol = headers.indexOf('name');
  const subCodeCol = headers.indexOf('sub-code');

  const index = {};

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const code = String(row[codeCol] || '').trim();
    const name = String(row[nameCol] || '').trim();
    const subCode = String(row[subCodeCol] || '').trim();
    if (!code || !name) continue;

    const candidates = YahooProductName_lookupCandidates_(code, subCode);
    const ref = subCode || code;

    candidates.forEach(function(candidate) {
      const key = YahooProductName_normCode_(candidate);
      if (!key) return;

      if (!index[key]) {
        index[key] = {
          names: [],
          refs: []
        };
      }

      if (index[key].names.indexOf(name) === -1) index[key].names.push(name);
      if (index[key].refs.indexOf(ref) === -1) index[key].refs.push(ref);
    });
  }

  return index;
}

function YahooProductName_buildCsvIndexFromText_(text) {
  const index = {};
  let codeCol = -1;
  let nameCol = -1;
  let subCodeCol = -1;

  YahooProductName_forEachCsvRow_(String(text || ''), function(row, rowIndex) {
    if (rowIndex === 0) {
      const headers = row.map(function(v) {
        return YahooProductName_normHeader_(v);
      });
      codeCol = headers.indexOf('code');
      nameCol = headers.indexOf('name');
      subCodeCol = headers.indexOf('sub-code');
      if (codeCol === -1 || nameCol === -1 || subCodeCol === -1) {
        throw new Error('CSVに code / name / sub-code の見出しが見つかりません。');
      }
      return true;
    }

    const code = String(row[codeCol] || '').trim();
    const name = String(row[nameCol] || '').trim();
    const subCode = String(row[subCodeCol] || '').trim();
    if (!code || !name) return true;

    const candidates = YahooProductName_lookupCandidates_(code, subCode);
    const ref = subCode || code;

    candidates.forEach(function(candidate) {
      const key = YahooProductName_normCode_(candidate);
      if (!key) return;

      if (!index[key]) {
        index[key] = {
          names: [],
          refs: []
        };
      }

      if (index[key].names.indexOf(name) === -1) index[key].names.push(name);
      if (index[key].refs.indexOf(ref) === -1) index[key].refs.push(ref);
    });

    return true;
  });

  return index;
}

function YahooProductName_forEachCsvRow_(text, callback) {
  let row = [];
  let cell = '';
  let inQuotes = false;
  let rowIndex = 0;

  for (let i = 0; i < text.length; i++) {
    const ch = text.charAt(i);

    if (ch === '"') {
      if (inQuotes && text.charAt(i + 1) === '"') {
        cell += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (ch === ',' && !inQuotes) {
      row.push(cell);
      cell = '';
      continue;
    }

    if ((ch === '\n' || ch === '\r') && !inQuotes) {
      if (ch === '\r' && text.charAt(i + 1) === '\n') i++;
      row.push(cell);

      const keepGoing = callback(row, rowIndex);
      rowIndex++;
      if (keepGoing === false) return;

      row = [];
      cell = '';
      continue;
    }

    cell += ch;
  }

  if (cell !== '' || row.length > 0) {
    row.push(cell);
    callback(row, rowIndex);
  }
}

function YahooProductName_normHeader_(value) {
  return String(value || '')
    .replace(/^\uFEFF/, '')
    .trim()
    .toLowerCase();
}

function YahooProductName_lookupCandidates_(code, subCode) {
  const codeText = String(code || '').trim();
  const subCodeText = String(subCode || '').trim();
  if (!subCodeText) {
    const parentCandidates = [codeText];
    YahooProductName_addTrailingNumberHyphenAlias_(parentCandidates, codeText);
    YahooProductName_addTrailingNumberNoHyphenAlias_(parentCandidates, codeText);
    return YahooProductName_uniqueCandidates_(parentCandidates);
  }

  const candidates = [subCodeText];

  const withoutLetterStockSuffix = subCodeText.replace(/[ab]$/i, '');
  const withoutNumberStockSuffix = subCodeText.replace(/[12]$/, '');

  if (withoutLetterStockSuffix !== subCodeText) candidates.push(withoutLetterStockSuffix);
  if (withoutNumberStockSuffix !== subCodeText) candidates.push(withoutNumberStockSuffix);

  YahooProductName_addNumericOptionAlias_(candidates, withoutLetterStockSuffix);
  YahooProductName_addNumericOptionAlias_(candidates, withoutNumberStockSuffix);
  YahooProductName_addNumericOptionAlias_(candidates, subCodeText);
  YahooProductName_addTrailingNumberHyphenAlias_(candidates, withoutLetterStockSuffix);
  YahooProductName_addTrailingNumberHyphenAlias_(candidates, withoutNumberStockSuffix);
  YahooProductName_addTrailingNumberHyphenAlias_(candidates, subCodeText);
  YahooProductName_addTrailingNumberNoHyphenAlias_(candidates, withoutLetterStockSuffix);
  YahooProductName_addTrailingNumberNoHyphenAlias_(candidates, withoutNumberStockSuffix);
  YahooProductName_addTrailingNumberNoHyphenAlias_(candidates, subCodeText);

  return YahooProductName_uniqueCandidates_(candidates);
}

function YahooProductName_uniqueCandidates_(candidates) {
  return candidates.filter(function(v, i, arr) {
    return v && arr.indexOf(v) === i;
  });
}

function YahooProductName_addNumericOptionAlias_(candidates, value) {
  const text = String(value || '').trim();
  const match = text.match(/^(.*-)(\d+)$/);
  if (!match) return;

  const num = Number(match[2]);
  if (!num || num < 1 || num > 26) return;

  const letter = String.fromCharCode(64 + num);
  candidates.push(match[1] + letter);
}

function YahooProductName_addTrailingNumberHyphenAlias_(candidates, value) {
  const text = String(value || '').trim();
  const match = text.match(/^(.*-[A-Za-z]+)(\d+)$/);
  if (!match) return;

  candidates.push(match[1] + '-' + match[2]);
}

function YahooProductName_addTrailingNumberNoHyphenAlias_(candidates, value) {
  const text = String(value || '').trim();
  const match = text.match(/^(.*-[A-Za-z]+)-(\d+)$/);
  if (!match) return;

  candidates.push(match[1] + match[2]);
}

function YahooProductName_normCode_(value) {
  let text = String(value || '').trim();
  if (text.normalize) text = text.normalize('NFKC');
  return text
    .replace(/[\s\u00A0\u3000]+/g, '')
    .replace(/[‐‑‒–—―ー－]/g, '-')
    .toUpperCase();
}
