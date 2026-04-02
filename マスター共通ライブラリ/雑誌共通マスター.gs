/**
 * 共通雑誌マスターセットアップ.gs
 * 商品登録_マスター共通 の Apps Script に配置する想定
 */

const 共通雑誌_シート = {
  master: '雑誌マスター（共通）',
  candidate: '雑誌マスター候補（共通）',
  edition: '版種ルール（雑誌共通）',
  type: 'タイプルール（雑誌共通）',
};

function 共通雑誌_正規化_(value) {
  return String(value || '')
    .trim()
    .replace(/[　\\s]+/g, ' ')
    .toUpperCase();
}

function 共通雑誌_ヘッダーMap_(sh) {
  const vals = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getDisplayValues()[0];
  const map = {};
  vals.forEach((h, i) => {
    if (h) map[String(h).trim()] = i + 1;
  });
  return map;
}

function 共通雑誌_比較文字列_(value) {
  return String(value || '').trim().replace(/[　\s]+/g, ' ').toUpperCase();
}

function 共通雑誌_正規化雑誌文字列_(value) {
  return String(value || '')
    .normalize('NFKC')
    .toUpperCase()
    .replace(/[（(].*?[)）]/g, ' ')
    .replace(/&/g, ' AND ')
    .replace(/\b(KOREA|TAIWAN|HONG\s*KONG|HONGKONG|CHINA|THAILAND|HK|TW|CN)\b/g, ' ')
    .replace(/(韓国版|韓國版|台湾版|臺灣版|中国版|中國版|香港版|タイ版)/g, ' ')
    .replace(/[^A-Z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function 共通雑誌_英字地域表記を除去_(value) {
  return String(value || '')
    .replace(/\b(?:KOREA|TAIWAN|HONG\s*KONG|HONGKONG|CHINA|THAILAND)\b/gi, ' ')
    .replace(/\b(?:HK|TW|CN|KR|TH)\b/gi, ' ')
    .replace(/(?:韓国版|韓國版|台湾版|臺灣版|中国版|中國版|香港版|タイ版)/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function 共通雑誌_英字表示キー_(value) {
  return String(value || '')
    .normalize('NFKC')
    .toUpperCase()
    .replace(/&/g, ' AND ')
    .replace(/[^A-Z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function 共通雑誌_英字表示名を整形_(value) {
  const text = String(value || '').trim();
  if (!text) return '';

  const map = {
    'MARIE CLAIRE': 'Marie Claire',
    'VOGUE': 'VOGUE',
    'BAZAAR': 'BAZAAR',
    'GQ': 'GQ',
    'ELLE': 'ELLE',
    'ESQUIRE': 'Esquire',
    'ALLURE': 'allure',
    'MENS HEALTH': "Men's Health",
    'MEN S HEALTH': "Men's Health",
    'L OFFICIEL': "L'OFFICIEL",
    'L OFFICIEL HOMMES': "L'OFFICIEL HOMMES",
    'THE BIG ISSUE': 'THE BIG ISSUE',
    'W': 'W'
  };
  return map[共通雑誌_英字表示キー_(text)] || text;
}

function 共通雑誌_表示カタカナ名を決定_(value) {
  return 共通雑誌_カタカナ地域表記を除去_(value);
}

function 共通雑誌_カタカナ地域表記を除去_(value) {
  return String(value || '')
    .trim()
    .replace(/(?:・|\s)*(?:台湾版|台湾|臺灣版|臺灣|コリア|韓国版|韓国|韓國版|韓國|香港版|香港|中国版|中国|中國版|中國|タイ版|タイ)$/u, '')
    .replace(/[・\s]+$/u, '')
    .trim();
}

function 共通雑誌_表示英字名を決定_(officialName) {
  const candidates = [];
  const seen = new Set();
  const pushCandidate = value => {
    const text = 共通雑誌_英字地域表記を除去_(value);
    const key = 共通雑誌_英字表示キー_(text);
    if (!text || !key || seen.has(key)) return;
    seen.add(key);
    candidates.push(text);
  };

  pushCandidate(officialName);
  if (!candidates.length) return String(officialName || '').trim();

  const latinCandidates = candidates.filter(text => /[A-Za-z]/.test(text));
  const pool = latinCandidates.length ? latinCandidates : candidates;
  let best = pool[0];
  let bestScore = -1;

  pool.forEach(candidate => {
    let score = 0;
    if (/[A-Za-z]/.test(candidate)) score += 100;
    if (/[A-Z]/.test(candidate) && /[a-z]/.test(candidate)) score += 25;
    else if (/^[A-Z0-9 '&+./-]+$/.test(candidate)) score += 20;
    else if (/^[a-z0-9 '&+./-]+$/.test(candidate)) score += 10;
    if (!/(?:KOREA|TAIWAN|HONG\s*KONG|HONGKONG|CHINA|THAILAND|HK|TW|CN|KR|TH)\b/i.test(candidate)) score += 40;
    score += Math.max(0, 40 - candidate.length);
    if (score > bestScore || (score == bestScore && candidate.length < best.length)) {
      best = candidate;
      bestScore = score;
    }
  });

  return 共通雑誌_英字表示名を整形_(best);
}

function 共通雑誌_別名一覧_(row, map) {
  const aliases = [];
  const seen = new Set();
  const pushIf = value => {
    const text = String(value || '').trim();
    if (!text || seen.has(text)) return;
    seen.add(text);
    aliases.push(text);
  };

  const officialName = String(row[(map['雑誌名（英字）'] || 2) - 1] || '').trim();
  pushIf(officialName);
  Object.keys(map)
    .filter(name => /^別名\d+$/.test(name))
    .sort((a, b) => map[a] - map[b])
    .forEach(name => pushIf(row[map[name] - 1]));

  const displayName = 共通雑誌_表示英字名を決定_(officialName);
  if (displayName) pushIf(displayName);

  return aliases;
}

function 共通雑誌_空き別名列_(master, masterMap, rowNum, value) {
  const aliasColumns = Object.keys(masterMap)
    .filter(name => /^別名\d+$/.test(name))
    .sort((a, b) => masterMap[a] - masterMap[b]);

  const currentValues = master.getRange(rowNum, 1, 1, master.getLastColumn()).getDisplayValues()[0];
  const normalized = 共通雑誌_正規化雑誌文字列_(value);
  for (let i = 0; i < aliasColumns.length; i += 1) {
    const col = masterMap[aliasColumns[i]];
    const current = String(currentValues[col - 1] || '').trim();
    if (current && 共通雑誌_正規化雑誌文字列_(current) === normalized) return -1;
    if (!current) return col;
  }
  return -1;
}

function 共通雑誌_略称候補_(name) {
  const cleaned = 共通雑誌_正規化雑誌文字列_(name).replace(/\s+/g, '');
  if (!cleaned) return '';
  return cleaned.slice(0, 8);
}

/* ===== 完全版 override start ===== */
const 共通雑誌_完全版設定 = {
  マスター列: [
    '略称コード', '雑誌名（英字）', '雑誌名（カタカナ）', '基本キー型',
    '通常タイプ結合記号', '版種ありタイプ結合記号', '対応言語',
    '別名1', '別名2', '別名3', '別名4', '別名5', '備考'
  ],
  候補列: [
    'ステータス', '反映', '略称コード候補',
    '雑誌名（英字）', '雑誌名（カタカナ）', '基本キー型候補',
    '通常タイプ結合記号', '版種ありタイプ結合記号', '対応言語',
    '別名1', '別名2', '別名3', '備考',
    '元ファイル', '元シート', '元行',
    '初回検出日時', '最終検出日時', '正規化キー'
  ],
  ルール列: ['判定キーワード', 'コードsuffix', 'タイトル表示', '優先順位', '有効'],
  ステータス候補: ['未対応', '確認中', '登録待ち', 'マスター登録済み', '無視'],
  基本キー型候補: ['年月型', '号数型'],
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('共通雑誌マスター')
    .addItem('共通雑誌マスターを初期化', '共通雑誌マスターを初期化')
    .addItem('雑誌名を国名なしへ統合', '共通雑誌マスターを国名なしへ統合')
    .addItem('候補を正式マスターへ反映', '共通雑誌候補を正式反映')
    .addSeparator()
    .addItem('候補シートへテスト追加', '共通雑誌候補テスト追加')
    .addToUi();
}
function 共通雑誌マスターを初期化() {
  const ss = SpreadsheetApp.getActive();
  const masterSh = 共通雑誌_マスター作成_(ss);
  共通雑誌_候補シート作成_(ss);
  const masterAdded = 共通雑誌マスターへ不足分を追加_(masterSh, 共通雑誌初期データ_());
  const editionAdded = 共通版種ルールシートを不足分追加_(ss);
  const typeAdded = 共通タイプルールシートを不足分追加_(ss);
  ss.toast(`共通雑誌マスター初期化完了 / 雑誌:${masterAdded}件 版種:${editionAdded}件 タイプ:${typeAdded}件`, '完了', 5);
}

function 共通雑誌マスターを国名なしへ統合() {
  const ss = SpreadsheetApp.getActive();
  const masterSh = 共通雑誌_マスター作成_(ss);
  const sourceRows = 共通雑誌_既存マスター行を移行_(masterSh);
  if (!sourceRows.length) {
    ss.toast('統合対象のマスターデータがありません', '完了', 3);
    return;
  }

  const mergedRows = 共通雑誌_マスター行を国名なしへ統合_(sourceRows);
  const headers = 共通雑誌_完全版設定.マスター列;
  masterSh.clearContents();
  共通雑誌_必要サイズを確保_(masterSh, Math.max(mergedRows.length + 1, 2), headers.length);
  masterSh.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (mergedRows.length) {
    masterSh.getRange(2, 1, mergedRows.length, headers.length).setValues(mergedRows);
  }
  共通雑誌_マスター作成_(ss);
  ss.toast(`国名なし統合完了 / ${sourceRows.length}件 → ${mergedRows.length}件`, '完了', 5);
}

function 共通雑誌_マスター作成_(ss) {
  const headers = 共通雑誌_完全版設定.マスター列;
  let sh = ss.getSheetByName(共通雑誌_シート.master);
  const needMigrate = sh && !共通雑誌_ヘッダー一致_(sh, headers);
  const migratedRows = needMigrate ? 共通雑誌_既存マスター行を移行_(sh) : [];

  if (!sh) sh = ss.insertSheet(共通雑誌_シート.master);
  共通雑誌_必要サイズを確保_(sh, Math.max(migratedRows.length + 1, 2), headers.length);

  if (!共通雑誌_ヘッダー一致_(sh, headers)) {
    sh.clearContents();
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (migratedRows.length) sh.getRange(2, 1, migratedRows.length, headers.length).setValues(migratedRows);
  }

  sh.getRange(1, 1, 1, headers.length)
    .setBackground('#4a86e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  [110, 220, 180, 90, 120, 150, 100, 180, 180, 180, 180, 180, 260]
    .forEach((w, i) => sh.setColumnWidth(i + 1, w));

  const maxRows = Math.max(sh.getMaxRows() - 1, Math.max(sh.getLastRow() + 200, 1000));
  共通雑誌_必要サイズを確保_(sh, maxRows + 1, headers.length);
  sh.getRange(2, 4, maxRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(共通雑誌_完全版設定.基本キー型候補, true)
      .build()
  );
  return sh;
}

function 共通雑誌マスターへ不足分を追加_(sh, rows) {
  if (!rows || !rows.length) return 0;
  const existing = {};
  if (sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues().forEach(r => {
      const key = `${共通雑誌_正規化_(r[0])}__${共通雑誌_正規化_(r[1])}`;
      if (key !== '__') existing[key] = true;
    });
  }

  const addRows = [];
  rows.forEach(row => {
    const key = `${共通雑誌_正規化_(row[0])}__${共通雑誌_正規化_(row[1])}`;
    if (!existing[key]) {
      addRows.push(row);
      existing[key] = true;
    }
  });

  if (addRows.length) {
    共通雑誌_必要サイズを確保_(sh, sh.getLastRow() + addRows.length, 共通雑誌_完全版設定.マスター列.length);
    sh.getRange(sh.getLastRow() + 1, 1, addRows.length, addRows[0].length).setValues(addRows);
  }
  return addRows.length;
}

function 共通雑誌_候補シート作成_(ss) {
  const headers = 共通雑誌_完全版設定.候補列;
  let sh = ss.getSheetByName(共通雑誌_シート.candidate);
  const needMigrate = sh && !共通雑誌_ヘッダー一致_(sh, headers);
  const migratedRows = needMigrate ? 共通雑誌_既存候補行を移行_(sh) : [];

  if (!sh) sh = ss.insertSheet(共通雑誌_シート.candidate);
  共通雑誌_必要サイズを確保_(sh, Math.max(migratedRows.length + 1, 2), headers.length);

  if (!共通雑誌_ヘッダー一致_(sh, headers)) {
    sh.clearContents();
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (migratedRows.length) sh.getRange(2, 1, migratedRows.length, headers.length).setValues(migratedRows);
  }

  sh.getRange(1, 1, 1, 13)
    .setBackground('#e69138')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sh.getRange(1, 14, 1, 6)
    .setBackground('#999999')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);
  [120, 45, 120, 220, 180, 110, 90, 140, 90, 180, 180, 180, 320, 160, 120, 70, 140, 140, 160]
    .forEach((w, i) => sh.setColumnWidth(i + 1, w));

  const maxRows = Math.max(sh.getMaxRows() - 1, Math.max(sh.getLastRow() + 200, 1000));
  共通雑誌_必要サイズを確保_(sh, maxRows + 1, headers.length);
  sh.getRange(2, 1, maxRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(共通雑誌_完全版設定.ステータス候補, true).build()
  );
  sh.getRange(2, 2, maxRows, 1).insertCheckboxes();
  sh.getRange(2, 6, maxRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(共通雑誌_完全版設定.基本キー型候補, true).build()
  );

  const range = sh.getRange(2, 1, maxRows, 13);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$A2="未対応"').setBackground('#fff2cc').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$A2="確認中"').setBackground('#fce5cd').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$A2="登録待ち"').setBackground('#cfe2f3').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$A2="マスター登録済み"').setBackground('#d9ead3').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$A2="無視"').setBackground('#efefef').setFontColor('#999999').setRanges([range]).build(),
  ]);
  sh.hideColumns(19);
  return sh;
}

function 共通版種ルールシートを不足分追加_(ss) {
  return 共通雑誌_ルールシート作成_(ss, 共通雑誌_シート.edition, [
    ['SPECIAL EDITION', 'SP', 'SPECIAL EDITION', 100, true],
    ['SP EDITION', 'SP', 'SPECIAL EDITION', 95, true],
    ['SPECIAL', 'SP', 'SPECIAL EDITION', 90, true],
    ['LIMITED EDITION', 'LE', 'LIMITED EDITION', 85, true],
    ['特装版', 'SP', '特装版', 80, true],
    ['特別版', 'SP', '特別版', 75, true],
    ['限定版', 'LE', '限定版', 70, true],
    ['超值版', '', '超値版', 65, true],
    ['超値版', '', '超値版', 64, true],
  ], '#e69138');
}

function 共通タイプルールシートを不足分追加_(ss) {
  return 共通雑誌_ルールシート作成_(ss, 共通雑誌_シート.type, [
    ['Aタイプ', 'A', 'Aタイプ', 100, true], ['Bタイプ', 'B', 'Bタイプ', 100, true],
    ['Cタイプ', 'C', 'Cタイプ', 100, true], ['Dタイプ', 'D', 'Dタイプ', 100, true],
    ['Eタイプ', 'E', 'Eタイプ', 100, true], ['TYPE A', 'A', 'Aタイプ', 90, true],
    ['TYPE B', 'B', 'Bタイプ', 90, true], ['TYPE C', 'C', 'Cタイプ', 90, true],
    ['TYPE D', 'D', 'Dタイプ', 90, true], ['TYPE E', 'E', 'Eタイプ', 90, true],
    ['VER.A', 'A', 'Aタイプ', 80, true], ['VER.B', 'B', 'Bタイプ', 80, true],
    ['VER.C', 'C', 'Cタイプ', 80, true], ['VER.D', 'D', 'Dタイプ', 80, true],
    ['VER.E', 'E', 'Eタイプ', 80, true], ['-A', 'A', 'Aタイプ', 70, true],
    ['-B', 'B', 'Bタイプ', 70, true], ['-C', 'C', 'Cタイプ', 70, true],
    ['-D', 'D', 'Dタイプ', 70, true], ['-E', 'E', 'Eタイプ', 70, true],
  ], '#6aa84f');
}

function 共通雑誌_ルールシート作成_(ss, sheetName, rows, headerColor) {
  const headers = 共通雑誌_完全版設定.ルール列;
  let sh = ss.getSheetByName(sheetName);
  const needMigrate = sh && !共通雑誌_ヘッダー一致_(sh, headers);
  const migratedRows = needMigrate ? 共通雑誌_既存ルール行を移行_(sh) : [];

  if (!sh) sh = ss.insertSheet(sheetName);
  共通雑誌_必要サイズを確保_(sh, Math.max(migratedRows.length + 1, 2), headers.length);
  if (!共通雑誌_ヘッダー一致_(sh, headers)) {
    sh.clearContents();
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (migratedRows.length) sh.getRange(2, 1, migratedRows.length, headers.length).setValues(migratedRows);
  }

  sh.getRange(1, 1, 1, headers.length).setBackground(headerColor).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  [220, 110, 220, 90, 60].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  const existing = {};
  if (sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues().forEach(r => {
      const key = `${共通雑誌_正規化_(r[0])}__${共通雑誌_正規化_(r[1])}`;
      if (key !== '__') existing[key] = true;
    });
  }

  const addRows = [];
  rows.forEach(row => {
    const key = `${共通雑誌_正規化_(row[0])}__${共通雑誌_正規化_(row[1])}`;
    if (!existing[key]) {
      addRows.push(row);
      existing[key] = true;
    }
  });

  if (addRows.length) {
    const startRow = sh.getLastRow() + 1;
    共通雑誌_必要サイズを確保_(sh, startRow + addRows.length, headers.length);
    sh.getRange(startRow, 1, addRows.length, headers.length).setValues(addRows);
  }

  const maxRows = Math.max(sh.getMaxRows() - 1, Math.max(sh.getLastRow() + 100, 500));
  共通雑誌_必要サイズを確保_(sh, maxRows + 1, headers.length);
  sh.getRange(2, 5, maxRows, 1).insertCheckboxes();
  if (sh.getLastRow() >= 2) {
    const bools = sh.getRange(2, 5, sh.getLastRow() - 1, 1).getValues().map(row => [row[0] !== false]);
    sh.getRange(2, 5, bools.length, 1).setValues(bools);
  }
  return addRows.length;
}

function 共通雑誌候補を追加_(params) {
  return 共通雑誌候補をSSへ追加_(SpreadsheetApp.getActive(), params);
}

function 共通雑誌候補を外部から追加_(masterSpreadsheetId, params) {
  return 共通雑誌候補をSSへ追加_(SpreadsheetApp.openById(masterSpreadsheetId), params);
}

function 共通雑誌候補をSSへ追加_(ss, params) {
  const masterSh = 共通雑誌_マスター作成_(ss);
  const candidateSh = 共通雑誌_候補シート作成_(ss);
  const candidateMap = 共通雑誌_ヘッダーMap_(candidateSh);

  const nameEn = String(params.nameEn || params.雑誌名 || params['雑誌名（英字）'] || '').trim();
  if (!nameEn) return -1;

  const language = 共通雑誌_言語コード_(params.langs || params.対応言語 || params.言語 || '');
  if (共通雑誌マスターに存在するか_(masterSh, nameEn, language)) return -1;

  const key = String(params.normalizedKey || 共通雑誌_候補キー_(language, nameEn)).trim();
  const note = 共通雑誌_備考結合_(
    params.note || params.備考 || '',
    params.原題 ? `原題:${params.原題}` : '',
    params.博客來URL ? `URL:${params.博客來URL}` : '',
    params.博客來商品コード ? `コード:${params.博客來商品コード}` : ''
  );
  const existingRow = 共通雑誌候補_既存行を取得_(candidateSh, key);
  const now = new Date();

  if (existingRow >= 2) {
    if (candidateMap['最終検出日時']) candidateSh.getRange(existingRow, candidateMap['最終検出日時']).setValue(now);
    if (candidateMap['備考'] && note) {
      const current = candidateSh.getRange(existingRow, candidateMap['備考']).getDisplayValue();
      candidateSh.getRange(existingRow, candidateMap['備考']).setValue(共通雑誌_備考結合_(current, note));
    }
    return existingRow;
  }

  const rowNum = Math.max(candidateSh.getLastRow() + 1, 2);
  共通雑誌_必要サイズを確保_(candidateSh, rowNum, 共通雑誌_完全版設定.候補列.length);
  candidateSh.getRange(rowNum, 1, 1, 19).setValues([[
    共通雑誌_候補状態を正規化_(params.status || '未対応'),
    false,
    String(params.code || params.略称コード候補 || 共通雑誌_略称候補_(nameEn)).trim(),
    nameEn,
    String(params.nameJa || params.nameKana || params['雑誌名（カタカナ）'] || '').trim(),
    String(params.keyType || params.基本キー型候補 || '年月型').trim() || '年月型',
    String(params.normalJoin || params['通常タイプ結合記号'] || '').trim(),
    String(params.specialJoin || params['版種ありタイプ結合記号'] || '-').trim() || '-',
    language,
    String(params.alias1 || '').trim(),
    String(params.alias2 || '').trim(),
    String(params.alias3 || '').trim(),
    note,
    String(params.sourceFile || params.元ファイル || '').trim(),
    String(params.sourceSheet || params.元シート || '').trim(),
    String(params.sourceRow || params.rowNum || params.元行 || '').trim(),
    now,
    now,
    key,
  ]]);
  return rowNum;
}

function 共通雑誌候補を正式反映() {
  return 共通雑誌候補をマスターへ反映();
}

function 共通雑誌候補をマスターへ反映() {
  const ss = SpreadsheetApp.getActive();
  const masterSh = 共通雑誌_マスター作成_(ss);
  const candidateSh = 共通雑誌_候補シート作成_(ss);
  if (candidateSh.getLastRow() < 2) {
    ss.toast('候補データがありません', '完了', 3);
    return;
  }

  const masterMap = 共通雑誌_ヘッダーMap_(masterSh);
  const candidateMap = 共通雑誌_ヘッダーMap_(candidateSh);
  const masterValues = masterSh.getLastRow() >= 2
    ? masterSh.getRange(2, 1, masterSh.getLastRow() - 1, masterSh.getLastColumn()).getDisplayValues()
    : [];
  const candidateValues = candidateSh.getRange(2, 1, candidateSh.getLastRow() - 1, 19).getValues();

  let added = 0;
  let aliasAdded = 0;
  let matched = 0;

  candidateValues.forEach((row, index) => {
    const rowNum = index + 2;
    const status = 共通雑誌_候補状態を正規化_(row[0]);
    const checked = row[1] === true;
    const nameEn = String(row[3] || '').trim();
    if (!checked || !nameEn || status === '無視') return;

    const language = 共通雑誌_言語コード_(row[8]);
    const code = String(row[2] || '').trim() || 共通雑誌_略称候補_(nameEn);

    if (共通雑誌マスターに存在するか_(masterSh, nameEn, language)) {
      共通雑誌_候補行を更新_(candidateSh, candidateMap, rowNum, 'マスター登録済み', false, '既存マスターに一致');
      matched += 1;
      return;
    }

    const matchedRow = 共通雑誌_最良マスター行を探す_(masterValues, masterMap, nameEn, language, code);
    if (matchedRow && !matchedRow.exact && matchedRow.languageMatched && matchedRow.score >= 360) {
      const aliasCol = 共通雑誌_空き別名列_(masterSh, masterMap, matchedRow.rowIndex + 2, nameEn);
      if (aliasCol > 0) {
        masterSh.getRange(matchedRow.rowIndex + 2, aliasCol).setValue(nameEn);
        masterValues[matchedRow.rowIndex][aliasCol - 1] = nameEn;
        共通雑誌_候補行を更新_(candidateSh, candidateMap, rowNum, 'マスター登録済み', false, `既存マスター ${matchedRow.displayName || matchedRow.officialName} に別名追加`);
        aliasAdded += 1;
        return;
      }
      if (aliasCol === -1) {
        共通雑誌_候補行を更新_(candidateSh, candidateMap, rowNum, 'マスター登録済み', false, `既存マスター ${matchedRow.displayName || matchedRow.officialName} に一致`);
        matched += 1;
        return;
      }
    }

    const appendRow = [
      code,
      nameEn,
      String(row[4] || '').trim(),
      String(row[5] || '年月型').trim() || '年月型',
      String(row[6] || '').trim(),
      String(row[7] || '-').trim() || '-',
      language,
      String(row[9] || '').trim(),
      String(row[10] || '').trim(),
      String(row[11] || '').trim(),
      '',
      '',
      String(row[12] || '').trim() || '候補シートから取込',
    ];

    masterSh.appendRow(appendRow);
    masterValues.push(appendRow.map(v => String(v == null ? '' : v)));
    共通雑誌_候補行を更新_(candidateSh, candidateMap, rowNum, 'マスター登録済み', false, '新規マスターとして追加');
    added += 1;
  });

  ss.toast(`正式反映完了 / 追加:${added}件 別名追加:${aliasAdded}件 一致:${matched}件`, '完了', 5);
}

function 共通雑誌候補テスト追加() {
  共通雑誌候補を追加_({
    code: 'TESTMAG',
    nameEn: 'TEST MAGAZINE',
    nameJa: 'テストマガジン',
    keyType: '年月型',
    langs: 'TW',
    alias1: 'TEST MAG',
    note: 'テスト追加',
    sourceFile: 'manual',
    sourceSheet: 'UI',
  });
  SpreadsheetApp.getActive().toast('候補シートへテスト追加しました', '完了', 3);
}
function 共通雑誌_ヘッダー一致_(sh, headers) {
  if (!sh || sh.getLastColumn() < headers.length) return false;
  const current = sh.getRange(1, 1, 1, headers.length).getDisplayValues()[0].map(v => String(v || '').trim());
  return headers.every((header, index) => current[index] === header);
}

function 共通雑誌_必要サイズを確保_(sh, rows, cols) {
  if (cols > sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());
  if (rows > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
}

function 共通雑誌_既存マスター行を移行_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  const map = 共通雑誌_行マップ取得_(sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getDisplayValues()[0]);
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues()
    .filter(共通雑誌_行に値があるか_)
    .map(row => [
      共通雑誌_旧行値_(row, map, ['略称コード']),
      共通雑誌_旧行値_(row, map, ['雑誌名（英字）', '雑誌名候補']),
      共通雑誌_旧行値_(row, map, ['雑誌名（カタカナ）']),
      共通雑誌_旧行値_(row, map, ['基本キー型', 'コード型']) || '年月型',
      共通雑誌_旧行値_(row, map, ['通常タイプ結合記号']),
      共通雑誌_旧行値_(row, map, ['版種ありタイプ結合記号']) || '-',
      共通雑誌_旧行値_(row, map, ['対応言語']),
      共通雑誌_旧行値_(row, map, ['別名1']),
      共通雑誌_旧行値_(row, map, ['別名2']),
      共通雑誌_旧行値_(row, map, ['別名3']),
      共通雑誌_旧行値_(row, map, ['別名4']),
      共通雑誌_旧行値_(row, map, ['別名5']),
      共通雑誌_旧行値_(row, map, ['備考']),
    ])
    .filter(row => String(row[0] || '').trim() || String(row[1] || '').trim());
}

function 共通雑誌_既存候補行を移行_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  const map = 共通雑誌_行マップ取得_(sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getDisplayValues()[0]);
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues()
    .filter(共通雑誌_行に値があるか_)
    .map(row => {
      const nameEn = String(共通雑誌_旧行値_(row, map, ['雑誌名（英字）', '雑誌名候補']) || '').trim();
      const language = String(共通雑誌_旧行値_(row, map, ['対応言語']) || '').trim();
      const note = 共通雑誌_備考結合_(
        共通雑誌_旧行値_(row, map, ['備考']),
        共通雑誌_旧行値_(row, map, ['サンプル商品コード']) ? `コード:${共通雑誌_旧行値_(row, map, ['サンプル商品コード'])}` : '',
        共通雑誌_旧行値_(row, map, ['サンプル商品名']) ? `商品名:${共通雑誌_旧行値_(row, map, ['サンプル商品名'])}` : '',
        共通雑誌_旧行値_(row, map, ['原題タイトルサンプル']) ? `原題:${共通雑誌_旧行値_(row, map, ['原題タイトルサンプル'])}` : '',
        共通雑誌_旧行値_(row, map, ['博客來URL']) ? `URL:${共通雑誌_旧行値_(row, map, ['博客來URL'])}` : '',
        共通雑誌_旧行値_(row, map, ['出現数']) ? `出現数:${共通雑誌_旧行値_(row, map, ['出現数'])}` : '',
        共通雑誌_旧行値_(row, map, ['信頼度']) ? `信頼度:${共通雑誌_旧行値_(row, map, ['信頼度'])}` : ''
      );
      const firstSeen = 共通雑誌_旧行値_(row, map, ['初回検出日時', '登録日時']) || '';
      const lastSeen = 共通雑誌_旧行値_(row, map, ['最終検出日時', '登録日時']) || firstSeen || '';
      return [
        共通雑誌_候補状態を正規化_(共通雑誌_旧行値_(row, map, ['ステータス', '状態']) || '未対応'),
        共通雑誌_真偽値_(共通雑誌_旧行値_(row, map, ['反映'])),
        共通雑誌_旧行値_(row, map, ['略称コード候補']),
        nameEn,
        共通雑誌_旧行値_(row, map, ['雑誌名（カタカナ）']),
        共通雑誌_旧行値_(row, map, ['基本キー型候補', '基本キー型', 'コード型']) || '年月型',
        共通雑誌_旧行値_(row, map, ['通常タイプ結合記号']),
        共通雑誌_旧行値_(row, map, ['版種ありタイプ結合記号']) || '-',
        language,
        共通雑誌_旧行値_(row, map, ['別名1']),
        共通雑誌_旧行値_(row, map, ['別名2']),
        共通雑誌_旧行値_(row, map, ['別名3']),
        note,
        共通雑誌_旧行値_(row, map, ['元ファイル']),
        共通雑誌_旧行値_(row, map, ['元シート']),
        共通雑誌_旧行値_(row, map, ['元行']),
        firstSeen,
        lastSeen,
        共通雑誌_旧行値_(row, map, ['正規化キー', '重複キー']) || 共通雑誌_候補キー_(language, nameEn),
      ];
    })
    .filter(row => String(row[3] || '').trim());
}

function 共通雑誌_既存ルール行を移行_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  const map = 共通雑誌_行マップ取得_(sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getDisplayValues()[0]);
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues()
    .filter(共通雑誌_行に値があるか_)
    .map(row => [
      共通雑誌_旧行値_(row, map, ['判定キーワード']),
      共通雑誌_旧行値_(row, map, ['コードsuffix']),
      共通雑誌_旧行値_(row, map, ['タイトル表示']),
      Number(共通雑誌_旧行値_(row, map, ['優先順位']) || 0),
      共通雑誌_真偽値_(共通雑誌_旧行値_(row, map, ['有効'])) !== false,
    ])
    .filter(row => String(row[0] || '').trim());
}

function 共通雑誌_行マップ取得_(headers) {
  const map = {};
  (headers || []).forEach((header, index) => {
    const key = String(header || '').trim();
    if (key && typeof map[key] !== 'number') map[key] = index;
  });
  return map;
}

function 共通雑誌_旧行値_(row, map, names) {
  for (let index = 0; index < names.length; index += 1) {
    const key = names[index];
    if (typeof map[key] === 'number') return row[map[key]];
  }
  return '';
}

function 共通雑誌_行に値があるか_(row) {
  return (row || []).some(value => {
    if (value === false) return true;
    if (value == null) return false;
    return String(value).trim() !== '';
  });
}

function 共通雑誌_言語コード_(value) {
  const raw = String(value || '').trim();
  const normalized = raw.toUpperCase();
  const map = {
    TW: 'TW', CN: 'CN', HK: 'HK', TH: 'TH', KR: 'KR', JP: 'JP', US: 'US', EN: 'EN',
    '台湾': 'TW', '臺灣': 'TW', '中国': 'CN', '中國': 'CN', '香港': 'HK',
    'タイ': 'TH', '泰国': 'TH', '泰國': 'TH', '韓国': 'KR', '韓國': 'KR',
    '日本': 'JP', 'アメリカ': 'US', '英語': 'EN'
  };
  return map[normalized] || map[raw] || '';
}

function 共通雑誌_対応言語配列_(raw) {
  return String(raw || '').split(',').map(value => 共通雑誌_言語コード_(value)).filter(Boolean);
}

function 共通雑誌_言語一致_(supported, language) {
  const code = 共通雑誌_言語コード_(language);
  if (!supported || !supported.length || !code) return true;
  return supported.includes(code);
}

function 共通雑誌_最良マスター行を探す_(masterValues, masterMap, candidateName, language, abbr) {
  const targetNorm = 共通雑誌_正規化雑誌文字列_(candidateName);
  if (!targetNorm) return null;

  let best = null;
  for (let index = 0; index < masterValues.length; index += 1) {
    const row = masterValues[index];
    const officialName = String(row[(masterMap['雑誌名（英字）'] || 2) - 1] || '').trim();
    const displayName = 共通雑誌_表示英字名を決定_(officialName);
    const officialNorm = 共通雑誌_正規化雑誌文字列_(officialName);
    const aliases = 共通雑誌_別名一覧_(row, masterMap);
    const supported = 共通雑誌_対応言語配列_(row[(masterMap['対応言語'] || 7) - 1]);
    const rowAbbr = String(row[(masterMap['略称コード'] || 1) - 1] || '').trim();
    const languageMatched = 共通雑誌_言語一致_(supported, language);

    let score = -1;
    let exact = false;
    aliases.forEach(alias => {
      const aliasNorm = 共通雑誌_正規化雑誌文字列_(alias);
      if (!aliasNorm) return;
      if (aliasNorm === targetNorm) {
        exact = true;
        score = Math.max(score, 1000);
      } else if (targetNorm.startsWith(aliasNorm) || aliasNorm.startsWith(targetNorm)) {
        score = Math.max(score, 320 + Math.min(aliasNorm.length, targetNorm.length));
      } else if (targetNorm.includes(aliasNorm) || aliasNorm.includes(targetNorm)) {
        score = Math.max(score, 240 + Math.min(aliasNorm.length, targetNorm.length));
      }
    });

    if (score < 0) continue;
    if (abbr && rowAbbr && abbr == rowAbbr) score += 40;
    if (languageMatched) score += 60;
    if (!best || score > best.score) {
      best = { rowIndex: index, score, exact, officialName, displayName, officialNorm, languageMatched };
    }
  }
  return best;
}

function 共通雑誌_候補状態を正規化_(value) {
  const text = String(value || '').trim();
  const map = {
    '': '未対応',
    '自動候補': '登録待ち',
    '要確認': '確認中',
    '取込済み': 'マスター登録済み',
    '既存一致': 'マスター登録済み',
    '別名追加済み': 'マスター登録済み',
    '除外候補': '無視',
  };
  return map[text] || (共通雑誌_完全版設定.ステータス候補.includes(text) ? text : '未対応');
}

function 共通雑誌_真偽値_(value) {
  if (value === true || value === false) return value;
  const text = String(value || '').trim().toUpperCase();
  return ['TRUE', '1', 'YES', 'ON', 'CHECKED'].includes(text);
}

function 共通雑誌_備考結合_() {
  const parts = [];
  const seen = new Set();
  for (let index = 0; index < arguments.length; index += 1) {
    const text = String(arguments[index] == null ? '' : arguments[index]).trim();
    if (!text || seen.has(text)) continue;
    seen.add(text);
    parts.push(text);
  }
  return parts.join(' / ');
}

function 共通雑誌_候補キー_(language, nameEn) {
  return `${共通雑誌_言語コード_(language)}|${共通雑誌_正規化雑誌文字列_(nameEn)}`;
}

function 共通雑誌候補_既存行を取得_(sh, key) {
  if (!sh || sh.getLastRow() < 2 || !key) return -1;
  const map = 共通雑誌_ヘッダーMap_(sh);
  const keyCol = map['正規化キー'] || 19;
  const values = sh.getRange(2, keyCol, sh.getLastRow() - 1, 1).getDisplayValues().flat();
  for (let index = 0; index < values.length; index += 1) {
    if (String(values[index] || '').trim() === key) return index + 2;
  }
  return -1;
}

function 共通雑誌_候補行を更新_(candidateSh, candidateMap, rowNum, status, checked, note) {
  if (candidateMap['ステータス']) candidateSh.getRange(rowNum, candidateMap['ステータス']).setValue(status);
  if (candidateMap['反映']) candidateSh.getRange(rowNum, candidateMap['反映']).setValue(checked === true);
  if (candidateMap['最終検出日時']) candidateSh.getRange(rowNum, candidateMap['最終検出日時']).setValue(new Date());
  if (candidateMap['備考'] && note) {
    const current = candidateSh.getRange(rowNum, candidateMap['備考']).getDisplayValue();
    candidateSh.getRange(rowNum, candidateMap['備考']).setValue(共通雑誌_備考結合_(current, note));
  }
}

function 共通雑誌マスターに存在するか_(sh, nameEn, language) {
  if (!sh || sh.getLastRow() < 2) return false;
  const target = 共通雑誌_正規化雑誌文字列_(nameEn);
  if (!target) return false;
  const map = 共通雑誌_ヘッダーMap_(sh);
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getDisplayValues().some(row => {
    const aliases = 共通雑誌_別名一覧_(row, map);
    const supported = 共通雑誌_対応言語配列_(row[(map['対応言語'] || 7) - 1]);
    if (!共通雑誌_言語一致_(supported, language)) return false;
    return aliases.some(alias => 共通雑誌_正規化雑誌文字列_(alias) === target);
  });
}

function 共通雑誌候補に存在するか_(sh, nameEn, language) {
  return 共通雑誌候補_既存行を取得_(sh, 共通雑誌_候補キー_(language, nameEn)) >= 2;
}

function 共通雑誌_配列へ重複なし追加_(list, value) {
  const text = String(value || '').trim();
  if (text && !list.includes(text)) list.push(text);
}

function 共通雑誌_言語一覧を連結_(values) {
  const order = ['TW', 'KR', 'CN', 'HK', 'TH', 'JP', 'US', 'EN'];
  const unique = [];
  (values || []).forEach(value => {
    const code = 共通雑誌_言語コード_(value);
    if (code && !unique.includes(code)) unique.push(code);
  });
  unique.sort((a, b) => {
    const ai = order.indexOf(a);
    const bi = order.indexOf(b);
    if (ai < 0 && bi < 0) return a.localeCompare(b);
    if (ai < 0) return 1;
    if (bi < 0) return -1;
    return ai - bi;
  });
  return unique.join(',');
}

function 共通雑誌_略称優先度_(code, displayName) {
  const text = String(code || '').trim().toUpperCase();
  if (!text) return -1;
  let score = 0;
  if (!/(?:TW|CN|HK|KR|TH)$/.test(text)) score += 100;
  score += Math.max(0, 40 - text.length);
  if (/^[A-Z0-9]+$/.test(text)) score += 10;
  const expected = String(共通雑誌_略称候補_(displayName) || '').toUpperCase();
  if (expected && text === expected) score += 20;
  return score;
}

function 共通雑誌_優先略称を選ぶ_(codes, displayName) {
  const values = Array.from(new Set((codes || []).map(value => String(value || '').trim()).filter(Boolean)));
  if (!values.length) return 共通雑誌_略称候補_(displayName);

  let best = values[0];
  let bestScore = 共通雑誌_略称優先度_(best, displayName);
  values.forEach(code => {
    const score = 共通雑誌_略称優先度_(code, displayName);
    if (score > bestScore || (score === bestScore && code.length < best.length)) {
      best = code;
      bestScore = score;
    }
  });
  return best;
}

function 共通雑誌_マスター行を国名なしへ統合_(rows) {
  const groups = new Map();

  (rows || []).forEach(row => {
    const code = String(row[0] || '').trim();
    const officialName = String(row[1] || '').trim();
    const kanaName = String(row[2] || '').trim();
    const keyType = String(row[3] || '年月型').trim() || '年月型';
    const normalJoin = String(row[4] || '').trim();
    const specialJoin = String(row[5] || '-').trim() || '-';
    const displayName = 共通雑誌_表示英字名を決定_(officialName) || officialName;
    const displayKana = 共通雑誌_表示カタカナ名を決定_(kanaName);
    const groupKey = `${共通雑誌_英字表示キー_(displayName || officialName)}__${keyType}`;
    if (!groupKey) return;

    if (!groups.has(groupKey)) {
      groups.set(groupKey, {
        displayName: displayName || officialName,
        displayKana: displayKana,
        keyType,
        normalJoin: normalJoin,
        specialJoin: specialJoin,
        codes: [],
        languages: [],
        aliases: [],
        notes: [],
      });
    }

    const group = groups.get(groupKey);
    if (!group.displayKana && displayKana) group.displayKana = displayKana;
    if (!group.normalJoin && normalJoin) group.normalJoin = normalJoin;
    if ((!group.specialJoin || group.specialJoin === '-') && specialJoin) group.specialJoin = specialJoin;

    共通雑誌_配列へ重複なし追加_(group.codes, code);
    共通雑誌_対応言語配列_(row[6]).forEach(language => 共通雑誌_配列へ重複なし追加_(group.languages, language));
    [officialName, row[7], row[8], row[9], row[10], row[11]].forEach(alias => {
      const text = String(alias || '').trim();
      if (!text) return;
      if (共通雑誌_比較文字列_(text) === 共通雑誌_比較文字列_(group.displayName)) return;
      共通雑誌_配列へ重複なし追加_(group.aliases, text);
    });
    共通雑誌_配列へ重複なし追加_(group.notes, String(row[12] || '').trim());
  });

  return Array.from(groups.values())
    .map(group => {
      const aliases = group.aliases.slice(0, 5);
      while (aliases.length < 5) aliases.push('');
      return [
        共通雑誌_優先略称を選ぶ_(group.codes, group.displayName),
        group.displayName,
        group.displayKana || '',
        group.keyType,
        group.normalJoin || '',
        group.specialJoin || '-',
        共通雑誌_言語一覧を連結_(group.languages),
        aliases[0], aliases[1], aliases[2], aliases[3], aliases[4],
        共通雑誌_備考結合_.apply(null, group.notes),
      ];
    })
    .sort((a, b) => String(a[1] || '').localeCompare(String(b[1] || ''), 'en'));
}

function 共通雑誌初期データ_() {
  const rows = [
    ['W', 'W KOREA', 'ダブリュー・コリア', '年月型', '', '-', 'KR', 'W', 'W Korea', '', '', '', ''],
    ['VOGU', 'VOGUE KOREA', 'ヴォーグ・コリア', '年月型', '', '-', 'KR', 'VOGUE', 'VOGUE Korea', 'VOGUE KOREA', '', '', ''],
    ['VOGUTW', 'VOGUE TAIWAN', 'ヴォーグ・台湾版', '年月型', '', '-', 'TW', 'VOGUE', 'VOGUE Taiwan', 'VOGUE TAIWAN', 'Vogue Taiwan', '', ''],
    ['ELLE', 'ELLE KOREA', 'エル・コリア', '年月型', '', '-', 'KR,CN', 'ELLE', 'ELLE Korea', 'ELLE KOREA', '', '', ''],
    ['ELLETW', 'ELLE TAIWAN', 'エル・台湾版', '年月型', '', '-', 'TW', 'ELLE', 'ELLE Taiwan', 'ELLE TAIWAN', 'Elle Taiwan', '', ''],
    ['ELLEHK', 'ELLE HONG KONG', 'エル・香港版', '年月型', '', '-', 'HK', 'ELLE Hong Kong', '', '', '', '', ''],
    ['BAZA', 'BAZAAR KOREA', 'バザー・コリア', '年月型', '', '-', 'KR', 'BAZAAR', "Harper's BAZAAR", "Harper's BAZAAR Korea", '', '', ''],
    ['BAZATW', 'BAZAAR TAIWAN', 'バザー・台湾版', '年月型', '', '-', 'TW', 'BAZAAR', "Harper's BAZAAR", "Harper's BAZAAR Taiwan", 'BAZAAR Taiwan', '', ''],
    ['GQ', 'GQ KOREA', 'ジーキュー・コリア', '年月型', '', '-', 'KR', 'GQ', 'GQ Korea', 'GQ KOREA', '', '', ''],
    ['GQTW', 'GQ TAIWAN', 'ジーキュー・台湾版', '年月型', '', '-', 'TW', 'GQ', 'GQ Taiwan', 'GQ TAIWAN', 'GQ台灣', '', ''],
    ['MARI', 'marie claire KOREA', 'マリ・クレール・コリア', '年月型', '', '-', 'KR', 'marie claire', 'Marie Claire', 'marie claire Korea', 'Marie Claire Korea', '', ''],
    ['MCLTW', 'marie claire TAIWAN', 'マリ・クレール・台湾', '年月型', '', '-', 'TW', 'marie claire', 'Marie Claire', 'Marie Claire美麗佳人', '美麗佳人', 'marie claire Taiwan', ''],
    ['COSM', 'COSMOPOLITAN KOREA', 'コスモポリタン', '年月型', '', '-', 'KR', 'COSMOPOLITAN', '', '', '', '', ''],
    ['MAXI', 'MAXIM KOREA', 'マキシム・コリア', '年月型', '', '-', 'KR', 'MAXIM', 'MAXIM KOREA', '', '', '', ''],
    ['ESQU', 'Esquire Korea', 'エスクァイア', '年月型', '', '-', 'KR,HK,CN', 'Esquire', 'Esquire Hong Kong', '時尚先生 Esquire', '', '', ''],
    ['ESQHK', 'Esquire Hong Kong', 'エスクァイア・香港版', '年月型', '', '-', 'HK', 'ESQUIRE HK', '', '', '', '', ''],
    ['DAZE', 'DAZED KOREA', 'デイズド・コリア', '年月型', '', '-', 'KR,TH', 'DAZED', 'DAZED&CONFUSED KOREA', '', '', '', ''],
    ['AKOR', 'allure KOREA', 'アルーア・コリア', '年月型', '', '-', 'KR', 'allure', 'allure Korea', '', '', '', ''],
    ['1STLOOK', '1st LOOK', 'ファーストルック', '号数型', '', '-', 'KR', '1ST LOOK', '1STLOOK', '', '', '', ''],
    ['CIN21', 'CINE21', 'シネ21', '号数型', '', '', 'KR', 'CINE 21', '', '', '', '', ''],
    ['SNGL', 'Singles', 'シングルズ', '年月型', '', '-', 'KR', 'SINGLES', '', '', '', '', ''],
    ['STAR', 'THE STAR', 'ザ・スター', '年月型', '', '-', 'KR', 'THE STAR', 'ザスター', '', '', '', ''],
    ['NYLN', 'NYLON KOREA', 'ナイロン・コリア', '年月型', '', '', 'KR', 'NYLON', 'NYLON Korea', '', '', '', ''],
    ['MENH', "Men's Health", 'メンズ・ヘルス', '年月型', '', '-', 'KR', "MEN'S HEALTH", 'Mens Health', '', '', '', ''],
    ['LOFI', "L'OFFICIEL KOREA", 'ロフィシェル', '年月型', '', '-', 'KR,HK,TH', "L'OFFICIEL", 'LOFFICIEL', 'L OFFICIEL', '', '', ''],
    ['AREN', 'ARENA HOMME+ KOREA', 'アリーナ・オム・プラス', '年月型', '', '-', 'KR,CN', 'ARENA HOMME+', 'ARENA', '', '', '', ''],
    ['PRAEW', 'Praew', 'プラーウ', '年月型', '', '', 'TH', '', '', '', '', '', ''],
    ['SAPDSP', 'Sudsapda', 'サッダパー', '年月型', '', '', 'TH', 'สุดสัปดาห์', '', '', '', '', ''],
    ['MUNO', "MEN'S UNO HK", 'メンズ・ウノ・香港', '年月型', '', '-', 'HK', "MEN'S UNO", 'S UNO HK', '', '', '', ''],
    ['WWD', 'WWD Korea', 'ダブリュー・ダブリュー・ディー', '年月型', '', '-', 'KR,TH', 'WWD', '', '', '', '', ''],
    ['MUSI', 'THE MUSICAL', 'ザ・ミュージカル', '年月型', '', '-', 'KR', 'THE MUSICAL', '', '', '', '', ''],
    ['SFOCUS', 'STAR FOCUS', 'スターフォーカス', '年月型', '', '-', 'KR', 'STAR FOCUS', '', '', '', '', ''],
    ['PERLA', 'Perla China', 'パーラ・チャイナ', '年月型', '', '-', 'CN', 'PERLA CHINA', '', '', '', '', ''],
    ['SENSECN', 'SENSE China', 'センス・チャイナ', '年月型', '', '-', 'CN', 'SENSE CHINA', '', '', '', '', ''],
    ['DELING', 'DE DELING', 'ドゥ・デリン', '年月型', '', '-', 'CN', '', '', '', '', '', ''],
    ['SPOTCN', 'SPOTLiGHT China', 'スポットライト・チャイナ', '年月型', '', '-', 'CN', 'SPOTLIGHT CHINA', '', '', '', '', ''],
    ['GRAZ', 'GRAZIA KOREA', 'グラーツィア', '年月型', '', '-', 'KR', 'GRAZIA', '', '', '', '', ''],
    ['HCUT', 'HIGH CUT', 'ハイカット', '号数型', '', '-', 'KR', 'High Cut', '', '', '', '', ''],
    ['LEGEND', 'LEGEND', 'レジェンド', '号数型', '', '-', 'HK', '', '', '', '', '', ''],
    ['MAPS', 'MAPS', 'マップス', '年月型', '', '-', 'KR', '', '', '', '', '', ''],
    ['MAXQ', 'MAXQ', 'マックスキュー', '年月型', '', '', 'KR', '', '', '', '', '', ''],
    ['NBLSS', 'Noblesse', 'ノブレス', '年月型', '', '', 'KR', 'NOBLESSE', '', '', '', '', ''],
    ['NBLSY', 'Noblesse Y MAGAZINE', 'ノブレス・ワイ', '号数型', '', '-', 'KR', 'Y MAGAZINE', 'Noblesse Y', '', '', '', ''],
    ['NBMEN', 'Noblesse MEN', 'ノブレス・メン', '号数型', '', '-', 'KR', 'MEN Noblesse', '', '', '', '', ''],
    ['TCLA', 'TOP Class', 'トップクラス', '年月型', '', '', 'KR', 'TOP CLASS', '', '', '', '', ''],
    ['THEA', 'THEATRE+', 'シアタープラス', '年月型', '', '-', 'KR,TH', 'THEATRE PLUS', '', '', '', '', ''],
    ['WSEN', 'Woman sense', 'ウーマンセンス', '年月型', '', '', 'KR', 'WOMAN SENSE', '', '', '', '', ''],
    ['YSDA', '女性東亜', 'ヨソンドンア', '年月型', '', '', 'KR', '', '', '', '', '', ''],
    ['YSJS', '女性朝鮮', 'ヨソンジョソン', '年月型', '', '', 'KR', '', '', '', '', '', ''],
    ['FORBES', 'Forbes', 'フォーブス・コリア', '年月型', '', '', 'KR', '', '', '', '', '', ''],
    ['RGST', 'Rolling Stone Korea', 'ローリングストーン・コリア', '号数型', '', '', 'KR', 'Rolling Stone', '', '', '', '', ''],
    ['SURE', 'SURE', 'シュア', '年月型', '', '', 'KR', '', '', '', '', '', ''],
    ['UBRK', 'URBANLIKE', 'アーバンライク', '年月型', '', '', 'KR', '', '', '', '', '', ''],
    ['VGIR', 'VOGUE girl', 'ヴォーグ・ガール', '年月型', '', '', 'KR', '', '', '', '', '', ''],
    ['WMUSIC', 'Weekly Music', 'ウィークリーミュージック', '年月型', '', '', 'KR', '', '', '', '', '', ''],
  ];
  return 共通雑誌_マスター行を国名なしへ統合_(rows);
}

/* ===== 完全版 override end ===== */