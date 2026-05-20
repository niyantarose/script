/**
 * 台湾グッズ.gs
 */

const 設定_台湾グッズ = {
  マスターシート名:   '台湾グッズ',
  作品シート名:       'Works（台湾グッズ）',
  作品略称マスター名: '作品略称マスター（台湾）',
  マスターファイルID: '1ZAy5wtQCq1ixl47MMEIGnrVWeae-F_NfN34lSOUsN5M',

  作品ヘッダー: ['WorksKey', '作品ID', '作品名（原題）', '作品名（日本語）', '作品略称', '登録数', '更新日時'],
  作品列数: 7,

  列名: {
    登録状況:         '登録状況',
    発番発行:         '発番発行',
    言語:             '言語',
    親コード:         '商品コード（SKU）',
    商品名出品用:     '商品名（出品用）',
    売価:             '売価',
    配送パターン:     '配送パターン',
    登録者:           '登録者',
    原価:             '原価',
    商品説明:         '商品説明',
    商品名日本語:     '商品名（日本語）',
    作品名原題:       '原題タイトル',
    作品名日本語:     '日本語タイトル',
    特典メモ:         '特典メモ',
    商品名原題:       '原題商品タイトル',
    サイト商品コード: 'サイト商品コード',
    メイン画像URL:    'メイン画像',
    追加画像URL:      '追加画像',
    発売日:           '発売日',
    重複チェックキー: '重複チェックキー',
    登録日:           '登録日',
    粗利益率:         '粗利益率'
  }
};

const 台湾グッズ_ヘッダー色 = {
  '登録状況':           '#4a86e8',
  '発番発行':           '#cc0000',
  '言語':               '#6aa84f',
  '商品コード（SKU）':  '#999999',
  '商品名（出品用）':   '#999999',
  '売価':               '#4a86e8',
  '配送パターン':       '#4a86e8',
  '登録者':             '#4a86e8',
  '原価':               '#e69138',
  '商品説明':           '#e69138',
  '商品名（日本語）':   '#f1c232',
  '原題タイトル':       '#6aa84f',
  '日本語タイトル':     '#999999',
  '特典メモ':           '#f1c232',
  '原題商品タイトル':   '#e69138',
  'サイト商品コード':   '#e69138',
  'メイン画像':         '#e69138',
  '追加画像':           '#e69138',
  '発売日':             '#e69138',
  '重複チェックキー':   '#999999',
  '登録日':             '#999999',
  '粗利益率':           '#999999',
};

/* ============================================================
 * 基本取得
 * ============================================================ */

function 台湾グッズ_マスターSSを取得_() {
  return SpreadsheetApp.openById(設定_台湾グッズ.マスターファイルID);
}

function 台湾グッズ_列マップを取得_(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(v => String(v).trim());
  const map = {};
  headers.forEach((h, i) => { if (h) map[h] = i + 1; });
  return map;
}

function 台湾グッズ_言語コードへ変換_(言語) {
  let raw = String(言語 || '').trim();
  if (!raw) return '';

  const fallback = {
    '台湾': 'TW', '臺灣': 'TW', 'TW': 'TW',
    '中国': 'CN', '中國': 'CN', 'CN': 'CN',
    '香港': 'HK', 'HK': 'HK',
    'タイ': 'TH', '泰国': 'TH', '泰國': 'TH', 'TH': 'TH',
    '韓国': 'KR', '韓國': 'KR', 'KR': 'KR',
    '日本': 'JP', 'JP': 'JP'
  };

  // fallbackに直接ヒットすればアクセス不要
  if (fallback[raw]) return fallback[raw];

  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = 'LANG_' + raw;
    const cached = cache.get(cacheKey);
    if (cached) return fallback[cached] || cached;

    const masterSS = 台湾グッズ_マスターSSを取得_();
    const 言語シート = masterSS.getSheetByName('言語マスター');
    if (言語シート && 言語シート.getLastRow() >= 2) {
      const 言語データ = 言語シート.getRange(2, 1, 言語シート.getLastRow() - 1, 2).getValues();
      for (const [名前, コード] of 言語データ) {
        if (String(名前).trim() === raw) {
          const code = String(コード).trim();
          cache.put(cacheKey, code, 3600);
          raw = code;
          break;
        }
      }
    }
  } catch (_) {}

  return fallback[raw] || raw.toUpperCase();
}

function 台湾グッズ_言語表示へ変換_(言語) {
  const code = 台湾グッズ_言語コードへ変換_(言語);
  const map = {
    TW: '台湾',
    CN: '中国',
    HK: '香港',
    TH: 'タイ',
    KR: '韓国',
    JP: '日本'
  };
  return map[code] || String(言語 || '').trim();
}

function 台湾グッズ_LOG_(label, obj) {
  try {
    if (obj === undefined) {
      console.log('[台湾グッズ] ' + label);
    } else {
      console.log('[台湾グッズ] ' + label + ' ' + JSON.stringify(obj));
    }
  } catch (err) {
    console.log('[台湾グッズ] ' + label + ' (log stringify error: ' + err + ')');
  }
}

/* ============================================================
 * 作品情報取得
 * ============================================================ */

function 台湾グッズ_作品マスターから取得_(作品名原題) {
  const ss = SpreadsheetApp.getActive();
  const マスター = ss.getSheetByName(設定_台湾グッズ.作品略称マスター名);
  if (!マスター || マスター.getLastRow() < 2) return null;

  const データ = マスター.getRange(2, 1, マスター.getLastRow() - 1, 4).getValues();
  for (const row of データ) {
    if (String(row[0]).trim() === String(作品名原題).trim()) {
      return {
        作品名日本語: String(row[1]).trim(),
        作品略称:     String(row[2]).trim()
      };
    }
  }
  return null;
}

function 台湾グッズ_Worksから取得_(作品名原題) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(設定_台湾グッズ.作品シート名);
  if (!sh || sh.getLastRow() < 2) return null;

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 設定_台湾グッズ.作品列数).getValues();
  for (const row of data) {
    if (String(row[2] || '').trim() === String(作品名原題).trim()) {
      return {
        作品ID:       String(row[1] || '').trim(),
        作品名原題:   String(row[2] || '').trim(),
        作品名日本語: String(row[3] || '').trim(),
        作品略称:     String(row[4] || '').trim()
      };
    }
  }
  return null;
}

function 台湾グッズ_作品情報を取得_(作品名原題) {
  const fromWorks = 台湾グッズ_Worksから取得_(作品名原題) || {};
  const fromMaster = 台湾グッズ_作品マスターから取得_(作品名原題) || {};

  return {
    作品ID:       fromWorks.作品ID || '',
    作品名原題:   String(作品名原題 || '').trim(),
    作品名日本語: fromWorks.作品名日本語 || fromMaster.作品名日本語 || '',
    作品略称:     fromWorks.作品略称 || fromMaster.作品略称 || ''
  };
}

/* ============================================================
 * 生成系
 * ============================================================ */

function 台湾グッズ_親コードを生成_(言語, 作品名原題, 現在行) {
  console.log('--- 親コード生成 start ---');
  console.log(JSON.stringify({ 言語, 作品名原題, 現在行 }));

  const 作品情報 = 台湾グッズ_作品情報を取得_(作品名原題);
  console.log('作品情報=' + JSON.stringify(作品情報));

  if (!作品情報 || !作品情報.作品略称) return null;

  const langCode = 台湾グッズ_言語コードへ変換_(言語) || 'TW';
  const prefix = `${langCode}-${作品情報.作品略称}-`;
  console.log('generated prefix=' + prefix);

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('台湾グッズ');
  if (!sh || sh.getLastRow() < 2) return `${prefix}0001`;

  const 列 = 台湾グッズ_列マップを取得_(sh);
  const SKU列 = 列['商品コード（SKU）'];
  if (!SKU列) return `${prefix}0001`;

  const 既存コード = sh.getRange(2, SKU列, sh.getLastRow() - 1, 1).getValues().flat();
  let 最大番号 = 0;

  既存コード.forEach((code, i) => {
    if (i + 2 === 現在行) return;
    if (code && String(code).startsWith(prefix)) {
      const n = parseInt(String(code).replace(prefix, ''), 10);
      if (!isNaN(n) && n > 最大番号) 最大番号 = n;
    }
  });

  console.log('最大番号=' + 最大番号);
  return `${prefix}${String(最大番号 + 1).padStart(4, '0')}`;
}

function 台湾グッズ_出品用商品名を生成_(言語, 作品名日本語, 商品名日本語, 商品名原題, 特典メモ) {
  const 作品名表示 = String(作品名日本語 || '').trim();
  const 商品名表示 = String(商品名日本語 || 商品名原題 || '').trim();

  if (!作品名表示 || !商品名表示) return '';

  const 言語表示 = 台湾グッズ_言語表示へ変換_(言語) || '台湾';
  let 名前 = `${言語表示} グッズ 『${作品名表示} ${商品名表示}』`;

  if (特典メモ && String(特典メモ).trim()) {
    名前 += ` ${String(特典メモ).trim()}`;
  }
  if (商品名原題 && String(商品名原題).trim() && String(商品名原題).trim() !== 商品名表示) {
    名前 += ` ${String(商品名原題).trim()}`;
  }

  return 名前.replace(/\s+/g, ' ').trim();
}

function 台湾グッズ_重複チェックキーを生成_(サイト商品コード, 商品名原題, 作品名原題, 商品名日本語) {
  const site = String(サイト商品コード || '').trim();
  const 原題商品 = String(商品名原題 || '').trim();
  const 原題作品 = String(作品名原題 || '').trim();
  const 商品名 = String(商品名日本語 || '').trim();

  if (site) return `SITE||${site}`;
  if (原題商品) return `ORGITEM||${原題商品}`;
  if (原題作品 || 商品名) return [原題作品, 商品名].filter(Boolean).join('||');
  return '';
}

function 台湾グッズ_粗利益率を計算_(売価, 原価) {
  const 売 = parseFloat(売価 || 0);
  const 原 = parseFloat(原価 || 0);
  if (!(売 > 0 && 原 > 0)) return null;

  let レート = 0;
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('RATE_TWD');
    if (cached) {
      レート = parseFloat(cached);
    } else {
      // IMPORTRANGEではなく共通マスターから直接取得
      const masterSS = SpreadsheetApp.openById('1ZAy5wtQCq1ixl47MMEIGnrVWeae-F_NfN34lSOUsN5M');
      const sh = masterSS.getSheetByName('為替レートマスター');
      if (sh) {
        レート = parseFloat(sh.getRange('B2').getValue() || 0); // B2=TWD
        if (レート > 0) cache.put('RATE_TWD', String(レート), 3600);
      }
    }
  } catch (_) {}

  if (!(レート > 0)) return null;
  return Math.round(((売 - 原 * レート * 1.3 - 売 * 0.22) / 売) * 1000) / 1000;
}

/* ============================================================
 * 1行補完
 * ============================================================ */

function 台湾グッズ_1行補完_(sh, row) {
  if (!sh || sh.getName() !== '台湾グッズ' || row < 2) {
    台湾グッズ_LOG_('1行補完 skip', {
      hasSheet: !!sh,
      sheet: sh ? sh.getName() : null,
      row
    });
    return false;
  }

  const 列 = 台湾グッズ_列マップを取得_(sh);
  const lastCol = sh.getLastColumn();
  const rowValues = sh.getRange(row, 1, 1, lastCol).getValues()[0];

  const g = (名前) => 列[名前] ? rowValues[列[名前] - 1] : '';
  const s = (名前, 値) => {
    if (!列[名前]) {
      台湾グッズ_LOG_('set skip: no column', { row, colName: 名前, value: 値 });
      return false;
    }

    const idx = 列[名前] - 1;
    const 現在値 = rowValues[idx];

    if (現在値 instanceof Date && 値 instanceof Date) {
      if (現在値.getTime() === 値.getTime()) return false;
    } else if (String(現在値) === String(値)) {
      return false;
    }

    台湾グッズ_LOG_('set value', {
      row,
      colName: 名前,
      before: 現在値 instanceof Date ? 現在値.toISOString() : String(現在値),
      after: 値 instanceof Date ? 値.toISOString() : String(値)
    });

    rowValues[idx] = 値;
    return true;
  };

  let changed = false;

  let 値 = {
    言語: String(g('言語') || '').trim(),
    作品名原題: String(g('原題タイトル') || '').trim(),
    作品名日本語: String(g('日本語タイトル') || '').trim(),
    商品名日本語: String(g('商品名（日本語）') || '').trim(),
    商品名原題: String(g('原題商品タイトル') || '').trim(),
    特典メモ: String(g('特典メモ') || '').trim(),
    サイト商品コード: String(g('サイト商品コード') || '').trim(),
    商品コード: String(g('商品コード（SKU）') || '').trim(),
    売価: g('売価'),
    原価: g('原価'),
    発番発行: g('発番発行') === true
  };

  台湾グッズ_LOG_('1行補完 input', {
    row,
    値
  });

  if (!値.作品名原題 && 値.商品名原題) {
    changed = s('原題タイトル', 値.商品名原題) || changed;
    値.作品名原題 = 値.商品名原題;
    台湾グッズ_LOG_('原題タイトルを原題商品タイトルから補完', {
      row,
      作品名原題: 値.作品名原題
    });
  }

  let 作品情報 = null;
  if (値.作品名原題) {
    作品情報 = 台湾グッズ_作品情報を取得_(値.作品名原題);
    台湾グッズ_LOG_('作品情報取得', {
      row,
      作品名原題: 値.作品名原題,
      作品情報
    });

    if (!値.作品名日本語 && 作品情報 && 作品情報.作品名日本語) {
      changed = s('日本語タイトル', 作品情報.作品名日本語) || changed;
      値.作品名日本語 = 作品情報.作品名日本語;
    }

    if (作品情報 && 作品情報.作品ID && !String(g('作品ID(W)(自動)') || '').trim()) {
      changed = s('作品ID(W)(自動)', 作品情報.作品ID) || changed;
    }
  } else {
    台湾グッズ_LOG_('作品情報取得 skip: 作品名原題なし', { row });
  }

  const 出品用商品名 = 台湾グッズ_出品用商品名を生成_(
    値.言語,
    値.作品名日本語 || 値.作品名原題,
    値.商品名日本語 || 値.商品名原題,
    値.商品名原題,
    値.特典メモ
  );

  台湾グッズ_LOG_('出品用商品名生成結果', {
    row,
    出品用商品名
  });

  if (出品用商品名) {
    changed = s('商品名（出品用）', 出品用商品名) || changed;
  }

  const 重複キー = 台湾グッズ_重複チェックキーを生成_(
    値.サイト商品コード,
    値.商品名原題,
    値.作品名原題,
    値.商品名日本語
  );

  台湾グッズ_LOG_('重複キー生成結果', {
    row,
    重複キー
  });

  if (重複キー) {
    changed = s('重複チェックキー', 重複キー) || changed;
  }

  const 粗利益率 = 台湾グッズ_粗利益率を計算_(値.売価, 値.原価);
  台湾グッズ_LOG_('粗利益率計算結果', {
    row,
    売価: 値.売価,
    原価: 値.原価,
    粗利益率
  });

  if (粗利益率 != null) {
    changed = s('粗利益率', 粗利益率) || changed;
  }

  let 親コード = 値.商品コード;

  const 期待prefix = `${台湾グッズ_言語コードへ変換_(値.言語)}-`;
const SKU不一致 = 親コード && !String(親コード).startsWith(期待prefix);

if ((!親コード || SKU不一致) && 値.言語 && 値.作品名原題) {
    const code = 台湾グッズ_親コードを生成_(値.言語, 値.作品名原題, row);

    台湾グッズ_LOG_('親コード生成結果', {
      row,
      言語: 値.言語,
      作品名原題: 値.作品名原題,
      code
    });

    if (code) {
      changed = s('商品コード（SKU）', code) || changed;
      親コード = code;
    } else if (値.作品名原題) {
      changed = s('商品コード（SKU）', 'ERROR:マスター未登録') || changed;
      親コード = 'ERROR:マスター未登録';
    }
  } else {
    台湾グッズ_LOG_('親コード生成 skip', {
      row,
      発番発行: 値.発番発行,
      既存商品コード: 親コード
    });
  }

  const 不足 = 台湾グッズ_不足項目を返す_(値);

  台湾グッズ_LOG_('不足項目判定', {
    row,
    不足,
    親コード
  });

  if (親コード && !String(親コード).startsWith('ERROR')) {
    changed = s('商品コードステータス', '生成済み') || changed;
  } else if (String(親コード).startsWith('ERROR')) {
    changed = s('商品コードステータス', 'マスター未登録') || changed;
  } else if (不足.length > 0) {
    changed = s('商品コードステータス', '情報不足: ' + 不足.join(',')) || changed;
  } else if (値.発番発行) {
    changed = s('商品コードステータス', '発番待ち') || changed;
  } else {
    changed = s('商品コードステータス', '入力途中') || changed;
  }

  const 何か入っている = [
    値.言語,
    値.作品名原題,
    値.作品名日本語,
    値.商品名日本語,
    値.商品名原題
  ].some(v => String(v || '').trim());

  if (何か入っている && !g('登録日')) {
    changed = s('登録日', new Date()) || changed;
  }

  if (changed) {
    sh.getRange(row, 1, 1, lastCol).setValues([rowValues]);
    if (列['粗利益率']) {
      sh.getRange(row, 列['粗利益率']).setNumberFormat('0.0%');
    }
    台湾グッズ_LOG_('1行補完 write done', { row });
  } else {
    台湾グッズ_LOG_('1行補完 no change', { row });
  }

  return changed;
}
/* ============================================================
 * シート作成
 * ============================================================ */

function 台湾グッズ作品マスターを作成() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const シート名 = 設定_台湾グッズ.作品略称マスター名;

  if (ss.getSheetByName(シート名)) {
    ui.alert(`「${シート名}」シートは既に存在します`);
    return;
  }

  const sh = ss.insertSheet(シート名);
  const ヘッダー = ['作品名（原題）', '作品名（日本語）', '作品略称', '備考'];
  sh.getRange(1, 1, 1, ヘッダー.length).setValues([ヘッダー]);
  sh.getRange(1, 1, 1, ヘッダー.length)
    .setBackground('#4a86e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  [200, 200, 100, 160].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  ui.alert('✅ 作品略称マスター（台湾）を作成しました');
}

function 台湾グッズWorksシートを作成() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const シート名 = 設定_台湾グッズ.作品シート名;

  if (ss.getSheetByName(シート名)) {
    ui.alert(`「${シート名}」シートは既に存在します`);
    return;
  }

  const sh = ss.insertSheet(シート名);
  sh.getRange(1, 1, 1, 設定_台湾グッズ.作品列数).setValues([設定_台湾グッズ.作品ヘッダー]);
  sh.getRange(1, 1, 1, 設定_台湾グッズ.作品列数)
    .setBackground('#cc0000')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  [150, 120, 200, 200, 100, 80, 150].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  ui.alert('✅ Works（台湾グッズ）シートを作成しました');
}

function 台湾グッズシートを作成() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const シート名 = '台湾グッズ';

  if (ss.getSheetByName(シート名)) {
    ui.alert(`「${シート名}」シートは既に存在します`);
    return;
  }

  const sh = ss.insertSheet(シート名);
  const ヘッダー = [
    '発番発行', '登録状況', '言語', '商品コード（SKU）', '商品名（出品用）',
    '売価', '原価', '粗利益率', '配送パターン',
    '商品名（日本語）', '日本語タイトル', '原題タイトル', '原題商品タイトル',
    '特典メモ', '商品説明', 'リンク', '発売日',
    'メイン画像', '追加画像', 'サイト商品コード',
    '作品ID(W)(自動)', '商品コードステータス', '重複チェックキー', '登録日', '登録者'
  ];

  sh.getRange(1, 1, 1, ヘッダー.length).setValues([ヘッダー]);

  ヘッダー.forEach((h, i) => {
    const 色 = 台湾グッズ_ヘッダー色[h] || '#4a86e8';
    sh.getRange(1, i + 1)
      .setBackground(色)
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  });

  sh.setRowHeight(1, 40);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(5);

  const 最終行 = 1000;
  sh.getRange(2, 1, 最終行 - 1, 1).insertCheckboxes();
  sh.getRange(2, 2, 最終行 - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['未登録', '登録済み', '売り切れ'], true)
      .build()
  );

  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$B2="未登録"')
      .setBackground('#ffff00')
      .setRanges([sh.getRange(2, 1, 最終行 - 1, ヘッダー.length)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$B2="登録済み"')
      .setBackground('#d9ead3')
      .setRanges([sh.getRange(2, 1, 最終行 - 1, ヘッダー.length)])
      .build(),
  ]);

  ui.alert('✅ 台湾グッズシートを作成しました');
}

/* ============================================================
 * onEdit / 取込後補完
 * ============================================================ */

function 台湾グッズ_onEdit(e) {
  const sh = e.range.getSheet();
  const 開始行 = e.range.getRow();
  const 行数 = e.range.getNumRows();
  const col = e.range.getColumn();

  台湾グッズ_LOG_('onEdit start', {
    sheet: sh.getName(),
    開始行,
    行数,
    col,
    value: e && e.value !== undefined ? e.value : null,
    oldValue: e && e.oldValue !== undefined ? e.oldValue : null
  });

  if (sh.getName() !== '台湾グッズ') {
    台湾グッズ_LOG_('return: sheet mismatch', { sheet: sh.getName() });
    return;
  }

  if (開始行 < 2) {
    台湾グッズ_LOG_('return: row < 2', { 開始行 });
    return;
  }

  const 列 = 台湾グッズ_列マップを取得_(sh);
  const 監視列 = [
    列['発番発行'],
    列['言語'],
    列['商品名（日本語）'],
    列['日本語タイトル'],
    列['原題タイトル'],
    列['原題商品タイトル'],
    列['特典メモ'],
    列['サイト商品コード'],
    列['売価'],
    列['原価']
  ].filter(Boolean);

  const 編集終了列 = e.range.getLastColumn();
  const 対象列あり = 監視列.some(c => c >= col && c <= 編集終了列);

  台湾グッズ_LOG_('watch columns', { col, 編集終了列, 監視列, 対象列あり });

  if (!対象列あり) {
    台湾グッズ_LOG_('skip: no target column in range');
    return;
  }

  const lock = LockService.getDocumentLock();
  try {
    if (!lock.tryLock(10000)) {
      台湾グッズ_LOG_('return: lock failed');
      return;
    }
  } catch (err) {
    台湾グッズ_LOG_('return: lock error', { error: String(err) });
    return;
  }

  try {
    for (let row = 開始行; row < 開始行 + 行数; row++) {
      if (row < 2) continue;
      台湾グッズ_LOG_('go: 1行補完', { row });
      const changed = 台湾グッズ_1行補完_(sh, row);
      台湾グッズ_LOG_('1行補完 end', { row, changed });
    }
  } catch (err) {
    台湾グッズ_LOG_('onEdit error', {
      col,
      error: String(err),
      stack: err && err.stack ? String(err.stack) : ''
    });
    throw err;
  } finally {
    lock.releaseLock();
    台湾グッズ_LOG_('lock released');
  }
}
/* ============================================================
 * 確定発行
 * ============================================================ */

function 台湾グッズ_確定発行() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('台湾グッズ');
  if (!sh || sh.getLastRow() < 2) {
    ui.alert('データがありません');
    return;
  }

  const 列 = 台湾グッズ_列マップを取得_(sh);
  if (!列['発番発行']) {
    ui.alert('発番発行列が見つかりません');
    return;
  }

  const 全データ = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const get行 = (row, 名前) => 列[名前] ? row[列[名前] - 1] : '';

  const 対象行リスト = [];
  全データ.forEach((row, i) => {
    if (get行(row, '発番発行') === true) {
      対象行リスト.push({ row, rowNum: i + 2 });
    }
  });

  if (対象行リスト.length === 0) {
    ui.alert('発番発行列にチェックが入っている行がありません');
    return;
  }

  const 既存サイトコードMap = {};
  const 既存商品名Map = {};

  全データ.forEach((row, i) => {
    const rowNum = i + 2;
    const サイトコード = String(get行(row, 'サイト商品コード') || '').trim();
    const 商品名原題   = String(get行(row, '原題商品タイトル') || '').trim();

    if (サイトコード) {
      if (!既存サイトコードMap[サイトコード]) 既存サイトコードMap[サイトコード] = [];
      既存サイトコードMap[サイトコード].push(rowNum);
    }
    if (商品名原題) {
      if (!既存商品名Map[商品名原題]) 既存商品名Map[商品名原題] = [];
      既存商品名Map[商品名原題].push(rowNum);
    }
  });

  const ブロックリスト = [];
  対象行リスト.forEach(({ row, rowNum }) => {
    const サイトコード = String(get行(row, 'サイト商品コード') || '').trim();
    const 商品名原題   = String(get行(row, '原題商品タイトル') || '').trim();

    if (サイトコード && 既存サイトコードMap[サイトコード]) {
      const 重複行 = 既存サイトコードMap[サイトコード].filter(r => r !== rowNum);
      if (重複行.length > 0) {
        ブロックリスト.push(`${rowNum}行目：サイト商品コード「${サイトコード}」が${重複行[0]}行目と重複`);
        return;
      }
    }

    if (商品名原題 && 既存商品名Map[商品名原題]) {
      const 重複行 = 既存商品名Map[商品名原題].filter(r => r !== rowNum);
      if (重複行.length > 0) {
        ブロックリスト.push(`${rowNum}行目：原題商品タイトル「${商品名原題}」が${重複行[0]}行目と重複`);
      }
    }
  });

  if (ブロックリスト.length > 0) {
    ui.alert(`⚠️ 重複が見つかりました。\n\n${ブロックリスト.join('\n')}\n\n重複を解消してから再度実行してください。`);
    return;
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(60000);
  try {
    const Worksシート = 台湾グッズ_Worksシートを確保_(ss);
    const Works更新Map = {};
    let 発行数 = 0;

    対象行リスト.forEach(({ row, rowNum }) => {
      const 親コード = String(get行(row, '商品コード（SKU）') || '').trim();

      if (!親コード || 親コード.startsWith('ERROR')) {
        ui.alert(`${rowNum}行目：商品コード（SKU）が未生成のためスキップしました`);
        return;
      }

      if (列['登録状況']) sh.getRange(rowNum, 列['登録状況']).setValue('登録済み');
      if (列['発番発行']) sh.getRange(rowNum, 列['発番発行']).setValue(false);

      発行数++;

      const 作品名 = String(get行(row, '原題タイトル') || '').trim();
      if (作品名) {
        if (!Works更新Map[作品名]) {
          Works更新Map[作品名] = {
            日本語: String(get行(row, '日本語タイトル') || '').trim(),
            件数: 0
          };
        }
        Works更新Map[作品名].件数++;
      }
    });

    台湾グッズ_Worksを更新_(Worksシート, Works更新Map);
    ui.alert(`✅ 確定発行完了: ${発行数}件`);
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
 * Works更新
 * ============================================================ */

function 台湾グッズ_Worksシートを確保_(ss) {
  let sh = ss.getSheetByName(設定_台湾グッズ.作品シート名);
  if (!sh) {
    sh = ss.insertSheet(設定_台湾グッズ.作品シート名);
    sh.getRange(1, 1, 1, 設定_台湾グッズ.作品列数).setValues([設定_台湾グッズ.作品ヘッダー]);
    sh.getRange(1, 1, 1, 設定_台湾グッズ.作品列数)
      .setBackground('#cc0000')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    sh.setFrozenRows(1);
  }
  return sh;
}

function 台湾グッズ_Worksを更新_(Worksシート, Works更新Map) {
  const 既存Map = {};

  if (Worksシート.getLastRow() >= 2) {
    Worksシート
      .getRange(2, 1, Worksシート.getLastRow() - 1, 設定_台湾グッズ.作品列数)
      .getValues()
      .forEach((row, i) => {
        if (row[2]) 既存Map[String(row[2]).trim()] = i + 2;
      });
  }

  Object.keys(Works更新Map).forEach(作品名 => {
    const info = Works更新Map[作品名];
    const WorksKey = 作品名.replace(/\s+/g, '').toLowerCase();

    if (既存Map[作品名]) {
      const 行 = 既存Map[作品名];
      Worksシート.getRange(行, 6).setValue(info.件数);
      Worksシート.getRange(行, 7).setValue(new Date());
    } else {
      const 新ID = 'TW-W-' + String(Worksシート.getLastRow()).padStart(4, '0');
      const マスター情報 = 台湾グッズ_作品マスターから取得_(作品名);

      Worksシート.appendRow([
        WorksKey,
        新ID,
        作品名,
        info.日本語,
        マスター情報 ? マスター情報.作品略称 : '',
        info.件数,
        new Date()
      ]);

      既存Map[作品名] = Worksシート.getLastRow();
    }
  });
}

/* ============================================================
 * プルダウン更新
 * ============================================================ */

function 台湾グッズ_プルダウン更新() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('台湾グッズ');
  if (!sh) {
    SpreadsheetApp.getUi().alert('台湾グッズシートが見つかりません');
    return;
  }

  const 最終行 = Math.max(sh.getLastRow() + 100, 200);
  const 列 = 台湾グッズ_列マップを取得_(sh);

  let masterSS = null;
  try {
    masterSS = 台湾グッズ_マスターSSを取得_();
  } catch (_) {}

  if (masterSS && 列['言語']) {
    const 言語シート = masterSS.getSheetByName('言語マスター');
    if (言語シート && 言語シート.getLastRow() >= 2) {
      const 言語値 = 言語シート
        .getRange(2, 1, 言語シート.getLastRow() - 1, 1)
        .getValues()
        .flat()
        .map(v => String(v).trim())
        .filter(v => v);

      if (言語値.length > 0) {
        sh.getRange(2, 列['言語'], 最終行 - 1, 1).setDataValidation(
          SpreadsheetApp.newDataValidation().requireValueInList(言語値, true).build()
        );
      }
    }
  }

  if (列['登録状況']) {
    sh.getRange(2, 列['登録状況'], 最終行 - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['未登録', '登録済み', '売り切れ'], true).build()
    );
  }

  if (列['発番発行']) {
    sh.getRange(2, 列['発番発行'], 最終行 - 1, 1).insertCheckboxes();
  }

  if (masterSS && 列['配送パターン']) {
    const 配送マスター = masterSS.getSheetByName('配送パターンマスター');
    if (配送マスター && 配送マスター.getLastRow() >= 2) {
      const 配送値 = 配送マスター
        .getRange(2, 1, 配送マスター.getLastRow() - 1, 1)
        .getValues()
        .flat()
        .map(v => String(v).trim())
        .filter(v => v);

      if (配送値.length > 0) {
        sh.getRange(2, 列['配送パターン'], 最終行 - 1, 1).setDataValidation(
          SpreadsheetApp.newDataValidation().requireValueInList(配送値, true).setAllowInvalid(false).build()
        );
      }
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('✅ 台湾グッズ プルダウン更新完了', '完了', 3);
}

/* ============================================================
 * 一括更新
 * ============================================================ */

function 台湾グッズ_一括更新() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert(
    '確認',
    '台湾グッズを一括更新します。\n\n・通常項目は全行再計算\n・Works / SKU の採番は B列=未登録 の行だけ\n\n続行しますか？',
    ui.ButtonSet.OK_CANCEL
  ) !== ui.Button.OK) return;

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('台湾グッズ');
  if (!sh || sh.getLastRow() < 2) {
    ui.alert('台湾グッズシートにデータがありません');
    return;
  }

  const Worksシート = 台湾グッズ_Worksシートを確保_(ss);
  const 列 = 台湾グッズ_列マップを取得_(sh);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const get = (row, 名前) => 列[名前] ? row[列[名前] - 1] : '';
  const set = (row, 名前, 値) => {
    if (!列[名前]) return false;
    const idx = 列[名前] - 1;

    const cur = row[idx];
    if (cur instanceof Date && 値 instanceof Date) {
      if (cur.getTime() === 値.getTime()) return false;
    } else if (String(cur || '') === String(値 || '')) {
      return false;
    }

    row[idx] = 値;
    return true;
  };

  const now = new Date();
  let 更新数 = 0;

  /* ============================================================
   * 1) Works既存情報を取得
   * ============================================================ */
  const worksColCount = 設定_台湾グッズ.作品列数;
  let worksData = [];
  if (Worksシート.getLastRow() >= 2) {
    worksData = Worksシート.getRange(2, 1, Worksシート.getLastRow() - 1, worksColCount).getValues();
  }

  const worksMap = {}; // 原題 -> { idx, 作品ID, 日本語, 略称 }
  let maxWorksIdNum = 0;

  worksData.forEach((row, idx) => {
    const 原題 = String(row[2] || '').trim();
    if (原題) {
      worksMap[原題] = {
        idx,
        作品ID: String(row[1] || '').trim(),
        日本語: String(row[3] || '').trim(),
        略称:   String(row[4] || '').trim()
      };
    }

    const m = String(row[1] || '').match(/^TW-W-(\d+)$/);
    if (m) {
      const n = parseInt(m[1], 10);
      if (!isNaN(n) && n > maxWorksIdNum) maxWorksIdNum = n;
    }
  });

  /* ============================================================
   * 2) SKU使用済み番号を収集
   *    ※ 未登録行は振り直し対象なので、使用済みから除外
   * ============================================================ */
  const usedByPrefix = {}; // { 'TW-ZCR-': Set([1,2,3]) }

  function prefixと番号を分解(code) {
    const m = String(code || '').trim().match(/^(.*-)(\d{4})$/);
    if (!m) return null;
    return {
      prefix: m[1],
      num: parseInt(m[2], 10)
    };
  }

  values.forEach(row => {
    const 登録状況 = String(get(row, '登録状況') || '').trim();
    const code = String(get(row, '商品コード（SKU）') || '').trim();

    // 未登録行は今回振り直す対象なので、使用済み番号には入れない
    if (登録状況 === '未登録') return;

    const parsed = prefixと番号を分解(code);
    if (!parsed) return;

    if (!usedByPrefix[parsed.prefix]) usedByPrefix[parsed.prefix] = new Set();
    usedByPrefix[parsed.prefix].add(parsed.num);
  });

  /* ============================================================
   * 3) 全行の通常項目を再計算
   *    採番(Works/SKU)だけは 未登録 行に限定
   * ============================================================ */
  const 採番対象行 = []; // {rowIndex, rowNum, prefix, 原題}
  const 全件数Map = {}; // 原題 -> 件数（全行ベース）

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowNum = i + 2;

    const 登録状況 = String(get(row, '登録状況') || '').trim();
    const 言語 = String(get(row, '言語') || '').trim();
    let 作品名原題 = String(get(row, '原題タイトル') || '').trim();
    let 作品名日本語 = String(get(row, '日本語タイトル') || '').trim();
    const 商品名日本語 = String(get(row, '商品名（日本語）') || '').trim();
    const 商品名原題 = String(get(row, '原題商品タイトル') || '').trim();
    const 特典メモ = String(get(row, '特典メモ') || '').trim();
    const サイト商品コード = String(get(row, 'サイト商品コード') || '').trim();
    const 売価 = get(row, '売価');
    const 原価 = get(row, '原価');

    // 原題タイトルが空なら原題商品タイトルで補完
    if (!作品名原題 && 商品名原題) {
      if (set(row, '原題タイトル', 商品名原題)) 更新数++;
      作品名原題 = 商品名原題;
    }

    // 作品情報補完
    let 作品情報 = null;
    if (作品名原題) {
      作品情報 = 台湾グッズ_作品情報を取得_(作品名原題);

      if (!作品名日本語 && 作品情報 && 作品情報.作品名日本語) {
        if (set(row, '日本語タイトル', 作品情報.作品名日本語)) 更新数++;
        作品名日本語 = 作品情報.作品名日本語;
      }
    }

    // 出品用商品名
    const 出品用商品名 = 台湾グッズ_出品用商品名を生成_(
      言語,
      作品名日本語 || 作品名原題,
      商品名日本語 || 商品名原題,
      商品名原題,
      特典メモ
    );
    if (出品用商品名) {
      if (set(row, '商品名（出品用）', 出品用商品名)) 更新数++;
    }

    // 重複チェックキー
    const 重複キー = 台湾グッズ_重複チェックキーを生成_(
      サイト商品コード,
      商品名原題,
      作品名原題,
      商品名日本語
    );
    if (重複キー) {
      if (set(row, '重複チェックキー', 重複キー)) 更新数++;
    }

    // 粗利益率
    const 粗利益率 = 台湾グッズ_粗利益率を計算_(売価, 原価);
    if (粗利益率 != null) {
      if (set(row, '粗利益率', 粗利益率)) 更新数++;
    }

    // 登録日
    const 何か入っている = [言語, 作品名原題, 作品名日本語, 商品名日本語, 商品名原題]
      .some(v => String(v || '').trim());
    if (何か入っている && !get(row, '登録日')) {
      if (set(row, '登録日', now)) 更新数++;
    }

    // 全件数Map（全行対象）
    if (作品名原題) {
      全件数Map[作品名原題] = (全件数Map[作品名原題] || 0) + 1;
    }

    // 採番対象は未登録のみ
    if (登録状況 === '未登録' && 作品名原題) {
      // Worksが無ければ作る
      if (!worksMap[作品名原題]) {
        maxWorksIdNum++;
        const 新作品ID = 'TW-W-' + String(maxWorksIdNum).padStart(4, '0');
        const worksKey = 作品名原題.replace(/\s+/g, '').toLowerCase();
        const 略称 = 作品情報 && 作品情報.作品略称 ? String(作品情報.作品略称).trim() : '';

        worksData.push([
          worksKey,
          新作品ID,
          作品名原題,
          作品名日本語 || '',
          略称,
          0,
          now
        ]);

        worksMap[作品名原題] = {
          idx: worksData.length - 1,
          作品ID: 新作品ID,
          日本語: 作品名日本語 || '',
          略称: 略称
        };
      } else {
        // 日本語名や略称が空なら補完
        const w = worksMap[作品名原題];
        const dataRow = worksData[w.idx];
        if (!String(dataRow[3] || '').trim() && (作品名日本語 || '')) dataRow[3] = 作品名日本語 || '';
        if (!String(dataRow[4] || '').trim() && 作品情報 && 作品情報.作品略称) dataRow[4] = String(作品情報.作品略称).trim();
      }

      // メインシート側の作品IDを補完
      if (worksMap[作品名原題].作品ID) {
        if (set(row, '作品ID(W)(自動)', worksMap[作品名原題].作品ID)) 更新数++;
      }

      const langCode = 台湾グッズ_言語コードへ変換_(言語) || '';
      const 略称 = worksMap[作品名原題].略称 || '';
      const prefix = (langCode && 略称) ? `${langCode}-${略称}-` : '';

      採番対象行.push({
        row,
        rowNum,
        原題: 作品名原題,
        prefix
      });
    }
  }

  /* ============================================================
   * 4) 未登録行だけ SKU を振り直し
   * ============================================================ */
  const prefixGroups = {};
  採番対象行.forEach(item => {
    if (!item.prefix) {
      set(item.row, '商品コードステータス', '作品略称未登録');
      return;
    }
    if (!prefixGroups[item.prefix]) prefixGroups[item.prefix] = [];
    prefixGroups[item.prefix].push(item);
  });

  Object.keys(prefixGroups).forEach(prefix => {
    if (!usedByPrefix[prefix]) usedByPrefix[prefix] = new Set();

    const items = prefixGroups[prefix].sort((a, b) => a.rowNum - b.rowNum);
    let nextNum = 1;

    items.forEach(item => {
      while (usedByPrefix[prefix].has(nextNum)) nextNum++;

      const newCode = `${prefix}${String(nextNum).padStart(4, '0')}`;
      if (set(item.row, '商品コード（SKU）', newCode)) 更新数++;
      if (set(item.row, '商品コードステータス', '生成済み')) 更新数++;

      usedByPrefix[prefix].add(nextNum);
      nextNum++;
    });
  });

  /* ============================================================
   * 5) Works の登録数・更新日時を全件数ベースで更新
   * ============================================================ */
  worksData.forEach(row => {
    const 原題 = String(row[2] || '').trim();
    if (!原題) return;

    const count = 全件数Map[原題] || 0;
    row[5] = count;
    row[6] = now;
  });

  /* ============================================================
   * 6) 一括書き込み
   * ============================================================ */
  sh.getRange(2, 1, values.length, lastCol).setValues(values);
  if (列['粗利益率']) {
    sh.getRange(2, 列['粗利益率'], values.length, 1).setNumberFormat('0.0%');
  }

  // Worksシート必要行数を確保
  const needRows = worksData.length + 1;
  const maxRows = Worksシート.getMaxRows();
  if (needRows > maxRows) {
    Worksシート.insertRowsAfter(maxRows, needRows - maxRows);
  }

  // ヘッダー再設定（念のため）
  Worksシート.getRange(1, 1, 1, worksColCount).setValues([設定_台湾グッズ.作品ヘッダー]);

  if (worksData.length > 0) {
    Worksシート.getRange(2, 1, worksData.length, worksColCount).setValues(worksData);
  }

  ui.alert(`✅ 一括更新完了: ${更新数}件更新\n（Works / SKU 採番は B列=未登録 の行のみ対象）`);
}

function 台湾グッズ_不足項目を返す_(値) {
  const lacks = [];
  if (!String(値.言語 || '').trim()) lacks.push('言語');
  if (!String(値.作品名日本語 || 値.作品名原題 || '').trim()) lacks.push('作品名');
  if (!String(値.商品名日本語 || 値.商品名原題 || '').trim()) lacks.push('商品名');
  return lacks;
}