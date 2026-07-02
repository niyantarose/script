
const MASTER_SPREADSHEET_ID_TW_MAG = '1ZAy5wtQCq1ixl47MMEIGnrVWeae-F_NfN34lSOUsN5M';

let _台湾雑誌GS_共通SSキャッシュ = null;

/* ============================================================
 * 設定
 * ============================================================ */
function 台湾雑誌_設定を取得_() {
  return {
    マスターファイルID: MASTER_SPREADSHEET_ID_TW_MAG,
    マスターシート名: '台湾雑誌',
    雑誌マスター名: '雑誌マスター（共通）',
    版種ルール名: '版種ルール（雑誌共通）',
    タイプルール名: 'タイプルール（雑誌共通）',
    候補シート名: '雑誌マスター候補（共通）',

    // 雑誌は正式マスター直参照で統一
    スナップ雑誌マスター名: '雑誌マスター（共通）',
    スナップ版種ルール名: '版種ルール（雑誌共通）',
    スナップタイプルール名: 'タイプルール（雑誌共通）',

    親コードに言語接頭辞を付ける: true,

    言語リスト: ['台湾', '中国', '香港', 'タイ'],

    言語コードマップ: {
      TW: 'TW', CN: 'CN', HK: 'HK', TH: 'TH',
      KR: 'KR', JP: 'JP', US: 'US', EN: 'EN',
      '台湾': 'TW', '臺灣': 'TW',
      '中国': 'CN', '中國': 'CN',
      '香港': 'HK',
      'タイ': 'TH', '泰国': 'TH', '泰國': 'TH',
      '韓国': 'KR', '韓國': 'KR',
      '日本': 'JP', 'アメリカ': 'US', '英語': 'EN'
    },

    言語表示マップ: {
      TW: '台湾', CN: '中国', HK: '香港', TH: 'タイ',
      KR: '韓国', JP: '日本', US: 'アメリカ', EN: '英語'
    },

    言語表記マップ: {
      TW: '台湾版',
      CN: '中国版',
      HK: '香港版',
      TH: 'タイ版',
      KR: '韓国版',
      JP: '日本版',
      US: 'アメリカ版',
      EN: '英語版'
    },

    監視列: [
      '言語',
      '雑誌名',
      '年',
      '月',
      '号数',
      '版種',
      'バリエーションコード',
      '原題商品タイトル',
      '特典メモ',
      '売価',
      '原価',
      '発番発行'
    ],

    列名: {
      発番発行: '発番発行',
      登録状況: '登録状況',
      商品コード: '商品コード（SKU）',
      商品名出品用: '商品名（出品用）',
      言語: '言語',
      雑誌名: '雑誌名',
      年: '年',
      月: '月',
      号数: '号数',
      版種: '版種',
      バリエーションコード: 'バリエーションコード',
      原題商品タイトル: '原題商品タイトル',
      特典メモ: '特典メモ',
      粗利益率: '粗利益率',
      登録日: '登録日',
      売価: '売価',
      原価: '原価',
      配送パターン: '配送パターン',
      登録者: '登録者',
      商品説明: '商品説明',
      原題タイトル: '原題タイトル',
      原題商品名: '原題商品名',
      博客來商品コード: '博客來商品コード',
      博客來URL: '博客來URL',
      メイン画像URL: 'メイン画像URL',
      追加画像URL: '追加画像URL',
      備考: '備考'
    }
  };
}

/* ============================================================
 * 共通SS
 * ============================================================ */
function 台湾雑誌GS_共通SS_() {
  if (_台湾雑誌GS_共通SSキャッシュ) return _台湾雑誌GS_共通SSキャッシュ;
  _台湾雑誌GS_共通SSキャッシュ = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID_TW_MAG);
  return _台湾雑誌GS_共通SSキャッシュ;
}

function テスト_台湾雑誌GS_共通SS() {
  const ss = 台湾雑誌GS_共通SS_();
  Logger.log(ss.getName());
}

/* ============================================================
 * 基本
 * ============================================================ */
function 台湾雑誌_正規化_(v) {
  return String(v == null ? '' : v)
    .replace(/（/g, '(')
    .replace(/）/g, ')')
    .replace(/／/g, '/')
    .replace(/\u3000/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function 台湾雑誌_英字キー_(v) {
  return String(v || '')
    .normalize('NFKC')
    .toUpperCase()
    .replace(/&/g, ' AND ')
    .replace(/[^\w\p{L}\p{N}]+/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function 台湾雑誌_地域表記を除去_(v) {
  return String(v || '')
    .normalize('NFKC')
    .replace(/\b(?:KOREA|TAIWAN|HONG\s*KONG|HONGKONG|CHINA|THAILAND|HK|TW|CN|KR|TH)\b/gi, ' ')
    .replace(/(?:韓国版|韓國版|韓国|韓國|台湾版|臺灣版|台湾|臺灣|中国版|中國版|中国|中國|香港版|香港|タイ版|タイ)/gu, ' ')
    .replace(/[（(].*?[)）]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function 台湾雑誌_雑誌名キー_(v) {
  return 台湾雑誌_英字キー_(台湾雑誌_地域表記を除去_(v));
}

function 台湾雑誌_英字表示名を整形_(v) {
  const key = 台湾雑誌_英字キー_(v);
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
  return map[key] || String(v || '').trim();
}

function 台湾雑誌_表示雑誌名を決定_(officialName, aliases) {
  const list = [officialName].concat(aliases || []).map(v => String(v || '').trim()).filter(Boolean);

  const unified = list.find(v => {
    const stripped = 台湾雑誌_地域表記を除去_(v);
    return stripped && 台湾雑誌_英字キー_(stripped) === 台湾雑誌_英字キー_(v);
  });
  if (unified) return 台湾雑誌_英字表示名を整形_(unified);

  const strippedOfficial = 台湾雑誌_地域表記を除去_(officialName);
  if (strippedOfficial) return 台湾雑誌_英字表示名を整形_(strippedOfficial);

  return 台湾雑誌_英字表示名を整形_(officialName);
}

function 台湾雑誌_言語コード_(v) {
  const cfg = 台湾雑誌_設定を取得_();
  const raw = 台湾雑誌_正規化_(v);
  const upper = raw.toUpperCase();
  return cfg.言語コードマップ[upper] || cfg.言語コードマップ[raw] || upper || '';
}

function 台湾雑誌_2桁_(v) {
  const n = parseInt(String(v || '').trim(), 10);
  if (isNaN(n)) return '';
  return String(n).padStart(2, '0');
}

function 台湾雑誌_年2桁_(v) {
  const n = parseInt(String(v || '').trim(), 10);
  if (isNaN(n)) return '';
  return String(n).slice(-2);
}

function 台湾雑誌_号数コード_(v) {
  return String(v || '')
    .normalize('NFKC')
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, '')
    .trim();
}

function 台湾雑誌_コード片_(v) {
  return String(v || '')
    .normalize('NFKC')
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, '')
    .trim();
}

function 台湾雑誌_列Map_(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0]
    .map(v => String(v || '').trim());
  const map = {};
  headers.forEach((h, i) => {
    if (h) map[h] = i + 1;
  });
  return map;
}

function 台湾雑誌_列番号_(map, names) {
  for (const name of names) {
    if (map[name]) return map[name];
  }
  return 0;
}

function 台湾雑誌_値_(rowValues, map, names) {
  const col = 台湾雑誌_列番号_(map, names);
  if (!col) return '';
  return 台湾雑誌_正規化_(rowValues[col - 1]);
}

function 台湾雑誌_生値_(rowValues, map, names) {
  const col = 台湾雑誌_列番号_(map, names);
  if (!col) return '';
  return rowValues[col - 1];
}

/* ============================================================
 * ヘッダー / シート整備
 * ============================================================ */
function 台湾雑誌_ヘッダー一覧_() {
  return [
    '発番発行',
    '登録状況',
    '商品コード（SKU）',
    '商品名（出品用）',
    '言語',
    '雑誌名',
    '年',
    '月',
    '号数',
    '版種',
    'バリエーションコード',
    '原題商品タイトル',
    '特典メモ',
    '売価',
    '原価',
    '粗利益率',
    '配送パターン',
    '商品説明',
    '原題タイトル',
    '原題商品名',
    '博客來商品コード',
    '博客來URL',
    'メイン画像URL',
    '追加画像URL',
    '登録日',
    '登録者',
    '備考'
  ];
}

function 台湾雑誌_ヘッダー色_(header) {
  const autoCols = ['商品コード（SKU）', '商品名（出品用）', '粗利益率', '登録日'];
  const apiCols = ['原題タイトル', '原題商品名', '博客來商品コード', '博客來URL', 'メイン画像URL', '追加画像URL', '原価'];

  if (autoCols.includes(header)) return '#999999';
  if (apiCols.includes(header)) return '#e69138';
  return '#4a86e8';
}

function 台湾雑誌_シートを整備() {
  const cfg = 台湾雑誌_設定を取得_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headers = 台湾雑誌_ヘッダー一覧_();
  let sh = ss.getSheetByName(cfg.マスターシート名);

  if (!sh) {
    sh = ss.insertSheet(cfg.マスターシート名);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const current = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getDisplayValues()[0].map(v => String(v || '').trim());
    const missing = headers.filter(h => !current.includes(h));
    if (missing.length) {
      const start = sh.getLastColumn() + 1;
      sh.insertColumnsAfter(sh.getLastColumn(), missing.length);
      sh.getRange(1, start, 1, missing.length).setValues([missing]);
    }
  }

  const map = 台湾雑誌_列Map_(sh);
  headers.forEach(h => {
    const col = map[h];
    if (!col) return;
    sh.getRange(1, col)
      .setValue(h)
      .setBackground(台湾雑誌_ヘッダー色_(h))
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  });

  sh.setRowHeight(1, 40);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(3);

  const maxRows = Math.max(sh.getMaxRows(), 1000);
  const colCheck = 台湾雑誌_列番号_(map, ['発番発行']);
  const colReg = 台湾雑誌_列番号_(map, ['登録状況']);

  if (colCheck) sh.getRange(2, colCheck, maxRows - 1, 1).insertCheckboxes();
  if (colReg) {
    sh.getRange(2, colReg, maxRows - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['未登録', '登録済み'], true)
        .build()
    );
  }

  const dataRange = sh.getRange(2, 1, maxRows - 1, sh.getLastColumn());
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$B2="未登録"')
      .setBackground('#ffff00')
      .setRanges([dataRange])
      .build()
  ]);

  ss.toast(`✅ ${cfg.マスターシート名} を整備しました`, '完了', 4);
}

function 台湾雑誌シートを作成() {
  台湾雑誌_シートを整備();
  SpreadsheetApp.getUi().alert('✅ 台湾雑誌シートを整備しました\n\n続けて「台湾雑誌_プルダウン更新」を実行してください');
}

/* ============================================================
 * マスター読込
 * ============================================================ */
function 台湾雑誌_雑誌マスター一覧_() {
  const cfg = 台湾雑誌_設定を取得_();
  const ss = 台湾雑誌GS_共通SS_();
  const sh = ss.getSheetByName(cfg.雑誌マスター名);
  if (!sh || sh.getLastRow() < 2) return [];

  const map = 台湾雑誌_列Map_(sh);
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getDisplayValues();

  return rows.map(row => {
    const official = 台湾雑誌_値_(row, map, ['雑誌名（英字）']);
    const aliases = ['別名1', '別名2', '別名3', '別名4', '別名5']
      .map(name => 台湾雑誌_値_(row, map, [name]))
      .filter(Boolean);

    const displayName = 台湾雑誌_表示雑誌名を決定_(official, aliases);
    const langs = 台湾雑誌_値_(row, map, ['対応言語'])
      .split(',')
      .map(v => 台湾雑誌_言語コード_(v))
      .filter(Boolean);

    const aliasPool = [official].concat(aliases).concat([displayName]).filter(Boolean);
    const keys = Array.from(new Set(aliasPool.map(v => 台湾雑誌_雑誌名キー_(v)).filter(Boolean)));

    return {
      略称コード: 台湾雑誌_値_(row, map, ['略称コード']),
      雑誌名英字: official,
      雑誌名表示: displayName,
      雑誌名カタカナ: 台湾雑誌_値_(row, map, ['雑誌名（カタカナ）']),
      基本キー型: 台湾雑誌_値_(row, map, ['基本キー型']) || '年月型',
      通常タイプ結合記号: 台湾雑誌_値_(row, map, ['通常タイプ結合記号']) || '-',
      版種ありタイプ結合記号: 台湾雑誌_値_(row, map, ['版種ありタイプ結合記号']) || '-',
      対応言語: langs,
      別名一覧: aliasPool,
      キー一覧: keys
    };
  }).filter(r => r.略称コード && r.雑誌名英字);
}

function 台湾雑誌_雑誌名候補一覧_() {
  const list = [];
  const seen = new Set();

  台湾雑誌_雑誌マスター一覧_().forEach(rec => {
    const v = rec.雑誌名表示 || rec.雑誌名英字;
    const key = 台湾雑誌_英字キー_(v);
    if (!v || seen.has(key)) return;
    seen.add(key);
    list.push(v);
  });

  return list.sort((a, b) => a.localeCompare(b, 'en'));
}

function 台湾雑誌_雑誌レコードを取得_(name, language) {
  const target = 台湾雑誌_雑誌名キー_(name);
  const langCode = 台湾雑誌_言語コード_(language);
  if (!target) return null;

  let best = null;

  台湾雑誌_雑誌マスター一覧_().forEach(rec => {
    let score = -1;
    rec.キー一覧.forEach(key => {
      if (!key) return;
      if (key === target) {
        score = Math.max(score, 1000);
      } else if (target.startsWith(key) || key.startsWith(target)) {
        score = Math.max(score, 700 + Math.min(key.length, target.length));
      } else if (target.includes(key) || key.includes(target)) {
        score = Math.max(score, 500 + Math.min(key.length, target.length));
      }
    });

    if (score < 0) return;

    const langMatched = !rec.対応言語.length || rec.対応言語.includes(langCode);
    score += langMatched ? 100 : -100;

    if (!best || score > best.score) {
      best = { score, rec };
    }
  });

  return best ? best.rec : null;
}

function 台湾雑誌_ルール一覧_(sheetName) {
  const ss = 台湾雑誌GS_共通SS_();
  const sh = ss.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];

  const map = 台湾雑誌_列Map_(sh);
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getDisplayValues();

  const list = rows.map(row => ({
    判定キーワード: 台湾雑誌_値_(row, map, ['判定キーワード']),
    コードsuffix: 台湾雑誌_値_(row, map, ['コードsuffix']),
    タイトル表示: 台湾雑誌_値_(row, map, ['タイトル表示']),
    優先順位: parseInt(台湾雑誌_値_(row, map, ['優先順位']) || '0', 10) || 0,
    有効: (() => {
      const v = 台湾雑誌_値_(row, map, ['有効']).toUpperCase();
      return v === 'TRUE' || v === '1' || v === 'YES' || v === 'ON' || v === '';
    })()
  }))
  .filter(r => r.判定キーワード && r.有効)
  .sort((a, b) => b.優先順位 - a.優先順位);

  return list;
}

function 台湾雑誌_版種候補一覧_() {
  const cfg = 台湾雑誌_設定を取得_();
  const rules = 台湾雑誌_ルール一覧_(cfg.版種ルール名);
  const list = [];
  const seen = new Set();

  // 通常版は空運用でもええけど、明示したい時用
  ['通常版'].concat(rules.map(r => r.タイトル表示 || r.判定キーワード)).forEach(v => {
    const text = 台湾雑誌_正規化_(v);
    if (!text || seen.has(text)) return;
    seen.add(text);
    list.push(text);
  });

  return list;
}

function 台湾雑誌_バリエーション候補一覧_() {
  const list = [];
  for (let i = 65; i <= 90; i++) list.push(String.fromCharCode(i)); // A-Z
  return list;
}

function 台湾雑誌_ルール適用_(inputValue, rules, kind) {
  const raw = 台湾雑誌_正規化_(inputValue);
  if (!raw) return { suffix: '', display: '' };

  // バリエーションは A / B / C など単独コードを優先
  if (kind === 'type') {
    const direct = 台湾雑誌_コード片_(raw);
    if (/^[A-Z0-9]{1,3}$/.test(direct)) {
      return { suffix: direct, display: direct };
    }
    const tail = raw.match(/([A-Z0-9])版?$/i);
    if (tail) {
      return { suffix: tail[1].toUpperCase(), display: tail[1].toUpperCase() };
    }
  }

  const norm = 台湾雑誌_英字キー_(raw);
  for (const rule of rules) {
    const key = 台湾雑誌_英字キー_(rule.判定キーワード);
    if (!key) continue;
    if (norm === key || norm.includes(key) || key.includes(norm)) {
      return {
        suffix: 台湾雑誌_コード片_(rule.コードsuffix),
        display: 台湾雑誌_正規化_(rule.タイトル表示 || raw)
      };
    }
  }

  return {
    suffix: kind === 'type' ? 台湾雑誌_コード片_(raw) : '',
    display: raw
  };
}

/* ============================================================
 * プルダウン更新
 * ============================================================ */
function 台湾雑誌_言語候補一覧_() {
  const ss = 台湾雑誌GS_共通SS_();
  const sh = ss.getSheetByName('言語マスター');
  if (!sh || sh.getLastRow() < 2) return [];

  return sh.getRange(2, 1, sh.getLastRow() - 1, 1)
    .getDisplayValues()
    .flat()
    .map(v => String(v || '').trim())
    .filter(Boolean);
}

function 台湾雑誌_配送パターン候補一覧_() {
  const ss = 台湾雑誌GS_共通SS_();
  const sh = ss.getSheetByName('配送パターンマスター');
  if (!sh || sh.getLastRow() < 2) return [];

  return sh.getRange(2, 1, sh.getLastRow() - 1, 1)
    .getDisplayValues()
    .flat()
    .map(v => String(v || '').trim())
    .filter(Boolean);
}

function 台湾雑誌_プルダウン更新() {
  台湾雑誌_雑誌名プルダウン更新_();
}

function 台湾雑誌_雑誌名プルダウン更新_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 入力側シート
  const 入力シート名 = '台湾雑誌';
  const 入力Sh = ss.getSheetByName(入力シート名);

  if (!入力Sh) {
    SpreadsheetApp.getUi().alert('台湾雑誌シートが見つかりません');
    return;
  }

  // 共通マスターのスプレッドシートID
  // 商品登録_マスター共通
  const MASTER_SPREADSHEET_ID = '1ZAy5wtQCq1ixl47MMEIGnrVWeae-F_NfN34lSOUsN5M';

  const masterSS = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  const 雑誌Sh = masterSS.getSheetByName('雑誌マスター（共通）');

  if (!雑誌Sh || 雑誌Sh.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('雑誌マスター（共通）が見つからないか、データがありません');
    return;
  }

  const 正規化 = (v) => String(v || '')
    .replace(/（/g, '(')
    .replace(/）/g, ')')
    .replace(/\u3000/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  const 列番号候補取得 = (headers, 候補) => {
    const h = headers.map(正規化);
    for (const name of 候補) {
      const idx = h.indexOf(正規化(name));
      if (idx >= 0) return idx + 1;
    }
    return 0;
  };

  const masterHeaders = 雑誌Sh.getRange(1, 1, 1, 雑誌Sh.getLastColumn())
    .getDisplayValues()[0];

  const col雑誌名 = 列番号候補取得(masterHeaders, [
    '雑誌名（英字）',
    '雑誌名(英字)',
    '雑誌名',
    '名称'
  ]);

  if (!col雑誌名) {
    SpreadsheetApp.getUi().alert('雑誌マスター側に「雑誌名（英字）」列が見つかりません');
    return;
  }

  const masterValues = 雑誌Sh
    .getRange(2, col雑誌名, 雑誌Sh.getLastRow() - 1, 1)
    .getDisplayValues();

  const 雑誌名候補 = [...new Set(
    masterValues
      .map(r => 正規化(r[0]))
      .filter(Boolean)
  )];

  if (雑誌名候補.length === 0) {
    SpreadsheetApp.getUi().alert('雑誌名候補が0件です');
    return;
  }

  const inputHeaders = 入力Sh.getRange(1, 1, 1, 入力Sh.getLastColumn())
    .getDisplayValues()[0];

  const col入力雑誌名 = 列番号候補取得(inputHeaders, [
    '雑誌名',
    '雑誌',
    '雑誌名（英字）',
    '雑誌名(英字)'
  ]);

  if (!col入力雑誌名) {
    SpreadsheetApp.getUi().alert('台湾雑誌シート側に「雑誌名」列が見つかりません');
    return;
  }

  // 入力ファイル側に隠し候補シートを作る
  const snapName = '_SNAP_雑誌名候補';
  let snap = ss.getSheetByName(snapName);

  if (!snap) {
    snap = ss.insertSheet(snapName);
  }

  snap.clearContents();

  snap.getRange(1, 1, 雑誌名候補.length, 1)
    .setValues(雑誌名候補.map(v => [v]));

  snap.hideSheet();

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(
      snap.getRange(1, 1, 雑誌名候補.length, 1),
      true
    )
    .setAllowInvalid(false)
    .build();

  入力Sh.getRange(2, col入力雑誌名, 入力Sh.getMaxRows() - 1, 1)
    .setDataValidation(rule);

  ss.toast(
    `雑誌名プルダウン更新完了：${雑誌名候補.length}件`,
    '完了',
    5
  );
}
function 台湾雑誌_雑誌名プルダウンを正式マスターから再設定() {
  return 台湾雑誌_プルダウン更新();
}

/* ============================================================
 * 粗利益率
 * ============================================================ */
function 台湾雑誌_粗利益率を計算_(売価, 原価) {
  const 売 = parseFloat(売価 || 0);
  const 本価格元 = parseFloat(原価 || 0);
  if (!(売 > 0 && 本価格元 >= 0)) return '';

  let レート = 0;

  // まず _kyoutuu から取得
  try {
    if (typeof _kyoutuu !== 'undefined' && typeof _kyoutuu.為替レートを取得_ === 'function') {
      レート =
        _kyoutuu.為替レートを取得_('TWD') ||
        _kyoutuu.為替レートを取得_('台湾ドル') ||
        _kyoutuu.為替レートを取得_('台湾元') ||
        _kyoutuu.為替レートを取得_('台湾') ||
        0;
    }
  } catch (_) {
    レート = 0;
  }

  // ダメなら為替レートマスターの D:E（検索キー / レート）を見る
  if (!(レート > 0)) {
    try {
      const ss = 台湾雑誌GS_共通SS_();
      const sh = ss.getSheetByName('為替レートマスター');
      if (sh && sh.getLastRow() >= 2) {
        const values = sh.getRange(2, 4, sh.getLastRow() - 1, 2).getDisplayValues(); // D:E
        for (const row of values) {
          const key = String(row[0] || '').trim().toUpperCase();
          const rate = parseFloat(String(row[1] || '').trim());
          if (
            !isNaN(rate) &&
            rate > 0 &&
            ['TWD', '台湾', 'TW', '台湾ドル', '台湾元'].includes(key)
          ) {
            レート = rate;
            break;
          }
        }
      }
    } catch (_) {
      レート = 0;
    }
  }

  if (!(レート > 0)) return '';

  const 送料率 = 0.35;
  const モール手数料率 = 0.09;
  const 消費税率 = 0.10;

  const 送料元 = 本価格元 * 送料率;
  const 実原価円 = (本価格元 + 送料元) * レート;
  const 粗利益額 = 売 - (売 * (モール手数料率 + 消費税率)) - 実原価円;
  const 粗利益率 = 粗利益額 / 売;

  return Math.round(粗利益率 * 1000) / 1000;
}
/* ============================================================
 * コード生成
 * ============================================================ */
function 台湾雑誌_商品コードを生成_(opts) {
  const langCode = 台湾雑誌_コード片_(opts.langCode) || 'XX';
  const magCode = 台湾雑誌_コード片_(opts.magazineCode) || '????';
  const keyType = 台湾雑誌_正規化_(opts.keyType) || '年月型';
  const editionSuffix = 台湾雑誌_コード片_(opts.editionSuffix);
  const variationCode = 台湾雑誌_コード片_(opts.variationCode);

  let base = '';

  if (keyType === '号数型') {
    const issue = 台湾雑誌_号数コード_(opts.issueNo) || '??';

    // 号数型は 雑誌コード と 号数 の間にハイフンを入れる
    // 例：TW-BSN-NO5
    base = `${langCode}-${magCode}-${issue}`;

  } else {
    const yy = 台湾雑誌_年2桁_(opts.year) || '??';
    const mm = 台湾雑誌_2桁_(opts.month) || '??';

    // 年月型は従来どおり直結
    // 例：TW-MARI2604-A
    base = `${langCode}-${magCode}${yy}${mm}`;
  }

  // 版種
  if (editionSuffix) {
    base += editionSuffix;
  } else if (opts.useUnknownPlaceholderForEdition) {
    base += '??';
  }

  // バリエーション
  if (variationCode) {
    base += `-${variationCode}`;
  } else if (opts.useUnknownPlaceholderForVariation) {
    base += '-?';
  }

  return base;
}

function 台湾雑誌_バリエーション表示_(display) {
  const d = 台湾雑誌_正規化_(display);
  if (!d) return '';
  if (/^[A-Z0-9]{1,3}$/i.test(d)) return `${d.toUpperCase()}版`;
  return d;
}

function 台湾雑誌_商品名を生成_(cfg, rec, rowData, editionInfo, variationInfo) {
  const parts = [];

  const langCode = 台湾雑誌_言語コード_(rowData.言語);
  const langLabel = cfg.言語表記マップ[langCode] || `${rowData.言語}版`;
  if (langLabel) parts.push(langLabel);

  parts.push('雑誌');

  const magazineDisplay = rec ? (rec.雑誌名表示 || rec.雑誌名英字) : rowData.雑誌名;
  if (magazineDisplay) parts.push(magazineDisplay);

  if ((rec ? rec.基本キー型 : '年月型') === '号数型') {
    if (rowData.号数) parts.push(`${rowData.号数}号`);
  } else {
    if (rowData.年 && rowData.月) parts.push(`${rowData.年}年${parseInt(rowData.月, 10)}月号`);
    else if (rowData.年) parts.push(`${rowData.年}年`);
    else if (rowData.月) parts.push(`${parseInt(rowData.月, 10)}月号`);
  }

  if (editionInfo.display) parts.push(editionInfo.display);
  if (variationInfo.display) parts.push(台湾雑誌_バリエーション表示_(variationInfo.display));
  if (rowData.原題商品タイトル) parts.push(rowData.原題商品タイトル);
  if (rowData.特典メモ) parts.push(rowData.特典メモ);

  return parts.join(' ').replace(/\s+/g, ' ').trim();
}

/* ============================================================
 * 1行再計算
 * ============================================================ */
function 台湾雑誌_1行再計算_(sh, row, options = {}) {
  if (!sh || row < 2) return;

  const cfg = 台湾雑誌_設定を取得_();
  const map = options.列マップ || 台湾雑誌_列Map_(sh);
  const lastCol = options.lastCol || sh.getLastColumn();
  const rowValues = options.rowValues || sh.getRange(row, 1, 1, lastCol).getValues()[0];

  const g = (names) => 台湾雑誌_値_(rowValues, map, Array.isArray(names) ? names : [names]);
  const modifiedCols = [];
  const s = (names, value) => {
    const col = 台湾雑誌_列番号_(map, Array.isArray(names) ? names : [names]);
    if (!col) return false;
    if (String(rowValues[col - 1]) === String(value)) return false;
    rowValues[col - 1] = value;
    if (!modifiedCols.includes(col)) {
      modifiedCols.push(col);
    }
    return true;
  };

  const rowData = {
    言語: g([cfg.列名.言語]),
    雑誌名: g([cfg.列名.雑誌名]),
    年: g([cfg.列名.年]),
    月: g([cfg.列名.月]),
    号数: g([cfg.列名.号数]),
    版種: g([cfg.列名.版種]),
    バリエーションコード: g([cfg.列名.バリエーションコード]),
    原題商品タイトル: g([cfg.列名.原題商品タイトル]),
    特典メモ: g([cfg.列名.特典メモ]),
    売価: g([cfg.列名.売価]),
    原価: g([cfg.列名.原価])
  };

  let changed = false;

  const rec = 台湾雑誌_雑誌レコードを取得_(rowData.雑誌名, rowData.言語);

  const langCode = 台湾雑誌_言語コード_(rowData.言語);
  const editionRules = 台湾雑誌_ルール一覧_(cfg.版種ルール名);
  const typeRules = 台湾雑誌_ルール一覧_(cfg.タイプルール名);

  const editionInfo = 台湾雑誌_ルール適用_(rowData.版種, editionRules, 'edition');
  const variationInfo = 台湾雑誌_ルール適用_(rowData.バリエーションコード, typeRules, 'type');

  const code = 台湾雑誌_商品コードを生成_({
  langCode,
  magazineCode: rec ? rec.略称コード : '',
  year: rowData.年,
  month: rowData.月,
  issueNo: rowData.号数,
  keyType: rec ? rec.基本キー型 : '年月型',
  editionSuffix: editionInfo.suffix,
  variationCode: variationInfo.suffix,
  normalJoin: rec ? rec.通常タイプ結合記号 : '-',
  specialJoin: rec ? rec.版種ありタイプ結合記号 : '-',
  useUnknownPlaceholderForEdition: false,
  useUnknownPlaceholderForVariation: false
});

  const title = 台湾雑誌_商品名を生成_(cfg, rec, rowData, editionInfo, variationInfo);
  const rate = 台湾雑誌_粗利益率を計算_(rowData.売価, rowData.原価);

  changed = s([cfg.列名.商品コード], code) || changed;
  changed = s([cfg.列名.商品名出品用], title) || changed;
  if (rate !== '') changed = s([cfg.列名.粗利益率], rate) || changed;
  if (!g([cfg.列名.登録状況])) changed = s([cfg.列名.登録状況], '未登録') || changed;
  if (code && !g([cfg.列名.登録日])) changed = s([cfg.列名.登録日], new Date()) || changed;

  if (changed) {
    modifiedCols.forEach(colIndex => {
      sh.getRange(row, colIndex).setValue(rowValues[colIndex - 1]);
    });

    const rateCol = 台湾雑誌_列番号_(map, [cfg.列名.粗利益率]);
    if (rateCol && modifiedCols.includes(rateCol)) sh.getRange(row, rateCol).setNumberFormat('0.0%');

    const dateCol = 台湾雑誌_列番号_(map, [cfg.列名.登録日]);
    if (dateCol && modifiedCols.includes(dateCol)) sh.getRange(row, dateCol).setNumberFormat('yyyy/mm/dd');
  }
}

/* ============================================================
 * onEdit / 再計算
 * ============================================================ */
function 台湾雑誌_onEdit(e) {
  if (!e || !e.range) return;

  const cfg = 台湾雑誌_設定を取得_();
  const sh = e.range.getSheet();
  if (!sh || sh.getName() !== cfg.マスターシート名) return;

  const startRow = e.range.getRow();
  const numRows = e.range.getNumRows();
  if (startRow + numRows - 1 < 2) return;

  const map = 台湾雑誌_列Map_(sh);
  const editStartCol = e.range.getColumn();
  const editEndCol = e.range.getLastColumn();
  const watchCols = cfg.監視列
    .map(name => 台湾雑誌_列番号_(map, [name]))
    .filter(Boolean);

  const touched = watchCols.some(c => c >= editStartCol && c <= editEndCol);
  if (!touched) return;

  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(15000);
  } catch (err) {
    return;
  }

  // ループ前に数式を確定し、必要な情報を一括取得
  SpreadsheetApp.flush();

  const lastCol = sh.getLastColumn();
  const allRowValues = sh.getRange(startRow, 1, numRows, lastCol).getValues();

  try {
    for (let r = startRow; r < startRow + numRows; r++) {
      if (r < 2) continue;
      const rowValues = allRowValues[r - startRow];
      if (!rowValues) continue;

      台湾雑誌_1行再計算_(sh, r, {
        列マップ: map,
        lastCol: lastCol,
        rowValues: rowValues
      });
    }
  } finally {
    lock.releaseLock();
  }
}

function 台湾雑誌_取込後補完_(startRow, numRows) {
  const cfg = 台湾雑誌_設定を取得_();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.マスターシート名);
  if (!sh || !numRows || numRows <= 0) return;

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    for (let r = startRow; r < startRow + numRows; r++) {
      if (r < 2) continue;
      台湾雑誌_1行再計算_(sh, r);
    }
  } finally {
    lock.releaseLock();
  }
}

function 台湾雑誌_現在行を再計算() {
  const cfg = 台湾雑誌_設定を取得_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  if (!sh || sh.getName() !== cfg.マスターシート名) {
    SpreadsheetApp.getUi().alert('台湾雑誌シートで実行してください');
    return;
  }

  const row = sh.getActiveCell().getRow();
  if (row < 2) {
    SpreadsheetApp.getUi().alert('2行目以降で実行してください');
    return;
  }

  台湾雑誌_1行再計算_(sh, row);
  ss.toast(`台湾雑誌 ${row}行目を再計算しました`, '完了', 4);
}

function 台湾雑誌_全行再計算() {
  const cfg = 台湾雑誌_設定を取得_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(cfg.マスターシート名);
  if (!sh || sh.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('台湾雑誌シートにデータがありません');
    return;
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    for (let r = 2; r <= sh.getLastRow(); r++) {
      台湾雑誌_1行再計算_(sh, r);
    }
  } finally {
    lock.releaseLock();
  }
  ss.toast('台湾雑誌を全行再計算しました', '完了', 5);
}

function テスト_台湾雑誌_粗利益率() {
  const rate = 台湾雑誌_粗利益率を計算_(2580, 188);
  SpreadsheetApp.getUi().alert('rate=[' + rate + ']');
}

/* ============================================================
 * 雑誌マスター候補（共通）関連
 * 候補シート・マスターはマスター共通ファイル（台湾雑誌GS_共通SS_）側にある。
 * ============================================================ */

/** 候補シートのステータス別件数を表示（読み取りのみ。台湾系言語＝TW/CN/HK/TH と未記入を対象） */
function 台湾雑誌_候補件数を確認() {
  const ui = SpreadsheetApp.getUi();
  const cfg = 台湾雑誌_設定を取得_();
  const sh = 台湾雑誌GS_共通SS_().getSheetByName(cfg.候補シート名);
  if (!sh || sh.getLastRow() < 2) {
    ui.alert('📋 雑誌マスター候補', '候補シートにデータがありません', ui.ButtonSet.OK);
    return;
  }

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0]
    .map(v => String(v || '').trim());
  const colステータス = headers.indexOf('ステータス') + 1;
  const col英字名 = headers.indexOf('雑誌名（英字）') + 1;
  const col言語 = headers.indexOf('対応言語') + 1;

  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getDisplayValues();
  const counts = {};
  let total = 0;
  vals.forEach(r => {
    const name = col英字名 ? String(r[col英字名 - 1] || '').trim() : '';
    if (!name) return; // 空行はスキップ
    const lang = col言語 ? String(r[col言語 - 1] || '').trim().toUpperCase() : '';
    if (lang && !/TW|CN|HK|TH/.test(lang)) return; // 台湾系言語と未記入のみ
    const s = colステータス ? (String(r[colステータス - 1] || '').trim() || '（空欄）') : '（空欄）';
    counts[s] = (counts[s] || 0) + 1;
    total++;
  });

  const 表示順 = ['未対応', '確認中', '登録待ち', 'マスター登録済み', '無視'];
  const lines = [];
  表示順.forEach(s => { if (counts[s]) { lines.push(`${s}: ${counts[s]}件`); delete counts[s]; } });
  Object.keys(counts).forEach(s => lines.push(`${s}: ${counts[s]}件`));

  ui.alert(
    '📋 雑誌マスター候補（台湾系言語＋未記入）',
    (lines.length ? lines.join('\n') : '対象の候補はありません') + `\n\n合計: ${total}件`,
    ui.ButtonSet.OK
  );
}

/** 候補シートで「反映」にチェックした行を、マスター共通ファイルの雑誌マスター（共通）へ反映 */
function 台湾雑誌_候補を正式マスターへ反映() {
  const ui = SpreadsheetApp.getUi();
  if (
    ui.alert(
      '確認',
      '雑誌マスター候補（共通）で「反映」にチェックした行を、雑誌マスター（共通）へ反映します。\n' +
      '（マスター共通ファイル側を更新します。この商品ファイルは変更しません）\n\n続行しますか？',
      ui.ButtonSet.OK_CANCEL
    ) !== ui.Button.OK
  ) {
    return;
  }

  // 反映本体はライブラリ（Project_08 ★雑誌共通マスター）。getActive固定だったため
  // マスター共通SSを引数で渡せるようにした版を呼ぶ（無引数の既存動作は不変）。
  const r = _kyoutuu.共通雑誌候補をマスターへ反映(台湾雑誌GS_共通SS_());
  ui.alert(
    '完了',
    r
      ? `新規追加: ${r.追加}件\n別名として追加: ${r.別名追加}件\n既存一致: ${r.一致}件`
      : '候補データがありませんでした（または反映チェックなし）',
    ui.ButtonSet.OK
  );
}