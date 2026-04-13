/**
 * 韓国雑誌.gs
 * 韓国雑誌シート専用の設定・シート作成・onEdit・確定発行
 *
 * 【設計方針】
 * - 正本マスターは共通ファイル「雑誌マスター（共通）」
 * - このファイルの「雑誌マスター（韓国）」は同期コピー
 * - 未解決雑誌は共通ファイル「雑誌マスター候補（共通）」へ追加
 *
 * ★ コード体系
 *   年月型: {略称}{年下2桁}{月2桁}   例: MAXI2603
 *   号数型: {略称}{号数}              例: 1STLOOK127
 */

const 韓国雑誌_MASTER_SPREADSHEET_ID = '1ZAy5wtQCq1ixl47MMEIGnrVWeae-F_NfN34lSOUsN5M';

const 設定_韓国雑誌 = {
  マスターシート名: '韓国雑誌',
  雑誌マスター名: '雑誌マスター（韓国）',
  共通雑誌マスター名: '雑誌マスター（共通）',
  候補シート名: '雑誌マスター候補（共通）',
  配送パターンマスター名: '配送パターンマスター',

  列名: {
    発番発行:         '発番発行',
    登録状況:         '登録状況',
    親コード:         '商品コード(SKU)',
    商品名出品用:     '商品名（出品用）',
    言語:             '言語',
    雑誌名:           '雑誌名',
    年:               '年',
    月:               '月',
    号数:             '号数',
    表紙情報:         '表紙情報',
    売価:             '売価',
    原価:             '原価',
    原題タイトル:     '原題タイトル',
    原題商品名:       '原題商品名',
    粗利益率:         '粗利益率',
    配送パターン:     '配送パターン',
    特典メモ:         '特典メモ',
    商品説明:         '商品説明',
    サイト商品コード: 'サイト商品コード',
    リンク:           'リンク',
    メイン画像URL:    'メイン画像',
    追加画像URL:      '追加画像',
    作品ID:           '作品ID(W)(自動)',
    コードステータス: '商品コードステータス',
    登録日:           '登録日',
    登録者:           '登録者',
  }
};

const 候補ステータス = {
  未対応: '未対応',
  確認中: '確認中',
  登録待ち: '登録待ち',
  登録済み: 'マスター登録済み',
  無視: '無視'
};

const 候補列 = {
  ステータス: 1,
  反映: 2,
  略称コード: 3,
  英字名: 4,
  カタカナ名: 5,
  基本キー型: 6,
  通常タイプ結合記号: 7,
  版種ありタイプ結合記号: 8,
  言語: 9,
  別名1: 10,
  別名2: 11,
  別名3: 12,
  備考: 13,
  元ファイル: 14,
  元シート: 15,
  元行: 16,
  初回検出: 17,
  最終検出: 18,
  正規化キー: 19
};
const 候補列数 = 19;

let _韓国雑誌_共通SSキャッシュ = null;

/* ============================================================
 * 基本ヘルパー
 * ============================================================ */

function 韓国雑誌_共通SS_() {
  if (_韓国雑誌_共通SSキャッシュ) return _韓国雑誌_共通SSキャッシュ;
  _韓国雑誌_共通SSキャッシュ = SpreadsheetApp.openById(韓国雑誌_MASTER_SPREADSHEET_ID);
  return _韓国雑誌_共通SSキャッシュ;
}

function 韓国雑誌_安全toast_(msg, title = '完了', sec = 3) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, title, sec);
  } catch (_) {}
}

function 韓国雑誌_安全alert_(msg, title) {
  try {
    const ui = SpreadsheetApp.getUi();
    if (title) ui.alert(title, msg, ui.ButtonSet.OK);
    else ui.alert(msg);
  } catch (_) {
    Logger.log(msg);
  }
}

function 韓国雑誌_ヘッダーMap_(sh) {
  const vals = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getDisplayValues()[0];
  const map = {};
  vals.forEach((h, i) => {
    const key = String(h || '').trim();
    if (key) map[key] = i + 1;
  });
  return map;
}

function 韓国雑誌_比較文字列_(value) {
  return String(value || '')
    .normalize('NFKC')
    .trim()
    .replace(/[　\s]+/g, ' ')
    .toUpperCase();
}

function 韓国雑誌_言語コード_(value) {
  const raw = String(value || '').trim();
  const normalized = raw.toUpperCase();
  const map = {
    KR: 'KR',
    '韓国': 'KR',
    '韓國': 'KR'
  };
  return map[normalized] || map[raw] || '';
}

function 韓国雑誌_必要サイズを確保_(sh, rows, cols) {
  if (cols > sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());
  if (rows > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
}

function 韓国雑誌_列番号を英字へ_(colNum) {
  let s = '';
  let n = colNum;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/* ============================================================
 * 雑誌名正規化
 * ============================================================ */

const 韓国雑誌_名称パターン = [
  { name: 'DAZED & CONFUSED', pattern: /dazed\s*&\s*confused/i },
  { name: "HARPER'S BAZAAR", pattern: /harper'?s\s*bazaar/i },
  { name: 'BAZAAR', pattern: /\bbazaar\b/i },
  { name: 'MARIE CLAIRE', pattern: /marie\s*claire/i },
  { name: 'COSMOPOLITAN', pattern: /cosmopolitan/i },
  { name: 'ESQUIRE', pattern: /esquire/i },
  { name: 'ALLURE', pattern: /allure/i },
  { name: 'VOGUE', pattern: /vogue/i },
  { name: 'ELLE', pattern: /elle/i },
  { name: 'GQ', pattern: /\bgq\b/i },
  { name: 'ARENA', pattern: /arena/i },
  { name: 'W KOREA', pattern: /\bw\s*korea\b/i },
  { name: 'CECI', pattern: /\bceci\b/i },
  { name: '1ST LOOK', pattern: /1st\s*look/i },
  { name: 'CINE21', pattern: /cine\s*21/i },
  { name: 'MAXIM', pattern: /maxim/i },
  { name: 'SINGLES', pattern: /singles/i },
  { name: 'NYLON', pattern: /nylon/i },
  { name: 'THE STAR', pattern: /the\s*star/i },
  { name: 'STAR1', pattern: /star\s*1|star1/i },
  { name: "MEN'S HEALTH", pattern: /men'?s\s*health/i }
];

function 韓国雑誌_雑誌名を正規化_(value) {
  const text = String(value || '').trim();
  if (!text) return '';

  for (const entry of 韓国雑誌_名称パターン) {
    if (entry.pattern.test(text)) return entry.name;
  }

  return text
    .replace(/\b20\d{2}[./-]\d{1,2}\b.*$/i, '')
    .replace(/\b\d{4}년\s*\d{1,2}월.*$/i, '')
    .replace(/\b\d{4}年\s*\d{1,2}月.*$/i, '')
    .replace(/\s+(KOREA|KOR|코리아)\b/ig, '')
    .replace(/[<＜【\[(].*$/, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
}

function 韓国雑誌_略称候補_(name) {
  const cleaned = 韓国雑誌_比較文字列_(name).replace(/\s+/g, '');
  return cleaned ? cleaned.slice(0, 8) : '';
}

/* ============================================================
 * ローカル韓国マスター作成
 * ============================================================ */

function 韓国雑誌マスターを作成() {
  const ss = SpreadsheetApp.getActive();
  const シート名 = 設定_韓国雑誌.雑誌マスター名;
  if (ss.getSheetByName(シート名)) {
    韓国雑誌_安全alert_(`「${シート名}」シートは既に存在します`);
    return;
  }

  const sh = ss.insertSheet(シート名);
  const ヘッダー = ['雑誌名（英字）', '雑誌名（カタカナ）', '略称コード', 'コード型', '備考'];
  sh.getRange(1, 1, 1, ヘッダー.length).setValues([ヘッダー]);

  sh.getRange(1, 1, 1, ヘッダー.length)
    .setBackground('#4a86e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  [200, 160, 120, 80, 220].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  韓国雑誌_安全toast_('韓国雑誌マスターを作成しました', '完了', 3);
}

/* ============================================================
 * 韓国雑誌シート作成
 * ============================================================ */

function 韓国雑誌シートを作成() {
  const ss = SpreadsheetApp.getActive();
  const シート名 = 設定_韓国雑誌.マスターシート名;
  if (ss.getSheetByName(シート名)) {
    韓国雑誌_安全alert_(`「${シート名}」シートは既に存在します`);
    return;
  }

  const sh = ss.insertSheet(シート名);

  const ヘッダー = [
    '発番発行', '登録状況', '雑誌名', '年', '月', '号数', '表紙情報', '特典メモ',
    '親コード', '商品名（出品用）', '粗利益率', '登録日',
    '売価', '配送パターン', '登録者', '商品説明',
    '原価', '原題タイトル', '原題商品名', 'アラジン商品コード', 'アラジンURL',
    'メイン画像URL', '追加画像URL'
  ];
  sh.getRange(1, 1, 1, ヘッダー.length).setValues([ヘッダー]);

  const 色 = {
    '発番発行': '#cc0000',
    '登録状況': '#4a86e8',
    '雑誌名': '#6aa84f',
    '年': '#6aa84f',
    '月': '#6aa84f',
    '号数': '#6aa84f',
    '表紙情報': '#f1c232',
    '特典メモ': '#f1c232',
    '親コード': '#999999',
    '商品名（出品用）': '#999999',
    '粗利益率': '#999999',
    '登録日': '#999999',
    '売価': '#4a86e8',
    '配送パターン': '#4a86e8',
    '登録者': '#4a86e8',
    '商品説明': '#4a86e8',
    '原価': '#e69138',
    '原題タイトル': '#e69138',
    '原題商品名': '#e69138',
    'アラジン商品コード': '#e69138',
    'アラジンURL': '#e69138',
    'メイン画像URL': '#e69138',
    '追加画像URL': '#e69138'
  };

  ヘッダー.forEach((h, i) => {
    sh.getRange(1, i + 1)
      .setBackground(色[h] || '#cccccc')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  });

  sh.setRowHeight(1, 40);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);

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
      .build()
  ]);

  const 列幅マップ = {
    '発番発行': 60, '登録状況': 80, '雑誌名': 180, '年': 60, '月': 50, '号数': 70,
    '表紙情報': 200, '特典メモ': 240,
    '親コード': 140, '商品名（出品用）': 360, '粗利益率': 90, '登録日': 120,
    '売価': 80, '配送パターン': 100, '登録者': 80, '商品説明': 150,
    '原価': 80, '原題タイトル': 200, '原題商品名': 200,
    'アラジン商品コード': 140, 'アラジンURL': 220, 'メイン画像URL': 200, '追加画像URL': 200
  };
  ヘッダー.forEach((h, i) => {
    if (列幅マップ[h]) sh.setColumnWidth(i + 1, 列幅マップ[h]);
  });

  韓国雑誌_安全alert_(
    '✅ 韓国雑誌シートを作成しました\n\n' +
    '続けて「韓国雑誌_プルダウン更新」を実行してください'
  );
}

/* ============================================================
 * 共通マスター → ローカル韓国マスター同期
 * ============================================================ */

function 韓国雑誌_共通マスターから同期_() {
  const localSS = SpreadsheetApp.getActive();
  let localSh = localSS.getSheetByName(設定_韓国雑誌.雑誌マスター名);
  if (!localSh) {
    localSh = localSS.insertSheet(設定_韓国雑誌.雑誌マスター名);
  }

  const commonSS = 韓国雑誌_共通SS_();
  const commonSh = commonSS.getSheetByName(設定_韓国雑誌.共通雑誌マスター名);
  if (!commonSh || commonSh.getLastRow() < 2) return 0;

  const map = 韓国雑誌_ヘッダーMap_(commonSh);
  const rows = commonSh.getRange(2, 1, commonSh.getLastRow() - 1, commonSh.getLastColumn()).getValues();

  const out = [];
  const seen = {};

  rows.forEach(r => {
    const langs = String(r[(map['対応言語'] || 0) - 1] || '').replace(/\s+/g, '').toUpperCase();
    if (!langs) return;
    if (!langs.split(',').includes('KR')) return;

    const 英字名 = String(r[(map['雑誌名（英字）'] || 0) - 1] || '').trim();
    const カタカナ = String(r[(map['雑誌名（カタカナ）'] || 0) - 1] || '').trim();
    const 略称 = String(r[(map['略称コード'] || 0) - 1] || '').trim();
    const 基本キー型 = String(r[(map['基本キー型'] || 0) - 1] || '年月型').trim() || '年月型';
    const 備考 = String(r[(map['備考'] || 0) - 1] || '').trim();

    if (!英字名 || !略称) return;

    const key = 韓国雑誌_雑誌名を正規化_(英字名);
    if (seen[key]) return;
    seen[key] = true;

    out.push([英字名, カタカナ, 略称, 基本キー型, 備考]);
  });

  localSh.clearContents();
  localSh.getRange(1, 1, 1, 5).setValues([['雑誌名（英字）', '雑誌名（カタカナ）', '略称コード', 'コード型', '備考']]);

  if (out.length > 0) {
    韓国雑誌_必要サイズを確保_(localSh, out.length + 1, 5);
    localSh.getRange(2, 1, out.length, 5).setValues(out);
    localSh.getRange(2, 4, out.length, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['年月型', '号数型'], true)
        .build()
    );
  }

  localSh.getRange(1, 1, 1, 5)
    .setBackground('#4a86e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  localSh.setFrozenRows(1);
  [200, 160, 120, 80, 220].forEach((w, i) => localSh.setColumnWidth(i + 1, w));

  return out.length;
}

/* ============================================================
 * ローカル韓国マスター検索
 * ============================================================ */

function 韓国雑誌_マスターを検索_(雑誌名英字) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(設定_韓国雑誌.雑誌マスター名);
  if (!sh || sh.getLastRow() < 2) return null;

  const target = 韓国雑誌_雑誌名を正規化_(雑誌名英字);
  if (!target) return null;

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  for (const row of data) {
    const masterName = String(row[0] || '').trim();
    if (韓国雑誌_雑誌名を正規化_(masterName) === target) {
      return {
        英字名: masterName,
        表示名: 韓国雑誌_雑誌名を正規化_(masterName),
        カタカナ名: String(row[1] || '').trim(),
        略称: String(row[2] || '').trim(),
        コード型: String(row[3] || '').trim() || '年月型'
      };
    }
  }
  return null;
}

function 韓国雑誌_マスターに完全一致するか_(候補名) {
  return !!韓国雑誌_マスターを検索_(候補名);
}

/* ============================================================
 * コード生成 / 商品名生成
 * ============================================================ */

function 韓国雑誌_親コードを生成_(雑誌名, 年, 月, 号数) {
  const info = 韓国雑誌_マスターを検索_(雑誌名);
  if (!info || !info.略称) return 'ERROR:マスター未登録';

  if (info.コード型 === '号数型') {
    if (!号数) return 'ERROR:号数未入力';
    return `${info.略称}${String(号数).trim()}`;
  }

  if (!年 || !月) return 'ERROR:年月未入力';
  return `${info.略称}${String(年).slice(-2)}${String(月).padStart(2, '0')}`;
}

function 韓国雑誌_出品用商品名を生成_(雑誌名, 年, 月, 号数, 表紙情報, 特典メモ) {
  if (!雑誌名) return '';

  const info = 韓国雑誌_マスターを検索_(雑誌名);
  const 英字 = info ? (info.表示名 || 韓国雑誌_雑誌名を正規化_(info.英字名)) : 韓国雑誌_雑誌名を正規化_(雑誌名);
  const カタカナ = info ? info.カタカナ名 : '';
  const コード型 = info ? info.コード型 : '年月型';

  let 名前 = `韓国 雑誌 ${英字}`;
  if (カタカナ) 名前 += ` (${カタカナ})`;

  if (コード型 === '号数型') {
    if (!号数) return '';
    名前 += ` Vol.${String(号数).trim()}`;
    if (年 && 月) 名前 += ` (${年}年${月}月)`;
  } else {
    if (!年 || !月) return '';
    名前 += ` ${年}年 ${monthAsNumber_(月)}月号`;
  }

  if (表紙情報 && String(表紙情報).trim()) 名前 += ` (${String(表紙情報).trim()})`;
  if (特典メモ && String(特典メモ).trim()) 名前 += ` ${String(特典メモ).trim()}`;

  return 名前;
}

function monthAsNumber_(v) {
  const n = parseInt(String(v || '').trim(), 10);
  return isNaN(n) ? String(v || '').trim() : n;
}

/* ============================================================
 * 行再計算
 * ============================================================ */

function 韓国雑誌_行を再計算_(sh, row) {
  if (!sh || sh.getName() !== '韓国雑誌' || row < 2) return;

  const 列 = 韓国雑誌_ヘッダーMap_(sh);
  const get = (名前) => 列[名前] ? sh.getRange(row, 列[名前]).getValue() : '';
  const set = (名前, v) => { if (列[名前]) sh.getRange(row, 列[名前]).setValue(v); };

  const 雑誌名入力 = get('雑誌名');
  const 雑誌名 = 韓国雑誌_雑誌名を正規化_(雑誌名入力);
  const 年 = String(get('年') || '').trim();
  const 月 = String(get('月') || '').trim();
  const 号数 = String(get('号数') || '').trim();
  const 表紙情報 = get('表紙情報');
  const 特典メモ = get('特典メモ');

  if (雑誌名 && 雑誌名 !== String(雑誌名入力 || '').trim()) {
    set('雑誌名', 雑誌名);
  }

  if (!雑誌名) return;

  set('原題タイトル', 雑誌名);

  const 親コード = 韓国雑誌_親コードを生成_(雑誌名, 年, 月, 号数);
set('商品コード(SKU)', 親コード);

  const 商品名 = 韓国雑誌_出品用商品名を生成_(雑誌名, 年, 月, 号数, 表紙情報, 特典メモ);
  if (商品名) set('商品名（出品用）', 商品名);

  if (雑誌名 && !get('登録日')) set('登録日', new Date());
}

/* ============================================================
 * onEdit
 * ============================================================ */

function 韓国雑誌_onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== '韓国雑誌') return;

  const row = e.range.getRow();
  if (row < 2) return;
  if (typeof 自己更新中か_ === 'function' && 自己更新中か_()) return;

  const lock = LockService.getDocumentLock();
  try {
    if (!lock.tryLock(5000)) return;
  } catch (_) {
    return;
  }

  try {
    if (typeof 自己更新を開始_ === 'function') 自己更新を開始_();

    const 列 = 韓国雑誌_ヘッダーMap_(sh);
    const col = e.range.getColumn();
    const 編集列名 = Object.keys(列).find(h => 列[h] === col);
    const 監視列 = ['雑誌名', '年', '月', '号数', '表紙情報', '特典メモ'];
    if (!監視列.includes(編集列名)) return;

    韓国雑誌_行を再計算_(sh, row);

    if (編集列名 === '雑誌名') {
      韓国雑誌_未解決候補を自動追加_(sh, row);
    }
  } finally {
    if (typeof 自己更新を終了_ === 'function') 自己更新を終了_();
    lock.releaseLock();
  }
}

/* ============================================================
 * 確定発行
 * ============================================================ */

function 韓国雑誌_重複キーを作成_(雑誌名, 年, 月, 号数) {
  const normalizedName = 韓国雑誌_雑誌名を正規化_(雑誌名);
  if (!normalizedName) return '';
  if (号数) return `${normalizedName}-${号数}`;
  if (年 && 月) return `${normalizedName}-${年}-${月}`;
  return '';
}

function 韓国雑誌_確定発行() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('韓国雑誌');
  if (!sh || sh.getLastRow() < 2) { 韓国雑誌_安全alert_('データがありません'); return; }

  const 列 = 韓国雑誌_ヘッダーMap_(sh);
  if (!列['発番発行']) { 韓国雑誌_安全alert_('発番発行列が見つかりません'); return; }

  const 全データ = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const get行 = (row, 名前) => 列[名前] ? row[列[名前] - 1] : '';

  const 対象行リスト = [];
  全データ.forEach((row, i) => {
    if (get行(row, '発番発行') === true) 対象行リスト.push({ row, rowNum: i + 2 });
  });

  if (対象行リスト.length === 0) { 韓国雑誌_安全alert_('発番発行列にチェックが入っている行がありません'); return; }

  // 重複チェック
  const 重複チェックMap = {};
  全データ.forEach((row, i) => {
    const key = 韓国雑誌_重複キーを作成_(
      String(get行(row, '雑誌名') || '').trim(),
      String(get行(row, '年') || '').trim(),
      String(get行(row, '月') || '').trim(),
      String(get行(row, '号数') || '').trim()
    );
    if (!key) return;
    if (!重複チェックMap[key]) 重複チェックMap[key] = [];
    重複チェックMap[key].push(i + 2);
  });

  const ブロックリスト = [];
  対象行リスト.forEach(({ row, rowNum }) => {
    const key = 韓国雑誌_重複キーを作成_(
      String(get行(row, '雑誌名') || '').trim(),
      String(get行(row, '年') || '').trim(),
      String(get行(row, '月') || '').trim(),
      String(get行(row, '号数') || '').trim()
    );
    if (!key) return;
    const 重複行 = (重複チェックMap[key] || []).filter(r => r !== rowNum);
    if (重複行.length > 0) ブロックリスト.push(`${rowNum}行目：「${key}」が${重複行[0]}行目と重複`);
  });

  if (ブロックリスト.length > 0) {
    韓国雑誌_安全alert_(`⚠️ 重複が見つかりました。\n\n${ブロックリスト.join('\n')}\n\n重複を解消してから再度実行してください。`);
    return;
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(60000);
  try {
    let 発行数 = 0;
    対象行リスト.forEach(({ row, rowNum }) => {
      const 親コード = String(get行(row, '商品コード(SKU)') || '').trim();
      if (!親コード || 親コード.startsWith('ERROR')) return;
      sh.getRange(rowNum, 列['登録状況']).setValue('登録済み');
      sh.getRange(rowNum, 列['発番発行']).setValue(false);
      発行数++;
    });

    // チェックを全解除
    const 最終行 = sh.getLastRow();
    if (列['発番発行'] && 最終行 >= 2) {
      sh.getRange(2, 列['発番発行'], 最終行 - 1, 1).setValues(Array(最終行 - 1).fill([false]));
    }

    韓国雑誌_安全alert_(`✅ 確定発行完了: ${発行数}件`);
  } finally { lock.releaseLock(); }
}

/* ============================================================
 * 候補シート確保
 * ============================================================ */

function 韓国雑誌_候補シートを確保_() {
  const ss = 韓国雑誌_共通SS_();
  let sh = ss.getSheetByName(設定_韓国雑誌.候補シート名);
  if (sh) return sh;

  sh = ss.insertSheet(設定_韓国雑誌.候補シート名);

  const headers = [
    'ステータス', '反映', '略称コード候補',
    '雑誌名（英字）', '雑誌名（カタカナ）', '基本キー型候補',
    '通常タイプ結合記号', '版種ありタイプ結合記号', '対応言語',
    '別名1', '別名2', '別名3', '備考',
    '元ファイル', '元シート', '元行',
    '初回検出日時', '最終検出日時', '正規化キー'
  ];

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

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

  sh.getRange(2, 1, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(
        [
          候補ステータス.未対応,
          候補ステータス.確認中,
          候補ステータス.登録待ち,
          候補ステータス.登録済み,
          候補ステータス.無視
        ],
        true
      )
      .build()
  );

  sh.getRange(2, 2, 1000, 1).insertCheckboxes();

  sh.getRange(2, 6, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['年月型', '号数型'], true)
      .build()
  );

  const range = sh.getRange(2, 1, 1000, 13);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2="未対応"')
      .setBackground('#fff2cc')
      .setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2="確認中"')
      .setBackground('#fce5cd')
      .setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2="登録待ち"')
      .setBackground('#cfe2f3')
      .setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2="マスター登録済み"')
      .setBackground('#d9ead3')
      .setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2="無視"')
      .setBackground('#efefef')
      .setFontColor('#999999')
      .setRanges([range]).build()
  ]);

  const 列幅 = {
    1: 120, 2: 45, 3: 120, 4: 220, 5: 180,
    6: 110, 7: 90, 8: 140, 9: 90,
    10: 160, 11: 160, 12: 160, 13: 320,
    14: 160, 15: 120, 16: 70, 17: 140, 18: 140, 19: 160
  };
  Object.entries(列幅).forEach(([col, width]) => sh.setColumnWidth(Number(col), width));

  sh.hideColumns(19);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);

  return sh;
}

function 韓国雑誌_候補シートの既存行番号_(candidateSh, 正規化キー) {
  if (!candidateSh || candidateSh.getLastRow() < 2 || !正規化キー) return -1;

  const values = candidateSh
    .getRange(2, 候補列.正規化キー, candidateSh.getLastRow() - 1, 1)
    .getDisplayValues().flat();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i] || '').trim() === 正規化キー) return i + 2;
  }
  return -1;
}

/* ============================================================
 * 未解決候補 自動追加
 * ============================================================ */

function 韓国雑誌_未解決候補を自動追加_(sh, row) {
  if (!sh || sh.getName() !== '韓国雑誌' || row < 2) return false;

  const 列 = 韓国雑誌_ヘッダーMap_(sh);
  const get = (名前) => 列[名前] ? String(sh.getRange(row, 列[名前]).getValue() || '').trim() : '';

  const 親コード = get('親コード');
  if (!String(親コード).startsWith('ERROR:マスター未登録')) return false;

  const 雑誌名 = get('雑誌名');
  const 表紙情報 = get('表紙情報');
  const 特典メモ = get('特典メモ');

  const 候補英字名 = 韓国雑誌_雑誌名を正規化_(雑誌名);
  if (!候補英字名) return false;

  const 正規化キー = `KR|${候補英字名.toUpperCase()}`;

  if (韓国雑誌_マスターに完全一致するか_(候補英字名)) return false;

  const candidateSh = 韓国雑誌_候補シートを確保_();

  const 既存行 = 韓国雑誌_候補シートの既存行番号_(candidateSh, 正規化キー);
  if (既存行 >= 2) {
    candidateSh.getRange(既存行, 候補列.最終検出).setValue(new Date());
    return false;
  }

  const 備考 = [
    表紙情報 ? `表紙:${表紙情報}` : '',
    特典メモ ? `特典:${特典メモ}` : ''
  ].filter(Boolean).join(' / ');

  const now = new Date();
  const writeRow = Math.max(candidateSh.getLastRow() + 1, 2);
  韓国雑誌_必要サイズを確保_(candidateSh, writeRow, 候補列数);

  candidateSh.getRange(writeRow, 1, 1, 候補列数).setValues([[
    候補ステータス.未対応,
    false,
    韓国雑誌_略称候補_(候補英字名),
    候補英字名,
    '',
    '年月型',
    '',
    '-',
    'KR',
    '',
    '',
    '',
    備考,
    sh.getParent().getName(),
    sh.getName(),
    row,
    now,
    now,
    正規化キー
  ]]);

  candidateSh.getRange(writeRow, 候補列.反映).insertCheckboxes();

  韓国雑誌_安全toast_(`未解決雑誌を候補シートへ追加: ${候補英字名}`, '📋 雑誌マスター候補', 4);
  return true;
}

/* ============================================================
 * 候補 → ローカル韓国マスター反映
 * ============================================================ */

function 韓国雑誌_候補を正式マスターへ反映() {
  const candidateSh = 韓国雑誌_共通SS_().getSheetByName(設定_韓国雑誌.候補シート名);
  if (!candidateSh || candidateSh.getLastRow() < 2) {
    韓国雑誌_安全alert_('候補シートにデータがありません');
    return;
  }

  const ss = SpreadsheetApp.getActive();
  const 韓国マスター = ss.getSheetByName(設定_韓国雑誌.雑誌マスター名);
  if (!韓国マスター) {
    韓国雑誌_安全alert_('雑誌マスター（韓国）シートが見つかりません');
    return;
  }

  const 全データ = candidateSh.getRange(2, 1, candidateSh.getLastRow() - 1, 候補列数).getValues();

  const 対象行リスト = [];
  全データ.forEach((row, i) => {
    if (row[候補列.反映 - 1] === true) {
      対象行リスト.push({ data: row, rowNum: i + 2 });
    }
  });

  if (対象行リスト.length === 0) {
    韓国雑誌_安全alert_('「反映」にチェックが入っている行がありません');
    return;
  }

  let 反映数 = 0;
  const スキップリスト = [];
  const 反映済み行番号 = [];

  対象行リスト.forEach(({ data, rowNum }) => {
    const 言語 = String(data[候補列.言語 - 1] || '').trim().toUpperCase();
    if (!(言語 === 'KR' || 言語 === '韓国' || 言語 === '')) return;

    const 英字名 = String(data[候補列.英字名 - 1] || '').trim();
    const 略称 = String(data[候補列.略称コード - 1] || '').trim();
    const カタカナ = String(data[候補列.カタカナ名 - 1] || '').trim();
    const 基本キー型 = String(data[候補列.基本キー型 - 1] || '').trim() || '年月型';
    const コード型 = 基本キー型 === '号数型' ? '号数型' : '年月型';

    if (!英字名) {
      スキップリスト.push(`${rowNum}行目: 雑誌名（英字）が空`);
      return;
    }
    if (!略称) {
      スキップリスト.push(`${rowNum}行目 [${英字名}]: 略称コードが空`);
      return;
    }
    if (韓国雑誌_マスターに完全一致するか_(英字名)) {
      スキップリスト.push(`${rowNum}行目 [${英字名}]: 既にローカル韓国マスターに登録済み`);
      candidateSh.getRange(rowNum, 候補列.ステータス).setValue(候補ステータス.登録済み);
      candidateSh.getRange(rowNum, 候補列.反映).setValue(false);
      return;
    }

    const writeRow = Math.max(韓国マスター.getLastRow() + 1, 2);
    韓国マスター.getRange(writeRow, 1, 1, 5).setValues([[
      英字名, カタカナ, 略称, コード型, ''
    ]]);

    反映済み行番号.push(rowNum);
    反映数++;
  });

  反映済み行番号.forEach(rowNum => {
    candidateSh.getRange(rowNum, 候補列.ステータス).setValue(候補ステータス.登録済み);
    candidateSh.getRange(rowNum, 候補列.反映).setValue(false);
  });

  let msg = `✅ 反映完了: ${反映数}件`;
  if (スキップリスト.length > 0) {
    msg += `\n\n⚠️ スキップ:\n${スキップリスト.join('\n')}`;
  }
  msg += '\n\n続けて「韓国雑誌_プルダウン更新」を実行してください';
  韓国雑誌_安全alert_(msg);
}

/* ============================================================
 * 候補件数確認
 * ============================================================ */

function 韓国雑誌_候補件数を確認() {
  const sh = 韓国雑誌_共通SS_().getSheetByName(設定_韓国雑誌.候補シート名);
  if (!sh || sh.getLastRow() < 2) {
    韓国雑誌_安全toast_('候補シートにデータがありません', '📋 雑誌マスター候補', 3);
    return;
  }

  const 全行 = sh.getRange(2, 1, sh.getLastRow() - 1, 候補列数).getValues();
  const KR行 = 全行.filter(row => {
    const 言語 = String(row[候補列.言語 - 1] || '').trim().toUpperCase();
    return 言語 === 'KR' || 言語 === '';
  });

  const カウント = {};
  Object.values(候補ステータス).forEach(s => { カウント[s] = 0; });

  KR行.forEach(row => {
    const s = String(row[候補列.ステータス - 1] || '');
    if (typeof カウント[s] !== 'undefined') カウント[s]++;
  });

  const msg = [
    '【韓国雑誌候補】',
    `未対応: ${カウント[候補ステータス.未対応]}件`,
    `確認中: ${カウント[候補ステータス.確認中]}件`,
    `登録待ち: ${カウント[候補ステータス.登録待ち]}件`,
    `登録済み: ${カウント[候補ステータス.登録済み]}件`,
    `無視: ${カウント[候補ステータス.無視]}件`
  ].join('\n');

  韓国雑誌_安全toast_(msg, '📋 雑誌マスター候補', 8);
}

/* ============================================================
 * プルダウン更新
 * ============================================================ */

function 韓国雑誌_プルダウン更新() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('韓国雑誌');
  if (!sh) {
    韓国雑誌_安全alert_('韓国雑誌シートが見つかりません');
    return;
  }

  const 同期件数 = 韓国雑誌_共通マスターから同期_();

  const 最終行 = Math.max(sh.getLastRow() + 100, 200);
  const 列 = 韓国雑誌_ヘッダーMap_(sh);

  const 雑誌マスター = ss.getSheetByName(設定_韓国雑誌.雑誌マスター名);
  if (雑誌マスター && 雑誌マスター.getLastRow() >= 2 && 列['雑誌名']) {
    const 雑誌値 = 雑誌マスター.getRange(2, 1, 雑誌マスター.getLastRow() - 1, 1)
      .getValues().flat()
      .map(v => String(v || '').trim())
      .filter(v => v);

    if (雑誌値.length > 0) {
      sh.getRange(2, 列['雑誌名'], 最終行 - 1, 1).setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(雑誌値, true)
          .build()
      );
    }
  }

  if (列['登録状況']) {
    sh.getRange(2, 列['登録状況'], 最終行 - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['未登録', '登録済み', '売り切れ'], true)
        .build()
    );
  }

  if (列['発番発行']) {
    sh.getRange(2, 列['発番発行'], 最終行 - 1, 1).insertCheckboxes();
  }

  const masterSS = 韓国雑誌_共通SS_();
  if (列['配送パターン']) {
    const 配送マスター = masterSS.getSheetByName(設定_韓国雑誌.配送パターンマスター名);
    if (配送マスター && 配送マスター.getLastRow() >= 2) {
      const 配送値 = 配送マスター.getRange(2, 1, 配送マスター.getLastRow() - 1, 1)
        .getValues().flat()
        .map(v => String(v || '').trim())
        .filter(v => v);

      if (配送値.length > 0) {
        sh.getRange(2, 列['配送パターン'], 最終行 - 1, 1).setDataValidation(
          SpreadsheetApp.newDataValidation()
            .requireValueInList(配送値, true)
            .build()
        );
      }
    }
  }

  if (雑誌マスター && 雑誌マスター.getLastRow() >= 2) {
    雑誌マスター.getRange(2, 4, 雑誌マスター.getLastRow() - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['年月型', '号数型'], true)
        .build()
    );
  }

  if (列['親コード']) {
    const 親コード列 = 列['親コード'];
    const colLetter = 韓国雑誌_列番号を英字へ_(親コード列);

    const rules = sh.getConditionalFormatRules().filter(rule =>
      !rule.getRanges().some(rng =>
        rng.getColumn() === 1 && rng.getLastColumn() === sh.getLastColumn()
      )
    );

    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND($${colLetter}2<>"",COUNTIF($${colLetter}:$${colLetter},$${colLetter}2)>1)`)
        .setBackground('#f4cccc')
        .setFontColor('#cc0000')
        .setRanges([sh.getRange(2, 1, 最終行 - 1, sh.getLastColumn())])
        .build()
    );

    sh.setConditionalFormatRules(rules);
  }

  韓国雑誌_安全toast_(
    `✅ 韓国雑誌 プルダウン更新完了（韓国マスター同期 ${同期件数}件）`,
    '完了',
    3
  );
}