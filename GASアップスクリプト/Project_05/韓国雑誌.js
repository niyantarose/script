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
  商品コード:       '商品コード(SKU)',
  商品名出品用:     '商品名（出品用）',
  言語:             '言語',
  雑誌名:           '雑誌名',
  年:               '年',
  月:               '月',
  号数:             '号数',
  バリエーション:   'バリエーションコード',
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

/**
 * バリエーションコードの正規化。
 * うちのコード体系はハイフン・アンダースコア等の区切りを入れない
 * （例: MARI2607A）ので、英数字以外を全部落とす。大文字小文字は保持。
 */
function 韓国雑誌_バリエーション正規化_(v) {
  return String(v || '').trim().replace(/[^A-Za-z0-9]/g, '');
}

/**
 * シート上のバリエーションコード列の実ヘッダー名を返す。
 * 設定の列名が見つからなければ「バリエーション」を含むヘッダーを探す。
 */
function 韓国雑誌_バリエーション列名_(列マップ) {
  const 設定名 = 設定_韓国雑誌.列名.バリエーション;
  if (設定名 && 列マップ[設定名]) return 設定名;
  const hit = Object.keys(列マップ).find(k => k.indexOf('バリエーション') !== -1);
  return hit || '';
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
    '発番発行', '登録状況', '雑誌名', '年', '月', '号数', 'バリエーションコード', '表紙情報', '特典メモ',
    '商品コード(SKU)', '商品名（出品用）', '粗利益率', '登録日',
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
    'バリエーションコード': '#6aa84f',
    '表紙情報': '#f1c232',
    '特典メモ': '#f1c232',
    '商品コード(SKU)': '#999999',
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
    'バリエーションコード': 110,
    '表紙情報': 200, '特典メモ': 240,
    '商品コード(SKU)': 140, '商品名（出品用）': 360, '粗利益率': 90, '登録日': 120,
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

function 韓国雑誌_商品コードSKUを生成_(雑誌名, 年, 月, 号数, バリエーション) {
  const info = 韓国雑誌_マスターを検索_(雑誌名);
  if (!info || !info.略称) return 'ERROR:マスター未登録';

  // 表紙違い・セット等の枝番は区切りなしで末尾に連結する（例: MARI2607A）
  const 枝番 = 韓国雑誌_バリエーション正規化_(バリエーション);

  if (info.コード型 === '号数型') {
    if (!号数) return 'ERROR:号数未入力';
    return `${info.略称}${String(号数).trim()}${枝番}`;
  }

  if (!年 || !月) return 'ERROR:年月未入力';
  return `${info.略称}${String(年).slice(-2)}${String(月).padStart(2, '0')}${枝番}`;
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

function 韓国雑誌_行を再計算_(sh, row, opts) {
  if (!sh || sh.getName() !== 設定_韓国雑誌.マスターシート名 || row < 2) return;

  // opts.列マップ / opts.rowValues を渡すと、複数行処理時に
  // ヘッダーやセル値の再読み込みを省略できる（読み取りのみ。書き込みはセル単位）
  const 列 = (opts && opts.列マップ) || 韓国雑誌_ヘッダーMap_(sh);
  const rowValues = opts && opts.rowValues;
  const get = (名前) => {
    if (!列[名前]) return '';
    if (rowValues) return rowValues[列[名前] - 1];
    return sh.getRange(row, 列[名前]).getValue();
  };
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

  const バリエーション列 = 韓国雑誌_バリエーション列名_(列);
  const バリエーション = バリエーション列 ? get(バリエーション列) : '';

  const 商品コードSKU = 韓国雑誌_商品コードSKUを生成_(雑誌名, 年, 月, 号数, バリエーション);

  // ダニエル取得で確定したコード（Set等の生成できない特殊コード含む）は
  // 生成コードで上書きしない。SKUセルのノートが目印。
  const SKU列番号 = 列['商品コード(SKU)'];
  const ダニエル固定 = SKU列番号
    ? String(sh.getRange(row, SKU列番号).getNote() || '').indexOf('ダニエル取得') === 0
    : false;
  if (!ダニエル固定) set('商品コード(SKU)', 商品コードSKU);

  const 商品名 = 韓国雑誌_出品用商品名を生成_(雑誌名, 年, 月, 号数, 表紙情報, 特典メモ);
  if (商品名) set('商品名（出品用）', 商品名);

  if (雑誌名 && !get('登録日')) set('登録日', new Date());
}

/* ============================================================
 * onEdit
 * ============================================================ */

function 韓国雑誌_onEdit(e) {
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  if (sh.getName() !== 設定_韓国雑誌.マスターシート名) return;

  // 複数行・複数列の貼り付けにも対応（台湾雑誌と同じ方式）
  const 開始行 = e.range.getRow();
  const 行数 = e.range.getNumRows();
  if (開始行 + 行数 - 1 < 2) return;
  if (typeof 自己更新中か_ === 'function' && 自己更新中か_()) return;

  const 列 = 韓国雑誌_ヘッダーMap_(sh);
  const 監視列 = ['雑誌名', '年', '月', '号数', '表紙情報', '特典メモ'];
  const バリエーション列 = 韓国雑誌_バリエーション列名_(列);
  if (バリエーション列) 監視列.push(バリエーション列);
  const 監視列番号 = 監視列.map(名前 => 列[名前]).filter(Boolean);

  const 編集開始列 = e.range.getColumn();
  const 編集終了列 = e.range.getLastColumn();
  if (!監視列番号.some(c => c >= 編集開始列 && c <= 編集終了列)) return;

  const 雑誌名列 = 列['雑誌名'];
  const 雑誌名を編集した =
    !!雑誌名列 && 雑誌名列 >= 編集開始列 && 雑誌名列 <= 編集終了列;

  const lock = LockService.getDocumentLock();
  try {
    if (!lock.tryLock(5000)) return;
  } catch (_) {
    return;
  }

  try {
    if (typeof 自己更新を開始_ === 'function') 自己更新を開始_();

    // 編集範囲の値を一括で読み、1行ずつ再計算する
    const lastCol = sh.getLastColumn();
    const 範囲値 = sh.getRange(開始行, 1, 行数, lastCol).getValues();

    for (let r = 開始行; r < 開始行 + 行数; r++) {
      if (r < 2) continue;

      韓国雑誌_行を再計算_(sh, r, {
        列マップ: 列,
        rowValues: 範囲値[r - 開始行]
      });

      if (雑誌名を編集した) {
        韓国雑誌_未解決候補を自動追加_(sh, r);
      }
    }
  } finally {
    if (typeof 自己更新を終了_ === 'function') 自己更新を終了_();
    try { lock.releaseLock(); } catch (_) {}
  }
}

/* ============================================================
 * 確定発行
 * ============================================================ */

function 韓国雑誌_重複キーを作成_(雑誌名, 年, 月, 号数, バリエーション) {
  const normalizedName = 韓国雑誌_雑誌名を正規化_(雑誌名);
  if (!normalizedName) return '';
  // 表紙違い・セットは別商品なので、バリエーションコードもキーに含める
  const 枝番 = 韓国雑誌_バリエーション正規化_(バリエーション).toUpperCase();
  const 枝 = 枝番 ? `-${枝番}` : '';
  if (号数) return `${normalizedName}-${号数}${枝}`;
  if (年 && 月) return `${normalizedName}-${年}-${月}${枝}`;
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
  const バリ列名 = 韓国雑誌_バリエーション列名_(列);
  const 重複チェックMap = {};
  全データ.forEach((row, i) => {
    const key = 韓国雑誌_重複キーを作成_(
      String(get行(row, '雑誌名') || '').trim(),
      String(get行(row, '年') || '').trim(),
      String(get行(row, '月') || '').trim(),
      String(get行(row, '号数') || '').trim(),
      バリ列名 ? String(get行(row, バリ列名) || '').trim() : ''
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
      String(get行(row, '号数') || '').trim(),
      バリ列名 ? String(get行(row, バリ列名) || '').trim() : ''
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
      const 商品コードSKU = String(get行(row, '商品コード(SKU)') || '').trim();
if (!商品コードSKU || 商品コードSKU.startsWith('ERROR')) return;
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

  const 商品コードSKU = get('商品コード(SKU)');
  if (!String(商品コードSKU).startsWith('ERROR:マスター未登録')) return false;

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
 * 候補 → 共通マスター反映（正本を更新する台湾雑誌と同じ方式）
 *
 * 以前はローカルの「雑誌マスター（韓国）」へ直接書き込んでいたが、
 * ローカルマスターはプルダウン更新のたびに共通マスターから
 * clearContents() で作り直される同期コピーのため、反映内容が
 * 次の同期で消えてしまっていた。正本（雑誌マスター（共通））側を
 * ライブラリ経由で更新する方式に統一。
 * ============================================================ */

function 韓国雑誌_候補を正式マスターへ反映() {
  const ui = SpreadsheetApp.getUi();
  if (
    ui.alert(
      '確認',
      '雑誌マスター候補（共通）で「反映」にチェックした行を、共通ファイルの雑誌マスター（共通）へ反映します。\n' +
      '（マスター共通ファイル側を更新します。このファイルの「雑誌マスター（韓国）」へは、続けてプルダウン更新を実行すると同期されます）\n\n' +
      '続行しますか？',
      ui.ButtonSet.OK_CANCEL
    ) !== ui.Button.OK
  ) {
    return;
  }

  const r = _kyoutuu.共通雑誌候補をマスターへ反映(韓国雑誌_共通SS_());
  ui.alert(
    '完了',
    (r
      ? `新規追加: ${r.追加}件\n別名として追加: ${r.別名追加}件\n既存一致: ${r.一致}件`
      : '候補データがありませんでした（または反映チェックなし）') +
    '\n\n続けて「韓国雑誌_プルダウン更新」を実行すると、雑誌マスター（韓国）へ同期されます。',
    ui.ButtonSet.OK
  );
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

  if (列['商品コード(SKU)']) {
    const 商品コードSKU列 = 列['商品コード(SKU)'];
const colLetter = 韓国雑誌_列番号を英字へ_(商品コードSKU列);

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

/* ============================================================
 * ダニエル(kstargate)商品コード取得
 *
 * 韓国雑誌の商品コードはダニエルの通販サイト kstargate.com の
 * 採番（商品詳細ページの「kstargate ID」）をそのまま使う。
 *
 * 取得の流れ:
 *   ① 略称+年月（例 MARI2607）/ 略称+号数（例 CIN211564）の
 *      ベースコードでサイト内検索（説明文のIDに部分一致するので
 *      表紙違い・セットも全部ヒットする）
 *   ② ヒットしない場合は 雑誌名(英字)+年 で検索し、商品名の
 *      年月・号数で絞り込む
 *   ③ 詳細ページから kstargate ID を読み取り、行のバリエーション
 *      コードと突き合わせて1件に決まったら書き込む
 *
 * 書き込み時はうちのコード体系に合わせて区切り文字を除去する
 * （AKOR2607_D → AKOR2607D）。取得したセルにはノートを付け、
 * onEditの再計算が生成コードで上書きしないよう保護する。
 * ============================================================ */

const ダニエル設定 = {
  ベースURL: 'https://www.kstargate.com',
  待機ミリ秒: 350,   // 仕入先サイトなので必ず間隔をあける
  最大フェッチ数: 100, // 1回の実行でのHTTPアクセス上限
  行上限: 20,          // 1回の実行で処理する行数上限
  候補詳細上限: 6      // 1行あたり詳細ページを見る候補数上限
};

let ダニエル_fetch回数_ = 0;

function ダニエル_コード正規化_(v) {
  return String(v || '').trim().replace(/[^A-Za-z0-9]/g, '');
}

function ダニエル_HTML取得_(url) {
  if (ダニエル_fetch回数_ >= ダニエル設定.最大フェッチ数) {
    throw new Error('FETCH_LIMIT');
  }
  ダニエル_fetch回数_++;
  Utilities.sleep(ダニエル設定.待機ミリ秒);

  const resp = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' }
  });
  if (resp.getResponseCode() !== 200) return '';
  // kstargateはEUC-JP
  return resp.getContentText('EUC-JP');
}

/** サイト内検索。結果一覧から brandcode(12桁) と商品名のペアを返す */
function ダニエル_検索_(キーワード) {
  const url = ダニエル設定.ベースURL + '/shop/shopbrand.html?search=' + encodeURIComponent(キーワード);
  const html = ダニエル_HTML取得_(url);
  if (!html) return [];

  const seen = {};
  const list = [];
  const re = /shopdetail\.html\?brandcode=(\d+)[^>]*>([^<]+)</g;
  let m;
  while ((m = re.exec(html)) !== null) {
    const 名前 = m[2].replace(/\s+/g, ' ').trim();
    if (!名前) continue;           // 画像リンク側はテキストが無いのでスキップ
    if (seen[m[1]]) continue;
    seen[m[1]] = true;
    list.push({ brandcode: m[1], 商品名: 名前 });
  }
  return list;
}

/** 商品詳細ページから「kstargate ID」を読み取る */
function ダニエル_商品コードを取得_(brandcode) {
  const html = ダニエル_HTML取得_(ダニエル設定.ベースURL + '/shopdetail/' + brandcode + '/');
  if (!html) return '';
  let m = html.match(/kstargate ID<\/strong>(?:&#160;|&nbsp;|\s)*([A-Za-z0-9_\-]+)/i);
  if (!m) m = html.match(/kstargate ID(?:&(?:amp;)?#160;|&nbsp;|\s)*([A-Za-z0-9_\-]+)/i);
  return m ? m[1] : '';
}

/** 候補（検索結果）の詳細ページを見て kstargate ID を集める */
function ダニエル_候補詳細_(候補) {
  const out = [];
  候補.slice(0, ダニエル設定.候補詳細上限).forEach(c => {
    const code = ダニエル_商品コードを取得_(c.brandcode);
    if (code) out.push({ brandcode: c.brandcode, 商品名: c.商品名, コード: code });
  });
  return out;
}

function ダニエル_枝番_(コード, ベース) {
  if (!ベース) return '';
  const n = ダニエル_コード正規化_(コード);
  if (n.toUpperCase().indexOf(ベース.toUpperCase()) !== 0) return '';
  return n.slice(ベース.length);
}

/**
 * 詳細リストの中から行に対応する1件を決める。
 * - バリエーション込みの完全一致が最優先
 * - バリエーション未入力なら、ベース前方一致が1件だけのとき採用
 * - 前方一致が複数なら「複数」として人に返す
 */
function ダニエル_詳細からマッチ_(詳細, ベース, 期待, バリエーション) {
  if (詳細.length === 0) return null;

  const norm = c => ダニエル_コード正規化_(c).toUpperCase();

  if (期待) {
    const exact = 詳細.filter(d => norm(d.コード) === 期待);
    if (exact.length === 1) {
      return {
        状態: '確定',
        コード: exact[0].コード,
        brandcode: exact[0].brandcode,
        枝番: ダニエル_枝番_(exact[0].コード, ベース)
      };
    }
  }

  if (ベース) {
    const pre = 詳細.filter(d => norm(d.コード).indexOf(ベース.toUpperCase()) === 0);
    if (pre.length === 1 && !バリエーション) {
      return {
        状態: '確定',
        コード: pre[0].コード,
        brandcode: pre[0].brandcode,
        枝番: ダニエル_枝番_(pre[0].コード, ベース)
      };
    }
    if (pre.length > 1) return { 状態: '複数', 候補: pre };
  }

  return null;
}

/** 1行分の探索本体 */
function ダニエル_行のコードを探す_(雑誌名, 年, 月, 号数, バリエーション) {
  // ベースコード（略称+号数 / 略称+年下2桁+月2桁）
  const info = 韓国雑誌_マスターを検索_(雑誌名);
  let ベース = '';
  if (info && info.略称) {
    if (info.コード型 === '号数型') {
      if (号数) ベース = info.略称 + String(号数).trim();
    } else if (年 && 月) {
      ベース = info.略称 + String(年).slice(-2) + String(月).padStart(2, '0');
    }
  }
  const 期待 = ベース ? (ベース + バリエーション).toUpperCase() : '';

  // ① ベースコードで直接検索
  if (ベース) {
    const 候補 = ダニエル_検索_(ベース);
    if (候補.length > 0) {
      const 詳細 = ダニエル_候補詳細_(候補);
      const r = ダニエル_詳細からマッチ_(詳細, ベース, 期待, バリエーション);
      if (r) return r;
    }
  }

  // ② 雑誌名(英字)+年（または号数）で検索して商品名で絞り込む
  const kw = (雑誌名 + ' ' + (号数 ? 号数 : (年 || '')))
    .replace(/[^\x20-\x7E]/g, ' ')  // サイトがEUC-JPなので検索語はASCIIに限定する
    .replace(/\s+/g, ' ')
    .trim();
  if (!kw) return { 状態: '不明', 理由: '検索キーワードを作れません' };

  const 候補 = ダニエル_検索_(kw);
  if (候補.length === 0) return { 状態: '不明', 理由: `検索ヒットなし (${kw})` };

  let 絞り込み = 候補;
  if (号数) {
    絞り込み = 候補.filter(c => c.商品名.indexOf(String(号数)) !== -1);
  } else if (年 && 月) {
    絞り込み = 候補.filter(c =>
      c.商品名.indexOf(`${年}年`) !== -1 &&
      c.商品名.indexOf(`${parseInt(月, 10)}月`) !== -1
    );
  }
  if (絞り込み.length === 0) 絞り込み = 候補;

  const 詳細 = ダニエル_候補詳細_(絞り込み);
  const r = ダニエル_詳細からマッチ_(詳細, ベース, 期待, バリエーション);
  if (r) return r;

  if (詳細.length === 1) {
    return { 状態: '確定', コード: 詳細[0].コード, brandcode: 詳細[0].brandcode, 枝番: '' };
  }
  if (詳細.length > 1) return { 状態: '複数', 候補: 詳細 };
  return { 状態: '不明', 理由: 'kstargate IDを読み取れませんでした' };
}

/** メニュー本体：未登録・未取得の行にダニエルの商品コードを記入する */
function 韓国雑誌_ダニエル商品コード取得() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(設定_韓国雑誌.マスターシート名);
  if (!sh || sh.getLastRow() < 2) { 韓国雑誌_安全alert_('データがありません'); return; }

  const 列 = 韓国雑誌_ヘッダーMap_(sh);
  const SKU列 = 列['商品コード(SKU)'];
  if (!列['雑誌名'] || !SKU列) {
    韓国雑誌_安全alert_('「雑誌名」「商品コード(SKU)」列が見つかりません');
    return;
  }

  const バリ列名 = 韓国雑誌_バリエーション列名_(列);
  const 最終行 = sh.getLastRow();
  const 全データ = sh.getRange(2, 1, 最終行 - 1, sh.getLastColumn()).getValues();
  const SKUノート = sh.getRange(2, SKU列, 最終行 - 1, 1).getNotes();

  const get行 = (row, 名前) =>
    列[名前] ? String(row[列[名前] - 1] == null ? '' : row[列[名前] - 1]).trim() : '';

  const 対象 = [];
  全データ.forEach((row, i) => {
    if (!get行(row, '雑誌名')) return;
    // 年月か号数が入っていない行は特定できないので対象外
    const 年月あり = get行(row, '年') && get行(row, '月');
    if (!年月あり && !get行(row, '号数')) return;
    if (get行(row, '登録状況').indexOf('登録済') === 0) return;
    if (String(SKUノート[i][0] || '').indexOf('ダニエル取得') === 0) return; // 取得済み
    対象.push({ row, rowNum: i + 2 });
  });

  if (対象.length === 0) {
    韓国雑誌_安全alert_('対象行がありません\n（未登録で、まだダニエル取得していない行が対象です）');
    return;
  }

  const 今回 = 対象.slice(0, ダニエル設定.行上限);
  const 確認 = ui.alert(
    '🚚 ダニエル商品コード取得',
    `対象 ${対象.length}行のうち ${今回.length}行を処理します。` +
    (対象.length > 今回.length ? `\n（1回の実行は${ダニエル設定.行上限}行まで。残りは再実行してください）` : '') +
    '\n\nkstargate.com に接続します。続行しますか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (確認 !== ui.Button.OK) return;

  ダニエル_fetch回数_ = 0;
  let 確定数 = 0, 複数数 = 0, 不明数 = 0;
  const レポート = [];

  try {
    今回.forEach(({ row, rowNum }, idx) => {
      ss.toast(`${idx + 1}/${今回.length}件目（${rowNum}行目）を照会中...`, '🚚 ダニエル取得', 5);

      const 雑誌名 = 韓国雑誌_雑誌名を正規化_(get行(row, '雑誌名'));
      const 年 = get行(row, '年');
      const 月 = get行(row, '月');
      const 号数 = get行(row, '号数');
      const バリエーション = 韓国雑誌_バリエーション正規化_(バリ列名 ? get行(row, バリ列名) : '');

      const 結果 = ダニエル_行のコードを探す_(雑誌名, 年, 月, 号数, バリエーション);
      const cell = sh.getRange(rowNum, SKU列);

      if (結果.状態 === '確定') {
        cell.setValue(ダニエル_コード正規化_(結果.コード));
        cell.setNote(
          'ダニエル取得: ' + 結果.コード + '\n' +
          ダニエル設定.ベースURL + '/shopdetail/' + 結果.brandcode + '/\n' +
          '雑誌名・年月・バリエーションを変えた場合はこのノートを消して再取得してください'
        );
        // バリエーションが空でサイト側に枝番があれば逆記入しておく
        if (バリ列名 && !バリエーション && 結果.枝番) {
          sh.getRange(rowNum, 列[バリ列名]).setValue(結果.枝番);
        }
        確定数++;
      } else if (結果.状態 === '複数') {
        cell.setNote(
          '候補(ダニエル): 以下から選んで商品コードを記入してください\n' +
          結果.候補.map(c => ダニエル_コード正規化_(c.コード) + ' … ' + c.商品名).join('\n')
        );
        レポート.push(`${rowNum}行目: 候補複数 → ${結果.候補.map(c => ダニエル_コード正規化_(c.コード)).join(', ')}`);
        複数数++;
      } else {
        レポート.push(`${rowNum}行目: 見つからず${結果.理由 ? '（' + 結果.理由 + '）' : ''}`);
        不明数++;
      }
    });
  } catch (e) {
    if (String(e.message) !== 'FETCH_LIMIT') throw e;
    レポート.push('⚠ 今回のアクセス上限に達したため途中で停止しました。残りは再実行してください。');
  }

  韓国雑誌_安全alert_(
    `✅ ダニエル商品コード取得 完了\n` +
    `確定: ${確定数}件 / 候補複数: ${複数数}件 / 見つからず: ${不明数}件` +
    (レポート.length ? '\n\n' + レポート.join('\n') : '') +
    '\n\n※ 候補複数の行は商品コード(SKU)セルのノートに候補一覧を書いてあります'
  );
}