/**
 * 台湾雑誌.gs
 * 台湾雑誌シート専用の設定・シート作成・onEdit・確定発行
 *
 * ★ コード体系
 *   年月型: {言語接頭辞}-{略称}{年下2桁}{月2桁}   例: TW-MAXI2603
 *   号数型: {言語接頭辞}-{略称}{号数}              例: TW-1STLOOK127
 *   ※ SP/A/B 等は原題タイトルから自動付与          例: TW-BAZA2510SP-A
 *
 * 【スナップショット運用】
 * プルダウン更新時に共通マスターの内容を
 * ローカルの隠しシートにコピーする
 *   _SNAP_雑誌マスター  … 雑誌マスター（共通）のコピー
 *   _SNAP_版種ルール    … 版種ルール（雑誌共通）のコピー
 *   _SNAP_タイプルール  … タイプルール（雑誌共通）のコピー
 * onEditはこのローカルコピーを参照するため高速動作する
 */

const MASTER_SPREADSHEET_ID = '1ZAy5wtQCq1ixl47MMEIGnrVWeae-F_NfN34lSOUsN5M';

const 設定_台湾雑誌 = {
  マスターシート名: '台湾雑誌',
  雑誌マスター名:   '雑誌マスター（共通）',
  版種ルール名:     '版種ルール（雑誌共通）',
  タイプルール名:   'タイプルール（雑誌共通）',
　候補シート名:     '雑誌マスター候補（共通）',  // ← 追加

  // スナップショットシート名（隠しシート）
  スナップ雑誌マスター名: '_SNAP_雑誌マスター',
  スナップ版種ルール名:   '_SNAP_版種ルール',
  スナップタイプルール名: '_SNAP_タイプルール',

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
    TW: '台湾版', CN: '中国版', HK: '香港版', TH: 'タイ版'
  },

  列名: {
    発番発行: '発番発行', 登録状況: '登録状況', 言語: '言語',
    雑誌名: '雑誌名', 年: '年', 月: '月', 号数: '号数',
    表紙情報: '表紙情報', 特典メモ: '特典メモ',
    親コード: '親コード', 商品名出品用: '商品名（出品用）',
    粗利益率: '粗利益率', 登録日: '登録日',
    売価: '売価', 配送パターン: '配送パターン',
    登録者: '登録者', 商品説明: '商品説明',
    原価: '原価', 原題タイトル: '原題タイトル', 原題商品名: '原題商品名',
    博客來商品コード: '博客來商品コード', 博客來URL: '博客來URL',
    メイン画像URL: 'メイン画像URL', 追加画像URL: '追加画像URL'
  }
};

let _台湾雑誌GS_共通SSキャッシュ = null;

function 台湾雑誌GS_共通SS_() {
  if (_台湾雑誌GS_共通SSキャッシュ) return _台湾雑誌GS_共通SSキャッシュ;
  _台湾雑誌GS_共通SSキャッシュ = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  return _台湾雑誌GS_共通SSキャッシュ;
}

/* ============================================================
 * スナップショット管理
 * ============================================================ */

/**
 * 共通マスターの内容をローカルにコピーする
 * プルダウン更新時に呼ぶ
 */
function 台湾雑誌_スナップショットを更新_() {
  const localSS  = SpreadsheetApp.getActive();
  const masterSS = 台湾雑誌GS_共通SS_();

  const コピー設定 = [
    { src: 設定_台湾雑誌.雑誌マスター名,  dst: 設定_台湾雑誌.スナップ雑誌マスター名  },
    { src: 設定_台湾雑誌.版種ルール名,    dst: 設定_台湾雑誌.スナップ版種ルール名    },
    { src: 設定_台湾雑誌.タイプルール名,  dst: 設定_台湾雑誌.スナップタイプルール名  },
  ];

  コピー設定.forEach(({ src, dst }) => {
    const srcSh = masterSS.getSheetByName(src);
    if (!srcSh || srcSh.getLastRow() < 1) return;

    let dstSh = localSS.getSheetByName(dst);
    if (!dstSh) {
      dstSh = localSS.insertSheet(dst);
      dstSh.hideSheet();
    } else {
      dstSh.clearContents();
    }

    const lastRow = srcSh.getLastRow();
    const lastCol = srcSh.getLastColumn();
    if (lastRow < 1 || lastCol < 1) return;

    const data = srcSh.getRange(1, 1, lastRow, lastCol).getValues();
    dstSh.getRange(1, 1, lastRow, lastCol).setValues(data);
  });

  Logger.log('スナップショット更新完了');
}

/**
 * ローカルスナップショットを取得（なければ共通SSを開く）
 */
function 台湾雑誌GS_スナップシートを取得_(スナップ名, 共通名) {
  const localSS = SpreadsheetApp.getActive();
  const snap = localSS.getSheetByName(スナップ名);
  if (snap && snap.getLastRow() >= 2) return snap;
  // スナップがなければ共通SSから直接取得
  return 台湾雑誌GS_共通SS_().getSheetByName(共通名);
}

/* ============================================================
 * セットアップ（初回1回実行）
 * ============================================================ */

function 台湾雑誌_雑誌コード自動生成セットアップ() {
  台湾雑誌_マスター拡張を適用();
  台湾雑誌_版種ルールシートを作成_();
  台湾雑誌_タイプルールシートを作成_();
  台湾雑誌_プルダウン更新();
  SpreadsheetApp.getActive().toast('✅ 台湾雑誌 コード自動生成セットアップ完了', '完了', 5);
}

function 台湾雑誌_マスター拡張を適用() {
  const ss = 台湾雑誌GS_共通SS_();
  let sh = ss.getSheetByName(設定_台湾雑誌.雑誌マスター名);

  if (!sh) {
    sh = ss.insertSheet(設定_台湾雑誌.雑誌マスター名);
    sh.getRange(1, 1, 1, 7).setValues([[
      '雑誌名（英字）', '雑誌名（カタカナ）', '略称コード',
      '基本キー型', '通常タイプ結合記号', '版種ありタイプ結合記号', '備考'
    ]]);
    sh.getRange(2, 1, 2, 7).setValues([
      ['MAXIM KOREA', 'マキシム・コリア', 'MAXI', '年月型', '', '-', ''],
      ['GQ KOREA',    'ジーキュー・コリア', 'GQ', '年月型', '', '-', ''],
    ]);
  } else {
    const lastCol = Math.max(sh.getLastColumn(), 1);
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());

    const ensureCol = (name) => {
      let idx = headers.indexOf(name);
      if (idx >= 0) return idx + 1;
      sh.insertColumnAfter(sh.getLastColumn());
      const newCol = sh.getLastColumn();
      sh.getRange(1, newCol).setValue(name);
      headers.push(name);
      return newCol;
    };

    const oldIdx = headers.indexOf('コード型');
    if (oldIdx >= 0 && headers.indexOf('基本キー型') < 0) {
      sh.getRange(1, oldIdx + 1).setValue('基本キー型');
      headers[oldIdx] = '基本キー型';
    }
    if (headers.indexOf('基本キー型') < 0) ensureCol('基本キー型');

    const colNormal  = ensureCol('通常タイプ結合記号');
    const colSpecial = ensureCol('版種ありタイプ結合記号');
    ensureCol('備考');

    const lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      const specialVals = sh.getRange(2, colSpecial, lastRow - 1, 1).getValues();
      specialVals.forEach((r, i) => { if (r[0] === '' || r[0] == null) specialVals[i][0] = '-'; });
      sh.getRange(2, colSpecial, lastRow - 1, 1).setValues(specialVals);
    }
  }

  const map = 台湾雑誌GS_ヘッダーMap_(sh);
  if (map['基本キー型'] && sh.getLastRow() >= 2) {
    sh.getRange(2, map['基本キー型'], sh.getLastRow() - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['年月型', '号数型'], true).build()
    );
  }

  sh.getRange(1, 1, 1, sh.getLastColumn())
    .setBackground('#6aa84f').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center');
  sh.setFrozenRows(1);

  SpreadsheetApp.getActive().toast('✅ 雑誌マスター（共通）を拡張しました', '完了', 3);
}

function 台湾雑誌_版種ルールシートを作成_() {
  台湾雑誌GS_ルールシートを確保_(設定_台湾雑誌.版種ルール名, [
    ['SPECIAL EDITION', 'SP', 'SPECIAL EDITION', 100, true],
    ['SPECIAL',         'SP', 'SPECIAL EDITION',  90, true],
    ['超值版',           '',  '超値版',            80, true],
    ['特別版',           'SP', '特別版',           70, true],
    ['限定版',           'LE', '限定版',           60, true],
  ]);
}

function 台湾雑誌_タイプルールシートを作成_() {
  台湾雑誌GS_ルールシートを確保_(設定_台湾雑誌.タイプルール名, [
    ['Aタイプ', 'A', 'Aタイプ', 100, true],
    ['Bタイプ', 'B', 'Bタイプ', 100, true],
    ['Cタイプ', 'C', 'Cタイプ', 100, true],
    ['TYPE A',  'A', 'Aタイプ',  90, true],
    ['TYPE B',  'B', 'Bタイプ',  90, true],
    ['TYPE C',  'C', 'Cタイプ',  90, true],
    ['VER.A',   'A', 'Aタイプ',  80, true],
    ['VER.B',   'B', 'Bタイプ',  80, true],
    ['VER.C',   'C', 'Cタイプ',  80, true],
  ]);
}

function 台湾雑誌GS_ルールシートを確保_(シート名, rows) {
  const ss = 台湾雑誌GS_共通SS_();
  let sh = ss.getSheetByName(シート名);
  if (sh) return sh;

  sh = ss.insertSheet(シート名);
  const headers = ['判定キーワード', 'コードsuffix', 'タイトル表示', '優先順位', '有効'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows && rows.length) {
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sh.getRange(2, 5, rows.length, 1).insertCheckboxes();
  }
  sh.getRange(1, 1, 1, headers.length)
    .setBackground('#6aa84f').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  return sh;
}

/* ============================================================
 * 雑誌マスター（共通）作成（新規）
 * ============================================================ */

function 台湾雑誌マスターを作成() {
  const ui = SpreadsheetApp.getUi();
  const ss = 台湾雑誌GS_共通SS_();
  if (ss.getSheetByName(設定_台湾雑誌.雑誌マスター名)) {
    ui.alert('「雑誌マスター（共通）」シートは既に存在します\n\n拡張する場合は「コード自動生成セットアップ」を実行してください');
    return;
  }
  台湾雑誌_マスター拡張を適用();
  ui.alert('✅ 雑誌マスター（共通）を作成しました\n\n例: TW-BAZA2510A / TW-BAZA2510SP-A');
}

/* ============================================================
 * 台湾雑誌シート作成
 * ============================================================ */

function 台湾雑誌シートを作成() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const シート名 = '台湾雑誌';
  if (ss.getSheetByName(シート名)) { ui.alert(`「${シート名}」シートは既に存在します`); return; }

  const sh = ss.insertSheet(シート名);
  const ヘッダー = [
    '発番発行', '登録状況', '言語', '雑誌名', '年', '月', '号数', '表紙情報', '特典メモ',
    '親コード', '商品名（出品用）', '粗利益率', '登録日',
    '売価', '配送パターン', '登録者', '商品説明',
    '原価', '原題タイトル', '原題商品名',
    '博客來商品コード', '博客來URL', 'メイン画像URL', '追加画像URL'
  ];

  sh.getRange(1, 1, 1, ヘッダー.length).setValues([ヘッダー]);

  const 色 = {
    '発番発行': '#cc0000', '登録状況': '#4a86e8',
    '言語': '#6aa84f', '雑誌名': '#6aa84f', '年': '#6aa84f', '月': '#6aa84f', '号数': '#6aa84f',
    '表紙情報': '#f1c232', '特典メモ': '#f1c232',
    '親コード': '#999999', '商品名（出品用）': '#999999', '粗利益率': '#999999', '登録日': '#999999',
    '売価': '#4a86e8', '配送パターン': '#4a86e8', '登録者': '#4a86e8', '商品説明': '#4a86e8',
    '原価': '#e69138', '原題タイトル': '#e69138', '原題商品名': '#e69138',
    '博客來商品コード': '#e69138', '博客來URL': '#e69138',
    'メイン画像URL': '#e69138', '追加画像URL': '#e69138',
  };

  for (let i = 0; i < ヘッダー.length; i++) {
    sh.getRange(1, i + 1)
      .setBackground(色[ヘッダー[i]] || '#cccccc')
      .setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  }

  sh.setRowHeight(1, 40);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);

  const 最終行 = 1000;
  sh.getRange(2, 1, 最終行 - 1, 1).insertCheckboxes();
  sh.getRange(2, 2, 最終行 - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['未登録', '登録済み', '売り切れ'], true).build()
  );
  sh.getRange(2, 3, 最終行 - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(設定_台湾雑誌.言語リスト, true).build()
  );

  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$B2="未登録"')
      .setBackground('#ffff00')
      .setRanges([sh.getRange(2, 1, 最終行 - 1, ヘッダー.length)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$B2="登録済み"')
      .setBackground('#d9ead3')
      .setRanges([sh.getRange(2, 1, 最終行 - 1, ヘッダー.length)]).build(),
  ]);

  const 列幅 = {
    '発番発行': 60, '登録状況': 80, '言語': 60, '雑誌名': 180, '年': 60, '月': 50, '号数': 70,
    '表紙情報': 200, '特典メモ': 240,
    '親コード': 150, '商品名（出品用）': 360, '粗利益率': 90, '登録日': 120,
    '売価': 80, '配送パターン': 100, '登録者': 80, '商品説明': 150,
    '原価': 80, '原題タイトル': 200, '原題商品名': 200,
    '博客來商品コード': 140, '博客來URL': 220, 'メイン画像URL': 200, '追加画像URL': 200
  };
  ヘッダー.forEach((h, i) => { if (列幅[h]) sh.setColumnWidth(i + 1, 列幅[h]); });

  ui.alert('✅ 台湾雑誌シートを作成しました\n\n【次の手順】\n① コード自動生成セットアップを実行\n② プルダウン更新を実行（スナップショットも同時に更新されます）');
}

/* ============================================================
 * 基本ヘルパー
 * ============================================================ */

function 台湾雑誌GS_ヘッダーMap_(sh) {
  const vals = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  const map = {};
  vals.forEach((h, i) => { if (h) map[String(h).trim()] = i + 1; });
  return map;
}

function 台湾雑誌GS_比較用文字列_(v) {
  return String(v || '').trim().replace(/[　\s]+/g, ' ').toUpperCase();
}

function 台湾雑誌GS_基本キー型を正規化_(raw) {
  const s = String(raw || '').trim().toUpperCase();
  if (s === 'ISSUE' || s === 'VOL' || s === '号数型' || s === '号数') return '号数型';
  return '年月型';
}

function 台湾雑誌GS_言語コードへ変換_(言語) {
  const raw = String(言語 || '').trim();
  if (!raw) return '';
  return 設定_台湾雑誌.言語コードマップ[raw] || 設定_台湾雑誌.言語コードマップ[raw.toUpperCase()] || '';
}

function 台湾雑誌GS_言語表示へ変換_(言語) {
  const code = 台湾雑誌GS_言語コードへ変換_(言語);
  return 設定_台湾雑誌.言語表示マップ[code] || String(言語 || '').trim();
}

/* ============================================================
 * マスター・ルール取得（スナップショット優先）
 * ============================================================ */

function 台湾雑誌GS_マスターを検索_(雑誌名, 言語 = '') {
  const sh = 台湾雑誌GS_スナップシートを取得_(
    設定_台湾雑誌.スナップ雑誌マスター名,
    設定_台湾雑誌.雑誌マスター名
  );
  if (!sh || sh.getLastRow() < 2) return null;

  const map = 台湾雑誌GS_ヘッダーMap_(sh);
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const target = 台湾雑誌GS_正規化雑誌文字列_(雑誌名);
  if (!target) return null;

  let best = null;
  let bestScore = -1;

  rows.forEach(row => {
    const info = 台湾雑誌GS_マスター行を情報化_(row, map);
    if (!info.aliases.length) return;

    info.aliases.forEach(alias => {
      const score = 台湾雑誌GS_別名一致スコア_(雑誌名, alias, info.languages, 言語);
      if (score > bestScore) {
        bestScore = score;
        best = info;
      }
    });
  });

  return best;
}

function 台湾雑誌GS_雑誌名を推定_(rawTitle, 言語 = '') {
  const sh = 台湾雑誌GS_スナップシートを取得_(
    設定_台湾雑誌.スナップ雑誌マスター名,
    設定_台湾雑誌.雑誌マスター名
  );
  if (!sh || sh.getLastRow() < 2) return '';

  const map = 台湾雑誌GS_ヘッダーMap_(sh);
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const raw = String(rawTitle || '').trim();
  if (!raw) return '';

  let best = null;
  let bestScore = -1;

  rows.forEach(row => {
    const info = 台湾雑誌GS_マスター行を情報化_(row, map);
    if (!info.aliases.length) return;

    info.aliases.forEach(alias => {
      const score = 台湾雑誌GS_別名一致スコア_(raw, alias, info.languages, 言語);
      if (score > bestScore) {
        bestScore = score;
        best = info;
      }
    });
  });

  return bestScore >= 120 && best ? best.表示英字名 : '';
}

function 台湾雑誌GS_雑誌マスター登録済みか_(雑誌名, 言語 = '') {
  return !!台湾雑誌GS_マスターを検索_(雑誌名, 言語);
}

function 台湾雑誌GS_マスター行を情報化_(row, map) {
  const aliases = [];
  const pushIf = value => {
    const text = String(value || '').trim();
    if (text && !aliases.includes(text)) aliases.push(text);
  };

  const 略称Col = (map['略称コード'] || map['略称'] || 1) - 1;
  const 英字Col = (map['雑誌名（英字）'] || map['英字名'] || 2) - 1;
  const カタカナCol = (map['雑誌名（カタカナ）'] || map['カタカナ名'] || 3) - 1;
  const 基本キー型Col = (map['基本キー型'] || map['コード型'] || 4) - 1;
  const 通常Col = map['通常タイプ結合記号'] ? map['通常タイプ結合記号'] - 1 : -1;
  const 版種Col = map['版種ありタイプ結合記号'] ? map['版種ありタイプ結合記号'] - 1 : -1;
  const 言語Col = (map['対応言語'] || 7) - 1;

  const 英字名 = String(row[英字Col] || '').trim();
  const カタカナ名 = String(row[カタカナCol] || '').trim();

  pushIf(英字名);
  Object.keys(map)
    .filter(name => /^別名\d*$/.test(name))
    .sort((a, b) => map[a] - map[b])
    .forEach(name => pushIf(row[map[name] - 1]));

  const 表示英字名 = 台湾雑誌GS_表示英字名を決定_(英字名, aliases);
  if (表示英字名 && !aliases.includes(表示英字名)) aliases.push(表示英字名);

  return {
    略称: String(row[略称Col] || '').trim(),
    英字名,
    カタカナ名,
    表示英字名,
    表示カタカナ名: 台湾雑誌GS_表示カタカナ名を決定_(カタカナ名),
    基本キー型: 台湾雑誌GS_基本キー型を正規化_(row[基本キー型Col]),
    通常タイプ結合記号: 通常Col >= 0 ? String(row[通常Col] || '').trim() : '',
    版種ありタイプ結合記号: 版種Col >= 0 ? String(row[版種Col] || '-').trim() : '-',
    languages: 台湾雑誌GS_対応言語一覧_(row[言語Col]),
    aliases,
  };
}

function 台湾雑誌GS_表示英字名を決定_(英字名, aliases) {
  const candidates = [];
  const seen = new Set();
  const pushCandidate = value => {
    const text = 台湾雑誌GS_英字地域表記を除去_(value);
    const key = 台湾雑誌GS_英字表示キー_(text);
    if (!text || !key || seen.has(key)) return;
    seen.add(key);
    candidates.push(text);
  };

  pushCandidate(英字名);
  (aliases || []).forEach(pushCandidate);

  if (!candidates.length) return String(英字名 || '').trim();

  const latinCandidates = candidates.filter(text => /[A-Za-z]/.test(text));
  const pool = latinCandidates.length ? latinCandidates : candidates;

  let best = pool[0];
  let bestScore = 台湾雑誌GS_表示英字名スコア_(best);

  pool.forEach(candidate => {
    const score = 台湾雑誌GS_表示英字名スコア_(candidate);
    if (score > bestScore || (score === bestScore && candidate.length < best.length)) {
      best = candidate;
      bestScore = score;
    }
  });

  return 台湾雑誌GS_英字表示名を整形_(best);
}

function 台湾雑誌GS_表示英字名スコア_(value) {
  const text = String(value || '').trim();
  if (!text) return -1;

  let score = 0;
  if (/[A-Za-z]/.test(text)) score += 100;
  if (/[A-Z]/.test(text) && /[a-z]/.test(text)) score += 25;
  else if (/^[A-Z0-9 '&+./-]+$/.test(text)) score += 20;
  else if (/^[a-z0-9 '&+./-]+$/.test(text)) score += 10;
  if (!/(?:KOREA|TAIWAN|HONG\s*KONG|HONGKONG|CHINA|THAILAND|HK|TW|CN|KR|TH)\b/i.test(text)) score += 40;
  score += Math.max(0, 40 - text.length);
  return score;
}

function 台湾雑誌GS_英字表示名を整形_(value) {
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
    'W': 'W',
  };
  return map[台湾雑誌GS_英字表示キー_(text)] || text;
}

function 台湾雑誌GS_英字表示キー_(value) {
  return String(value || '')
    .normalize('NFKC')
    .toUpperCase()
    .replace(/&/g, ' AND ')
    .replace(/[^A-Z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function 台湾雑誌GS_英字地域表記を除去_(value) {
  return String(value || '')
    .replace(/\b(?:KOREA|TAIWAN|HONG\s*KONG|HONGKONG|CHINA|THAILAND)\b/gi, ' ')
    .replace(/\b(?:HK|TW|CN|KR|TH)\b/gi, ' ')
    .replace(/(?:韓国版|韓國版|台湾版|臺灣版|中国版|中國版|香港版|タイ版)/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function 台湾雑誌GS_表示カタカナ名を決定_(カタカナ名) {
  return 台湾雑誌GS_カタカナ地域表記を除去_(カタカナ名);
}

function 台湾雑誌GS_カタカナ地域表記を除去_(value) {
  return String(value || '')
    .trim()
    .replace(/(?:・|\s)*(?:台湾版|台湾|臺灣版|臺灣|コリア|韓国版|韓国|韓國版|韓國|香港版|香港|中国版|中国|中國版|中國|タイ版|タイ)$/u, '')
    .replace(/[・\s]+$/u, '')
    .trim();
}

function 台湾雑誌GS_対応言語一覧_(raw) {
  return String(raw || '')
    .split(',')
    .map(v => 台湾雑誌GS_言語コードへ変換_(v))
    .filter(Boolean);
}

function 台湾雑誌GS_言語一致_(supported, language) {
  if (!supported || !supported.length) return true;
  const code = 台湾雑誌GS_言語コードへ変換_(language);
  return !code || supported.includes(code);
}

function 台湾雑誌GS_別名一致スコア_(rawTitle, alias, supportedLanguages, language) {
  const raw = String(rawTitle || '').trim();
  const aliasText = String(alias || '').trim();
  if (!raw || !aliasText) return -1;

  const rawNorm = 台湾雑誌GS_正規化雑誌文字列_(raw);
  const aliasNorm = 台湾雑誌GS_正規化雑誌文字列_(aliasText);
  if (!aliasNorm) return -1;

  let score = -1;
  if (rawNorm === aliasNorm) score = 260;
  else if (rawNorm.startsWith(aliasNorm)) score = 210 + aliasNorm.length;
  else if (rawNorm.includes(aliasNorm)) score = 180 + aliasNorm.length;
  else if (aliasNorm.length <= 2 && 台湾雑誌GS_短い別名一致_(raw, aliasText)) score = 190;

  if (score < 0) return -1;
  if (台湾雑誌GS_言語一致_(supportedLanguages, language)) score += 100;
  return score;
}

function 台湾雑誌GS_短い別名一致_(rawTitle, alias) {
  const escaped = String(alias || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp(`(^|[^A-Z0-9])${escaped.toUpperCase()}([^A-Z0-9]|$)`);
  return re.test(String(rawTitle || '').toUpperCase());
}
function 台湾雑誌GS_正規化雑誌文字列_(value) {
  return String(value || '')
    .normalize('NFKC')
    .toUpperCase()
    .replace(/[（(].*?[)）]/g, ' ')
    .replace(/&/g, ' AND ')
    .replace(/\b(?:KOREA|TAIWAN|HONG\s*KONG|HONGKONG|CHINA|THAILAND|HK|TW|CN|KR|TH)\b/g, ' ')
    .replace(/(?:韓国版|韓國版|台湾版|臺灣版|中国版|中國版|香港版|タイ版)/g, ' ')
    .replace(/[^A-Z0-9一-龥ぁ-んァ-ヶ]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/* ============================================================
 * 原題解析
 * ============================================================ */

function 台湾雑誌GS_原題から年月号数を抽出_(rawTitle) {
  const text = String(rawTitle || '')
    .replace(/[　]/g, ' ').replace(/[／]/g, '/').replace(/[－—–ー]/g, '-').replace(/\s+/g, ' ').trim();

  let year = '', month = '', issue = '', issueDisplay = '';

  let m = text.match(/(\d{1,2}(?:\s*[.\/\-・~～]\s*\d{1,2})?)\s*月(?:號|号)?\s*\/\s*(20\d{2})/i);
  if (m) { month = m[1].replace(/\s*/g, ''); year = m[2]; }

  if (!year || !month) {
    m = text.match(/(20\d{2})\s*(?:年|\/|\-|\.)\s*(\d{1,2}(?:\s*[.\/\-・~～]\s*\d{1,2})?)\s*月?(?:號|号)?/i);
    if (m) { year = m[1]; month = m[2].replace(/\s*/g, ''); }
  }

  if (!year) { m = text.match(/\b(20\d{2})\b/); if (m) year = m[1]; }
  if (!month) { m = text.match(/(\d{1,2}(?:\s*[.\/\-・~～]\s*\d{1,2})?)\s*月(?:號|号)?/i); if (m) month = m[1].replace(/\s*/g, ''); }

  m = text.match(/\bVOL\.?\s*([0-9]{1,5})\b/i);
  if (m) { issue = m[1]; issueDisplay = `Vol.${issue}`; }

  if (!issue) {
    m = text.match(/第\s*([0-9]{1,5})\s*(號|号|期)/i);
    if (m) { issue = m[1]; issueDisplay = `第${issue}${m[2]}`; }
  }

  if (!issue) {
    const re = /([0-9]{1,5})\s*(號|号|期)/ig;
    let mm;
    while ((mm = re.exec(text)) !== null) {
      const prev = text.substring(Math.max(0, mm.index - 1), mm.index);
      if (prev === '月') continue;
      issue = mm[1]; issueDisplay = `${issue}${mm[2]}`; break;
    }
  }

  return { year, month, issue, issueDisplay };
}

function 台湾雑誌GS_原題を解析_(rawTitle) {
  const base = 台湾雑誌GS_原題から年月号数を抽出_(rawTitle);
  return {
    ...base,
    edition: 台湾雑誌GS_ルール一致_(rawTitle, 台湾雑誌GS_ルール一覧を取得_(設定_台湾雑誌.スナップ版種ルール名, 設定_台湾雑誌.版種ルール名)),
    type:    台湾雑誌GS_ルール一致_(rawTitle, 台湾雑誌GS_ルール一覧を取得_(設定_台湾雑誌.スナップタイプルール名, 設定_台湾雑誌.タイプルール名))
  };
}

function 台湾雑誌GS_月コードへ変換_(月値) {
  const nums = String(月値 || '').match(/\d+/g);
  if (!nums) return '';
  return nums.map(n => String(parseInt(n, 10)).padStart(2, '0')).join('');
}

function 台湾雑誌GS_月表示へ変換_(月値) {
  const nums = String(月値 || '').match(/\d+/g);
  if (!nums) return '';
  const arr = nums.map(n => String(parseInt(n, 10)));
  return arr.length === 1 ? `${arr[0]}月号` : `${arr.join('・')}月号`;
}

function 台湾雑誌GS_雑誌名候補を整形_(rawTitle) {
  let title = String(rawTitle || '').trim();
  if (!title) return '';
  const patterns = [
    /\s+\d{1,2}(?:\s*[.\/\-・~～]\s*\d{1,2})?\s*月(?:號|号)?\s*\/\s*20\d{2}.*$/i,
    /\s+20\d{2}\s*(?:年|\/|\-|\.)\s*\d{1,2}(?:\s*[.\/\-・~～]\s*\d{1,2})?\s*月?(?:號|号)?.*$/i,
    /\s+VOL\.?\s*\d+.*$/i,
    /\s+第\s*\d+\s*(?:號|号|期).*$/i,
    /\s+\d+\s*(?:號|号|期).*$/i
  ];
  for (const pat of patterns) {
    if (pat.test(title)) { title = title.replace(pat, ''); break; }
  }
  return title.replace(/\s+/g, ' ').trim();
}

/* ============================================================
 * 親コード・商品名生成
 * ============================================================ */

function 台湾雑誌GS_親コードを生成_(言語, 雑誌名, 年, 月, 号数, parsed) {
  const info = 台湾雑誌GS_マスターを検索_(雑誌名, 言語);
  if (!info || !info.略称) return 'ERROR:マスター未登録';

  let codeCore = '';
  if (info.基本キー型 === '号数型') {
    if (!号数) return 'ERROR:号数未入力';
    codeCore = `${info.略称}${号数}`;
  } else {
    if (!年 || !月) return 'ERROR:年月未入力';
    const yy = String(年).slice(-2);
    const mmKey = 台湾雑誌GS_月コードへ変換_(月);
    if (!mmKey) return 'ERROR:月形式不正';
    codeCore = `${info.略称}${yy}${mmKey}`;
  }

  if (parsed.edition.codeSuffix) codeCore += parsed.edition.codeSuffix;
  if (parsed.type.codeSuffix) {
    const joiner = parsed.edition.codeSuffix
      ? (info.版種ありタイプ結合記号 || '-')
      : (info.通常タイプ結合記号 || '');
    codeCore += `${joiner}${parsed.type.codeSuffix}`;
  }

  if (!設定_台湾雑誌.親コードに言語接頭辞を付ける) return codeCore;
  const languageCode = 台湾雑誌GS_言語コードへ変換_(言語);
  return languageCode ? `${languageCode}-${codeCore}` : codeCore;
}
function 台湾雑誌GS_商品名を生成_(言語, 雑誌名, 年, 月, 号数, 表紙情報, 特典メモ, parsed) {
  if (!言語 || !雑誌名) return '';

  const info = 台湾雑誌GS_マスターを検索_(雑誌名, 言語);
  const 英字 = info
    ? (info.表示英字名 || info.英字名)
    : (台湾雑誌GS_英字表示名を整形_(台湾雑誌GS_英字地域表記を除去_(雑誌名)) || 雑誌名);
  const カタカナ = info
    ? (info.表示カタカナ名 || info.カタカナ名)
    : 台湾雑誌GS_表示カタカナ名を決定_('');
  const 基本キー型 = info ? info.基本キー型 : '年月型';
  const languageCode = 台湾雑誌GS_言語コードへ変換_(言語);
  const 言語表記 = 設定_台湾雑誌.言語表記マップ[languageCode] || `${台湾雑誌GS_言語表示へ変換_(言語)}版`;

  let name = `${言語表記} 雑誌 ${英字}`;
  if (カタカナ) name += ` (${カタカナ})`;

  if (基本キー型 === '号数型') {
    if (!号数) return '';
    name += ` ${parsed.issueDisplay || `Vol.${号数}`}`;
    if (年 && 月) name += ` (${年}年${台湾雑誌GS_月表示へ変換_(月)})`;
  } else {
    if (!年 || !月) return '';
    name += ` ${年}年${台湾雑誌GS_月表示へ変換_(月)}`;
    if (parsed.issueDisplay) name += ` ${parsed.issueDisplay}`;
  }

  if (parsed.edition.titleDisplay) name += ` ${parsed.edition.titleDisplay}`;
  if (表紙情報 && String(表紙情報).trim()) name += ` (${String(表紙情報).trim()})`;
  if (parsed.type.titleDisplay) name += ` ${parsed.type.titleDisplay}`;
  if (特典メモ && String(特典メモ).trim()) name += ` ${String(特典メモ).trim()}`;

  return name.replace(/\s+/g, ' ').trim();
}
/* ============================================================
 * 行再計算
 * ============================================================ */

function 台湾雑誌_行を再計算_(sh, row, 編集列番号 = 0, 強制上書き = false) {
  if (!sh || sh.getName() !== '台湾雑誌' || row < 2) return;

  const map = 台湾雑誌GS_ヘッダーMap_(sh);
  const get = (name) => map[name] ? sh.getRange(row, map[name]).getValue() : '';
  const set = (name, val) => { if (map[name]) sh.getRange(row, map[name]).setValue(val); };

  const 原題タイトル列 = map['原題タイトル'] || -1;
  const 原題商品名列 = map['原題商品名'] || -1;

  const 言語Raw = String(get('言語') || '').trim();
  const 言語 = 台湾雑誌GS_言語表示へ変換_(言語Raw);
  if (言語 && 言語 !== 言語Raw) set('言語', 言語);

  let 雑誌名 = String(get('雑誌名') || '').trim();

  const 原題タイトル = String(get('原題タイトル') || '').trim();
  const 原題商品名 = String(get('原題商品名') || '').trim();
  const 解析用原題 = 原題商品名 || String(get('表紙情報') || '').trim() || 原題タイトル;
  const 推定用原題 = 原題タイトル || 原題商品名 || String(get('表紙情報') || '').trim();

  const 表紙情報 = String(get('表紙情報') || '').trim();
  const 特典メモ = String(get('特典メモ') || '').trim();

  const 現在情報 = 雑誌名 ? 台湾雑誌GS_マスターを検索_(雑誌名, 言語) : null;
  const 正式雑誌名 = 現在情報 ? (現在情報.表示英字名 || 現在情報.英字名) : '';
  if (正式雑誌名 && 正式雑誌名 !== 雑誌名) {
    雑誌名 = 正式雑誌名;
    set('雑誌名', 雑誌名);
  }

  const 推定雑誌名 = 台湾雑誌GS_雑誌名を推定_(推定用原題, 言語);
  if (推定雑誌名 && (!雑誌名 || 雑誌名 !== 推定雑誌名)) {
    雑誌名 = 推定雑誌名;
    set('雑誌名', 雑誌名);
  }
  if (雑誌名 && 原題タイトル !== 雑誌名) {
    set('原題タイトル', 雑誌名);
  }

  const 計算用雑誌名 = 雑誌名 || 台湾雑誌GS_雑誌名候補を整形_(推定用原題 || 解析用原題);
  const parsed = 台湾雑誌GS_原題を解析_(解析用原題);

  let 年 = String(get('年') || '').trim();
  let 月 = String(get('月') || '').trim();
  let 号数 = String(get('号数') || '').trim();

  const 原題系を編集した = 強制上書き || [原題タイトル列, 原題商品名列].includes(編集列番号);

  if (原題系を編集した) {
    if (parsed.year) {
      年 = String(parsed.year).trim();
      set('年', 年);
    }
    if (parsed.month) {
      月 = String(parsed.month).trim();
      set('月', 月);
    }
    if (parsed.issue) {
      号数 = String(parsed.issue).trim();
      set('号数', 号数);
    }
  } else {
    if (!年 && parsed.year) {
      年 = String(parsed.year).trim();
      set('年', 年);
    }
    if (!月 && parsed.month) {
      月 = String(parsed.month).trim();
      set('月', 月);
    }
    if (!号数 && parsed.issue) {
      号数 = String(parsed.issue).trim();
      set('号数', 号数);
    }
  }

  if (!言語 || !計算用雑誌名) return;

  const 親コード = 台湾雑誌GS_親コードを生成_(言語, 計算用雑誌名, 年, 月, 号数, parsed);
  if (親コード && !String(親コード).startsWith('ERROR')) {
    set('親コード', 親コード);
  }

  const 商品名 = 台湾雑誌GS_商品名を生成_(言語, 計算用雑誌名, 年, 月, 号数, 表紙情報, 特典メモ, parsed);
  if (商品名) {
    set('商品名（出品用）', 商品名);
  }

  if (計算用雑誌名 && !get('登録日')) {
    set('登録日', new Date());
  }
}
/* ============================================================
 * onEdit
 * ============================================================ */

function 台湾雑誌_onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== '台湾雑誌') return;
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < 2) return;

  const map = 台湾雑誌GS_ヘッダーMap_(sh);
  const 監視列番号 = [
    map['言語'], map['雑誌名'], map['年'], map['月'], map['号数'],
    map['表紙情報'], map['特典メモ'], map['原題タイトル'], map['原題商品名']
  ].filter(Boolean);
  if (!監視列番号.includes(col)) return;

  // 候補追加をトリガーする列（雑誌名・原題系のみ）
  const 候補追加対象列 = [
    map['雑誌名'], map['原題タイトル'], map['原題商品名']
  ].filter(Boolean);

  台湾雑誌_行を再計算_(sh, row, col);

  if (候補追加対象列.includes(col)) {
    台湾雑誌_未解決候補を自動追加_(sh, row);
  }
}

function 台湾雑誌_現在行を再計算() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (!sh || sh.getName() !== '台湾雑誌') {
    SpreadsheetApp.getUi().alert('台湾雑誌シートで実行してください');
    return;
  }
  const row = sh.getActiveRange().getRow();
  if (row < 2) {
    SpreadsheetApp.getUi().alert('2行目以降を選択してください');
    return;
  }
  const col = sh.getActiveRange().getColumn();

  台湾雑誌_行を再計算_(sh, row, col, true);
  台湾雑誌_未解決候補を自動追加_(sh, row);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `台湾雑誌 ${row}行目を再計算しました`, '完了', 3
  );
}

/* ============================================================
 * 確定発行
 * ============================================================ */

function 台湾雑誌_確定発行() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('台湾雑誌');
  if (!sh || sh.getLastRow() < 2) { ui.alert('データがありません'); return; }

  const ヘッダー = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const 列 = {};
  ヘッダー.forEach((h, i) => { if (h) 列[String(h).trim()] = i + 1; });

  if (!列['発番発行']) { ui.alert('発番発行列が見つかりません'); return; }

  const 全データ = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const get行 = (row, 名前) => 列[名前] ? row[列[名前] - 1] : '';

  const 対象行リスト = [];
  全データ.forEach((row, i) => {
    if (get行(row, '発番発行') === true) 対象行リスト.push({ row, rowNum: i + 2 });
  });

  if (対象行リスト.length === 0) { ui.alert('発番発行列にチェックが入っている行がありません'); return; }

  const 重複チェックMap = {};
  全データ.forEach((row, i) => {
    const キー = 台湾雑誌_重複キーを作成_(
      String(get行(row, '言語') || '').trim(),
      String(get行(row, '雑誌名') || '').trim(),
      String(get行(row, '年') || '').trim(),
      String(get行(row, '月') || '').trim(),
      String(get行(row, '号数') || '').trim()
    );
    if (!キー) return;
    if (!重複チェックMap[キー]) 重複チェックMap[キー] = [];
    重複チェックMap[キー].push(i + 2);
  });

  const ブロックリスト = [];
  対象行リスト.forEach(({ row, rowNum }) => {
    const キー = 台湾雑誌_重複キーを作成_(
      String(get行(row, '言語') || '').trim(),
      String(get行(row, '雑誌名') || '').trim(),
      String(get行(row, '年') || '').trim(),
      String(get行(row, '月') || '').trim(),
      String(get行(row, '号数') || '').trim()
    );
    if (!キー) return;
    const 重複行 = (重複チェックMap[キー] || []).filter(r => r !== rowNum);
    if (重複行.length > 0) ブロックリスト.push(`${rowNum}行目：「${キー}」が${重複行[0]}行目と重複`);
  });

if (ブロックリスト.length > 0) {
    ui.alert(`⚠️ 重複が見つかりました。\n\n${ブロックリスト.join('\n')}\n\n重複を解消してから再度実行してください。`);
    return;
  }

  // ★ ここに追加 ↓
  const 親コード重複Map = {};
  全データ.forEach((row, i) => {
    const code = String(get行(row, '親コード') || '').trim();
    if (!code || code.startsWith('ERROR')) return;
    if (!親コード重複Map[code]) 親コード重複Map[code] = [];
    親コード重複Map[code].push(i + 2);
  });

  const 親コードブロックリスト = [];
  対象行リスト.forEach(({ row, rowNum }) => {
    const code = String(get行(row, '親コード') || '').trim();
    if (!code || code.startsWith('ERROR')) return;
    const 重複行 = (親コード重複Map[code] || []).filter(r => r !== rowNum);
    if (重複行.length > 0) {
      親コードブロックリスト.push(`${rowNum}行目：親コード「${code}」が${重複行[0]}行目と重複`);
    }
  });

  if (親コードブロックリスト.length > 0) {
    ui.alert(`⚠️ 親コードの重複が見つかりました。\n\n${親コードブロックリスト.join('\n')}\n\n重複を解消してから再度実行してください。`);
    return;
  }
  // ★ ここまで
  const lock = LockService.getDocumentLock();
  lock.waitLock(60000);
  try {
    let 発行数 = 0;
    対象行リスト.forEach(({ row, rowNum }) => {
      const 親コード = String(get行(row, '親コード') || '').trim();
      if (!親コード || 親コード.startsWith('ERROR')) {
        ui.alert(`${rowNum}行目：親コードが未生成のためスキップしました`);
        return;
      }
      sh.getRange(rowNum, 列['登録状況']).setValue('未登録');
      sh.getRange(rowNum, 列['発番発行']).setValue(false);
      発行数++;
    });
    ui.alert(`✅ 確定発行完了: ${発行数}件`);
  } finally {
    lock.releaseLock();
  }
}

function 台湾雑誌_重複キーを作成_(言語, 雑誌名, 年, 月, 号数) {
  const 言語表示 = 台湾雑誌GS_言語表示へ変換_(言語);
  if (!雑誌名 || !言語表示) return '';
  if (号数) return `${言語表示}-${雑誌名}-${号数}`;
  if (年 && 月) return `${言語表示}-${雑誌名}-${年}-${月}`;
  return '';
}

/* ============================================================
 * プルダウン更新（スナップショット更新も同時実行）
 * ============================================================ */

function 台湾雑誌_プルダウン更新() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('台湾雑誌');
  if (!sh) { SpreadsheetApp.getUi().alert('台湾雑誌シートが見つかりません'); return; }

  台湾雑誌_スナップショットを更新_();

  const 最終行 = Math.max(sh.getLastRow() + 100, 200);
  const 列 = 台湾雑誌GS_ヘッダーMap_(sh);
  const masterSS = 台湾雑誌GS_共通SS_();

  if (列['言語']) {
    const 言語マスターSh = masterSS.getSheetByName('言語マスター');
    let 言語値 = [];
    if (言語マスターSh && 言語マスターSh.getLastRow() >= 2) {
      言語値 = 言語マスターSh.getRange(2, 1, 言語マスターSh.getLastRow() - 1, 1)
        .getValues().flat().map(v => String(v).trim()).filter(v => v);
    }
    sh.getRange(2, 列['言語'], 最終行 - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(言語値.length ? 言語値 : 設定_台湾雑誌.言語リスト, true).build()
    );
  }

  const snapMaster = ss.getSheetByName(設定_台湾雑誌.スナップ雑誌マスター名);
  if (snapMaster && snapMaster.getLastRow() >= 2 && 列['雑誌名']) {
    const snapMap = 台湾雑誌GS_ヘッダーMap_(snapMaster);
    const rows = snapMaster.getRange(2, 1, snapMaster.getLastRow() - 1, snapMaster.getLastColumn()).getValues();
    const seen = new Set();
    const 雑誌値 = [];

    rows.forEach(row => {
      const info = 台湾雑誌GS_マスター行を情報化_(row, snapMap);
      const displayName = info.表示英字名 || info.英字名;
      const key = 台湾雑誌GS_英字表示キー_(displayName);
      if (!displayName || !key || seen.has(key)) return;
      seen.add(key);
      雑誌値.push(displayName);
    });

    雑誌値.sort((a, b) => a.localeCompare(b, 'en', { sensitivity: 'base' }));

    if (雑誌値.length) {
      sh.getRange(2, 列['雑誌名'], 最終行 - 1, 1).setDataValidation(
        SpreadsheetApp.newDataValidation().requireValueInList(雑誌値, true).build()
      );
    }
  }

  if (列['登録状況']) {
    sh.getRange(2, 列['登録状況'], 最終行 - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['未登録', '登録済み', '売り切れ'], true).build()
    );
  }

  if (列['発番発行']) sh.getRange(2, 列['発番発行'], 最終行 - 1, 1).insertCheckboxes();

  if (列['配送パターン']) {
    const 配送マスター = masterSS.getSheetByName('配送パターンマスター');
    if (配送マスター && 配送マスター.getLastRow() >= 2) {
      const 配送値 = 配送マスター.getRange(2, 1, 配送マスター.getLastRow() - 1, 1)
        .getValues().flat().map(v => String(v).trim()).filter(v => v);
      if (配送値.length) {
        sh.getRange(2, 列['配送パターン'], 最終行 - 1, 1).setDataValidation(
          SpreadsheetApp.newDataValidation().requireValueInList(配送値, true).build()
        );
      }
    }
  }

  if (列['親コード']) {
    const 親コード列 = 列['親コード'];
    const colLetter = String.fromCharCode(64 + 親コード列);
    const rules = sh.getConditionalFormatRules();
    const 既存除外 = rules.filter(r =>
      !r.getRanges().some(rng => rng.getColumn() === 1 && rng.getLastColumn() === sh.getLastColumn())
      || !r.getBooleanCondition()
      || (r.getBooleanCondition().getCriteriaType() !== SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA)
    );
    既存除外.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND(${colLetter}2<>"",COUNTIF(${colLetter}:${colLetter},${colLetter}2)>1)`)
        .setBackground('#f4cccc')
        .setFontColor('#cc0000')
        .setRanges([sh.getRange(2, 1, 最終行 - 1, sh.getLastColumn())])
        .build()
    );
    sh.setConditionalFormatRules(既存除外);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('✅ 台湾雑誌 プルダウン更新完了（スナップショット更新済み）', '完了', 3);
}/* ============================================================
 * 雑誌マスター候補 自動収集・正式反映
 *
 * 【列構成】
 *   前面（人が触る列）:
 *     A: ステータス  B: 反映  C: 略称コード候補
 *     D: 雑誌名（英字）  E: 雑誌名（カタカナ）
 *     F: 基本キー型候補  G: 通常タイプ結合記号
 *     H: 版種ありタイプ結合記号  I: 対応言語
 *     J: 別名1  K: 別名2  L: 別名3  M: 備考
 *   裏（参照用）:
 *     N: 元ファイル  O: 元シート  P: 元行
 *     Q: 初回検出日時  R: 最終検出日時  S: 正規化キー
 *
 * 【ステータス5段階】
 *   未対応 → 確認中 → 登録待ち → マスター登録済み / 無視
 *
 * 【運用フロー】
 *   1. 台湾雑誌シートで未解決 → 候補シートへ自動追加（未対応）
 *   2. 担当者が確認 → ステータスを「確認中」「登録待ち」に変更
 *   3. 略称コード・別名などを入力して完成させる
 *   4. 「反映」にチェック → メニューから「正式マスターへ反映」実行
 *   5. 自動で本番マスターへコピー・ステータスが「マスター登録済み」に変わる
 * ============================================================ */

const 候補ステータス = {
  未対応:   '未対応',
  確認中:   '確認中',
  登録待ち: '登録待ち',
  登録済み: 'マスター登録済み',
  無視:     '無視'
};

// 列番号定数（候補シート用）
const 候補列 = {
  ステータス: 1, 反映: 2, 略称コード: 3,
  英字名: 4, カタカナ名: 5, 基本キー型: 6,
  通常結合: 7, 版種結合: 8, 言語: 9,
  別名1: 10, 別名2: 11, 別名3: 12, 備考: 13,
  元ファイル: 14, 元シート: 15, 元行: 16,
  初回検出: 17, 最終検出: 18, 正規化キー: 19
};

const 候補列数 = 19;

/**
 * 候補シートを取得（共通マスターファイル内に作成）
 */
function 台湾雑誌_候補シートを確保_(masterSS) {
  const シート名 = 設定_台湾雑誌.候補シート名;
  let sh = masterSS.getSheetByName(シート名);
  if (sh) return sh;

  sh = masterSS.insertSheet(シート名);

  const headers = [
    'ステータス', '反映', '略称コード候補',
    '雑誌名（英字）', '雑誌名（カタカナ）', '基本キー型候補',
    '通常タイプ結合記号', '版種ありタイプ結合記号', '対応言語',
    '別名1', '別名2', '別名3', '備考',
    '元ファイル', '元シート', '元行',
    '初回検出日時', '最終検出日時', '正規化キー'
  ];

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダー色（前面列=オレンジ、裏列=グレー）
  sh.getRange(1, 1, 1, 13)
    .setBackground('#e69138').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(1, 14, 1, 6)
    .setBackground('#999999').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center');

  // ステータスプルダウン（A列）
  sh.getRange(2, 候補列.ステータス, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(Object.values(候補ステータス), true).build()
  );

  // 反映チェックボックス（B列）
  sh.getRange(2, 候補列.反映, 1000, 1).insertCheckboxes();

  // 基本キー型プルダウン（F列）
  sh.getRange(2, 候補列.基本キー型, 1000, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['年月型', '号数型'], true).build()
  );

  // 条件付き書式（ステータスで行全体を色分け）
  const range = sh.getRange(2, 1, 1000, 13);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2="未対応"')
      .setBackground('#fff2cc').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2="確認中"')
      .setBackground('#fce5cd').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2="登録待ち"')
      .setBackground('#cfe2f3').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2="マスター登録済み"')
      .setBackground('#d9ead3').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2="無視"')
      .setBackground('#efefef').setFontColor('#999999').setRanges([range]).build(),
  ]);

  // 列幅
  const 列幅 = {
    1: 120, 2: 45, 3: 120, 4: 200, 5: 160,
    6: 100, 7: 75, 8: 75, 9: 75,
    10: 160, 11: 160, 12: 160, 13: 300,
    14: 160, 15: 100, 16: 60, 17: 140, 18: 140, 19: 0
  };
  Object.entries(列幅).forEach(([col, width]) => {
    if (width > 0) sh.setColumnWidth(Number(col), width);
  });

  // 正規化キー列（S列）を非表示
  sh.hideColumns(候補列.正規化キー);

  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);

  return sh;
}

/**
 * 既存マスターに完全一致するか（英字名＋別名1〜3）
 */
function 台湾雑誌_マスターに完全一致するか_(masterSh, 候補名) {
  if (!masterSh || masterSh.getLastRow() < 2 || !候補名) return false;
  const map = 台湾雑誌GS_ヘッダーMap_(masterSh);
  const rows = masterSh.getRange(2, 1, masterSh.getLastRow() - 1, masterSh.getLastColumn()).getValues();
  const target = 台湾雑誌GS_比較用文字列_(候補名);
  return rows.some(row => {
    return [
      String(row[(map['雑誌名（英字）'] || 1) - 1] || '').trim(),
      String(row[(map['別名1'] || 0) - 1] || '').trim(),
      String(row[(map['別名2'] || 0) - 1] || '').trim(),
      String(row[(map['別名3'] || 0) - 1] || '').trim(),
    ].filter(v => v).some(v => 台湾雑誌GS_比較用文字列_(v) === target);
  });
}

/**
 * 候補シートに既に同じ正規化キーがあるか
 */
function 台湾雑誌_候補シートに既にあるか_(candidateSh, 正規化キー) {
  if (!candidateSh || candidateSh.getLastRow() < 2 || !正規化キー) return false;
  const values = candidateSh
    .getRange(2, 候補列.正規化キー, candidateSh.getLastRow() - 1, 1)
    .getDisplayValues().flat();
  return values.some(v => String(v).trim() === 正規化キー);
}

/**
 * 候補シートの既存行（正規化キーで検索）の行番号を返す（なければ-1）
 */
function 台湾雑誌_候補シートの既存行番号_(candidateSh, 正規化キー) {
  if (!candidateSh || candidateSh.getLastRow() < 2 || !正規化キー) return -1;
  const values = candidateSh
    .getRange(2, 候補列.正規化キー, candidateSh.getLastRow() - 1, 1)
    .getDisplayValues().flat();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i]).trim() === 正規化キー) return i + 2;
  }
  return -1;
}

/**
 * 未解決行を候補シートへ自動追加する
 * - 親コードが「ERROR:マスター未登録」の行だけ対象
 * - 正規化キーで重複防止
 * - 既にある場合は最終検出日時だけ更新
 */
function 台湾雑誌_未解決候補を自動追加_(sh, row) {
  if (!sh || sh.getName() !== '台湾雑誌' || row < 2) return false;

  const map = 台湾雑誌GS_ヘッダーMap_(sh);
  const get = (name) => map[name]
    ? String(sh.getRange(row, map[name]).getValue() || '').trim() : '';

  const 親コード = get('親コード');
  if (!String(親コード).startsWith('ERROR:マスター未登録')) return false;

  const 雑誌名    = get('雑誌名');
  const 原題タイトル = get('原題タイトル');
  const 原題商品名  = get('原題商品名');
  const 表紙情報   = get('表紙情報');
  const 言語      = get('言語');
  const 原題      = 原題商品名 || 原題タイトル || '';

  const 候補英字名 = (
    雑誌名 || 台湾雑誌GS_雑誌名候補を整形_(原題 || 表紙情報)
  ).trim();
  if (!候補英字名) return false;

  const 正規化キー = 台湾雑誌GS_比較用文字列_(候補英字名);
  const masterSS   = 台湾雑誌GS_共通SS_();
  const masterSh   = masterSS.getSheetByName(設定_台湾雑誌.雑誌マスター名);
  const candidateSh = 台湾雑誌_候補シートを確保_(masterSS);

  // 既存マスターに一致 → 追加しない
  if (台湾雑誌_マスターに完全一致するか_(masterSh, 候補英字名)) return false;

  const 既存行 = 台湾雑誌_候補シートの既存行番号_(candidateSh, 正規化キー);

  // 既に候補シートにある → 最終検出日時だけ更新
  if (既存行 >= 2) {
    candidateSh.getRange(既存行, 候補列.最終検出).setValue(new Date());
    return false;
  }

  // 基本キー型を原題から推定
  const parsed = 台湾雑誌GS_原題を解析_(原題);
  const 基本キー型候補 = (parsed.issue && !parsed.year) ? '号数型' : '年月型';
  const 対応言語 = 台湾雑誌GS_言語コードへ変換_(言語) || '';

  const 備考 = [
    原題    ? `原題:${原題}` : '',
    表紙情報 ? `表紙:${表紙情報}` : ''
  ].filter(Boolean).join(' / ');

  const now = new Date();
  const writeRow = Math.max(candidateSh.getLastRow() + 1, 2);

  candidateSh.getRange(writeRow, 1, 1, 候補列数).setValues([[
    候補ステータス.未対応,  // A: ステータス
    false,                  // B: 反映
    '',                     // C: 略称コード候補
    候補英字名,             // D: 雑誌名（英字）
    '',                     // E: カタカナ
    基本キー型候補,         // F: 基本キー型
    '',                     // G: 通常タイプ結合記号
    '-',                    // H: 版種ありタイプ結合記号
    対応言語,               // I: 対応言語
    '', '', '',             // J〜L: 別名1〜3
    備考,                   // M: 備考
    sh.getParent().getName(), // N: 元ファイル
    sh.getName(),           // O: 元シート
    row,                    // P: 元行
    now,                    // Q: 初回検出日時
    now,                    // R: 最終検出日時
    正規化キー              // S: 正規化キー（隠し列）
  ]]);

  // B列をチェックボックスに設定
  candidateSh.getRange(writeRow, 候補列.反映).insertCheckboxes();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `未解決雑誌を候補シートへ追加: ${候補英字名}`,
    '📋 雑誌マスター候補',
    4
  );

  return true;
}

/**
 * 候補シートから正式マスターへ反映する
 * 「反映」チェックONの行を処理する
 *
 * 【チェック内容】
 *   - 略称コードが空 → スキップ＋警告リスト
 *   - 雑誌名（英字）が空 → スキップ
 *   - 既にマスターに同名がある → スキップ＋警告リスト
 *   - 問題なければ本番マスターへコピー
 *   - 反映後にステータスを「マスター登録済み」・チェックをOFF
 */
function 台湾雑誌_候補を正式マスターへ反映() {
  const ui = SpreadsheetApp.getUi();
  const masterSS  = 台湾雑誌GS_共通SS_();
  const candidateSh = masterSS.getSheetByName(設定_台湾雑誌.候補シート名);

  if (!candidateSh || candidateSh.getLastRow() < 2) {
    ui.alert('候補シートにデータがありません');
    return;
  }

  const masterSh = masterSS.getSheetByName(設定_台湾雑誌.雑誌マスター名);
  if (!masterSh) {
    ui.alert('雑誌マスター（共通）シートが見つかりません');
    return;
  }

  const lastRow = candidateSh.getLastRow();
  const 全データ = candidateSh.getRange(2, 1, lastRow - 1, 候補列数).getValues();

  // 反映チェックONの行を抽出
  const 対象行リスト = [];
  全データ.forEach((row, i) => {
    if (row[候補列.反映 - 1] === true) {
      対象行リスト.push({ data: row, rowNum: i + 2 });
    }
  });

  if (対象行リスト.length === 0) {
    ui.alert('「反映」にチェックが入っている行がありません\n\n反映したい行のB列にチェックを入れてから実行してください');
    return;
  }

  // 確認ダイアログ
  const res = ui.alert(
    '確認',
    `${対象行リスト.length}件を正式マスターへ反映します。続行しますか？`,
    ui.ButtonSet.OK_CANCEL
  );
  if (res !== ui.Button.OK) return;

  // マスターのヘッダーマップを取得
  const masterMap = 台湾雑誌GS_ヘッダーMap_(masterSh);

  let 反映数 = 0;
  const スキップリスト = [];
  const 反映済み行番号 = [];

  対象行リスト.forEach(({ data, rowNum }) => {
    const 英字名  = String(data[候補列.英字名  - 1] || '').trim();
    const 略称    = String(data[候補列.略称コード - 1] || '').trim();
    const カタカナ = String(data[候補列.カタカナ名 - 1] || '').trim();
    const 基本キー型 = String(data[候補列.基本キー型 - 1] || '').trim() || '年月型';
    const 通常結合 = String(data[候補列.通常結合 - 1] || '').trim();
    const 版種結合 = String(data[候補列.版種結合 - 1] || '-').trim() || '-';
    const 別名1   = String(data[候補列.別名1 - 1] || '').trim();
    const 別名2   = String(data[候補列.別名2 - 1] || '').trim();
    const 別名3   = String(data[候補列.別名3 - 1] || '').trim();

    // バリデーション
    if (!英字名) {
      スキップリスト.push(`${rowNum}行目: 雑誌名（英字）が空`);
      return;
    }
    if (!略称) {
      スキップリスト.push(`${rowNum}行目 [${英字名}]: 略称コードが空`);
      return;
    }
    if (台湾雑誌_マスターに完全一致するか_(masterSh, 英字名)) {
      スキップリスト.push(`${rowNum}行目 [${英字名}]: 既にマスターに登録済み`);
      return;
    }

    // マスターへ書き込む行を構築
    // マスターの列順: 雑誌名（英字）/ 雑誌名（カタカナ）/ 略称コード /
    //                 基本キー型 / 通常タイプ結合記号 / 版種ありタイプ結合記号 / 備考 / 別名1〜3
    const writeRow = Math.max(masterSh.getLastRow() + 1, 2);

    // マスターの既存列順に合わせて書き込む
    const 列順 = [
      '雑誌名（英字）', '雑誌名（カタカナ）', '略称コード',
      '基本キー型', '通常タイプ結合記号', '版種ありタイプ結合記号', '備考',
      '別名1', '別名2', '別名3'
    ];
    const 値マップ = {
      '雑誌名（英字）': 英字名,
      '雑誌名（カタカナ）': カタカナ,
      '略称コード': 略称,
      '基本キー型': 基本キー型,
      '通常タイプ結合記号': 通常結合,
      '版種ありタイプ結合記号': 版種結合,
      '備考': '',
      '別名1': 別名1,
      '別名2': 別名2,
      '別名3': 別名3
    };

    // マスターの実際の列数に合わせて書き込み
    const masterLastCol = masterSh.getLastColumn();
    const masterHeaders = masterSh.getRange(1, 1, 1, masterLastCol).getValues()[0];
    const writeData = masterHeaders.map(h => 値マップ[String(h).trim()] ?? '');
    masterSh.getRange(writeRow, 1, 1, writeData.length).setValues([writeData]);

    反映済み行番号.push(rowNum);
    反映数++;
  });

  // 候補シートのステータス更新・チェックOFF
  if (反映済み行番号.length > 0) {
    反映済み行番号.forEach(rowNum => {
      candidateSh.getRange(rowNum, 候補列.ステータス).setValue(候補ステータス.登録済み);
      candidateSh.getRange(rowNum, 候補列.反映).setValue(false);
    });
  }

  // 結果レポート
  let msg = `✅ 反映完了: ${反映数}件`;
  if (スキップリスト.length > 0) {
    msg += `\n\n⚠️ スキップ: ${スキップリスト.length}件\n${スキップリスト.join('\n')}`;
  }
  if (反映数 > 0) {
    msg += '\n\n台湾雑誌ファイルで「プルダウン更新」を実行してください';
  }
  ui.alert(msg);
}

/**
 * 候補シートの件数をステータス別に表示
 */
function 台湾雑誌_候補件数を確認() {
  const masterSS = 台湾雑誌GS_共通SS_();
  const sh = masterSS.getSheetByName(設定_台湾雑誌.候補シート名);

  if (!sh || sh.getLastRow() < 2) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '候補シートにデータがありません', '📋 雑誌マスター候補', 3
    );
    return;
  }

  const ステータス一覧 = sh.getRange(2, 候補列.ステータス, sh.getLastRow() - 1, 1)
    .getValues().flat().filter(v => v);

  const カウント = {};
  Object.values(候補ステータス).forEach(s => { カウント[s] = 0; });
  ステータス一覧.forEach(s => { if (カウント[s] !== undefined) カウント[s]++; });

  const msg = [
    `未対応: ${カウント[候補ステータス.未対応]}件`,
    `確認中: ${カウント[候補ステータス.確認中]}件`,
    `登録待ち: ${カウント[候補ステータス.登録待ち]}件`,
    `登録済み: ${カウント[候補ステータス.登録済み]}件`,
    `無視: ${カウント[候補ステータス.無視]}件`,
    `\n共通マスターの「${設定_台湾雑誌.候補シート名}」を確認してください`
  ].join('\n');

  SpreadsheetApp.getActiveSpreadsheet().toast(msg, '📋 雑誌マスター候補', 8);
}