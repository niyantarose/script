const 設定_台湾まんが = {
  マスターシート名: '台湾まんが',
  作品シート名: 'Works（書籍専用）',
  言語マスター名: '言語マスター',
  カテゴリマスター名: 'カテゴリマスター',
  形態マスター名: '形態マスターシート',
  マスターファイルID: '1ZAy5wtQCq1ixl47MMEIGnrVWeae-F_NfN34lSOUsN5M',

  作品ヘッダー: [
    'WorksKey', '作品ID', '日本語タイトル', '作者', '原題タイトル',
    '登録済み巻', '最新巻', '更新日時', '最新巻(予約込み)', '予約更新日時'
  ],
  作品列数: 10,

  色パレット: ['#b7e1cd', '#fce8b2', '#f4c7c3', '#c9daf8', '#d9d2e9', '#fce5cd', '#fff2cc', '#d9ead3', '#cfe2f3', '#f4cccc'],

  監視列: [
    '作者', '日本語タイトル', '原題タイトル',
    '言語', 'カテゴリ', '形態(通常/初回限定/特装)',
    '単巻数', 'セット巻数開始番号', 'セット巻数終了番号', '特典メモ',
    'ISBN'
  ],

  列名: {
    発行チェック:     '発番発行',
    商品コード:       '親コード',
    タイトル:         'タイトル',
    作者:             '作者',
    日本語タイトル:   '日本語タイトル',
    原題:             '原題タイトル',
    形態:             '形態(通常/初回限定/特装)',
    言語:             '言語',
    カテゴリ:         'カテゴリ',
    単巻数:           '単巻数',
    セット開始:       'セット巻数開始番号',
    セット終了:       'セット巻数終了番号',
    特典メモ:         '特典メモ',
    ISBN:             'ISBN',
    作品ID:           '作品ID(W)(自動)',
    SKU自動:          'SKU(自動)',
    コードステータス: '商品コードステータス',
    登録状況:         '登録状況',
    博客來商品コード: 'サイト商品コード'
  }
};

/* ============================================================
 * シート作成
 * ============================================================ */

function 台湾まんがシートを作成() {
  台湾書籍系シートを作成する_共通('台湾まんが');
}

/**
 * ※ この関数は同名で他ファイルに重複定義しないこと
 *    もし台湾書籍その他.gsにも同じ関数があるなら、どちらか一方だけ残す
 */
function 台湾書籍系シートを作成する_共通(シート名) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();

  if (ss.getSheetByName(シート名)) {
    ui.alert(`「${シート名}」シートは既に存在します`);
    return;
  }

  const sh = ss.insertSheet(シート名);
  const ヘッダー = [
    '発番発行', '登録状況', '親コード', 'タイトル',
    '作者', '日本語タイトル', '原題タイトル',
    '形態(通常/初回限定/特装)', '言語', 'カテゴリ',
    '単巻数', 'セット巻数開始番号', 'セット巻数終了番号', '特典メモ',
    'ISBN', '作品ID(W)(自動)', 'SKU(自動)', '商品コードステータス', '処理メモ',
    '売価', '原価', '粗利益率', '配送パターン', '登録者', '備考',
    'サイト商品コード', '博客來URL', 'メイン画像URL', '追加画像URL', '発売日',
    '重複チェックキー', '登録日'
  ];

  sh.getRange(1, 1, 1, ヘッダー.length).setValues([ヘッダー]);

  const 自動列 = ['タイトル', '作品ID(W)(自動)', 'SKU(自動)', '商品コードステータス', '重複チェックキー', '登録日'];
  const API列  = ['原題タイトル', 'サイト商品コード', '博客來URL', 'メイン画像URL', '追加画像URL', '発売日', '原価'];

  for (let i = 0; i < ヘッダー.length; i++) {
    let 色 = '#4a86e8';
    if (自動列.includes(ヘッダー[i])) 色 = '#999999';
    if (API列.includes(ヘッダー[i]))  色 = '#e69138';

    sh.getRange(1, i + 1)
      .setBackground(色)
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  }

  sh.setRowHeight(1, 40);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(3);

  const 最終行 = 1000;

  sh.getRange(2, 1, 最終行 - 1, 1).insertCheckboxes();
  sh.getRange(2, 2, 最終行 - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['未登録', '登録済み'], true)
      .build()
  );

  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$B2="未登録"')
      .setBackground('#ffff00')
      .setRanges([sh.getRange(2, 1, 最終行 - 1, ヘッダー.length)])
      .build()
  ]);

  ui.alert(`✅ ${シート名}シートを作成しました\n\nメニュー「プルダウン更新」を実行してください`);
}

/* ============================================================
 * onEdit / 取込後補完
 * ============================================================ */

/**
 * installable onEdit から呼ばれる本体
 * ※ 台湾書籍系_共通.gs 側の共通処理を使う
 */
function 台湾まんが_onEdit(e) {
  台湾書籍系_onEdit_共通_(e, 設定_台湾まんが);
}

/**
 * 拡張機能や setValues() 後に呼ぶ
 * 例:
 *   const startRow = sh.getLastRow() + 1;
 *   sh.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
 *   台湾まんが_取込後補完_(startRow, rows.length);
 */
function 台湾まんが_取込後補完_(startRow, numRows) {
  const sh = SpreadsheetApp.getActive().getSheetByName('台湾まんが');
  if (!sh) return;
  if (!startRow || !numRows || numRows <= 0) return;

  台湾書籍系_追加行補完_共通_(sh, startRow, numRows, 設定_台湾まんが);
}

/* ============================================================
 * データ整理
 * ============================================================ */

function 台湾まんがからまんが以外を削除する() {
  const 実際に削除する = true;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('台湾まんが');

  if (!sh) {
    ss.toast('台湾まんが シートが見つかりません');
    return;
  }

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  if (lastRow < 2) {
    ss.toast('削除対象データがありません');
    return;
  }

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());
  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const カテゴリ列 = headers.indexOf('カテゴリ');
  const titleCol   = headers.indexOf('タイトル');

  const deleteRows = [];

  for (let i = 0; i < data.length; i++) {
    const カテゴリ = カテゴリ列 >= 0 ? String(data[i][カテゴリ列] || '').trim() : '';
    const title    = titleCol   >= 0 ? String(data[i][titleCol] || '').trim() : '';

    const isManga =
      /まんが|漫画/i.test(カテゴリ) ||
      /^台湾版\s*まんが/.test(title);

    if (!isManga) {
      deleteRows.push(i + 2);
    }
  }

  if (deleteRows.length === 0) {
    ss.toast('削除する行がありません');
    return;
  }

  if (!実際に削除する) {
    ss.toast(`削除対象は ${deleteRows.length} 行です`, '確認', 8);
    return;
  }

  deleteRows.sort((a, b) => b - a);
  for (const rowNum of deleteRows) {
    sh.deleteRow(rowNum);
  }

  ss.toast(`✅ ${deleteRows.length} 行を削除しました`, '完了', 8);
}