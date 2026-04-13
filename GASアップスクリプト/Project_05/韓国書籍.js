/**
 * 韓国書籍.gs
 * 韓国書籍シート専用の設定とメニューラッパー
 * ※ 共通関数は _kyoutuu ライブラリを使用
 */

const 設定_韓国書籍 = {
  マスターシート名: '韓国書籍',
  作品シート名: 'Works（韓国書籍）',
  言語マスター名: '言語マスター',
  カテゴリマスター名: 'カテゴリマスター',
  形態マスター名: '形態マスターシート',
  マスターファイルID: '1ZAy5wtQCq1ixl47MMEIGnrVWeae-F_NfN34lSOUsN5M',

  作品ヘッダー: ['WorksKey', '作品ID', '日本語タイトル', '作者', '原題タイトル', '登録済み巻', '最新巻', '更新日時', '最新巻(予約込み)', '予約更新日時'],
  作品列数: 10,

  色パレット: ['#b7e1cd', '#fce8b2', '#f4c7c3', '#c9daf8', '#d9d2e9', '#fce5cd', '#fff2cc', '#d9ead3', '#cfe2f3', '#f4cccc'],

  監視列: [
    '著者', '商品名(日本語)', '商品名(原題)',
    '言語', 'カテゴリ', '版種名',
    '単巻数', 'セット巻数開始', 'セット巻数終了', '特典メモ',
    'ISBN'
  ],

  列名: {
    発行チェック:     '発番発行',
    商品コード:       '管理コード',
    タイトル:         'タイトル(自動)',
    作者:             '著者',
    日本語タイトル:   '商品名(日本語)',
    原題:             '商品名(原題)',
    形態:             '版種名',
    言語:             '言語',
    カテゴリ:         'カテゴリ',
    単巻数:           '単巻数',
    セット開始:       'セット巻数開始',
    セット終了:       'セット巻数終了',
    特典メモ:         '特典メモ',
    ISBN:             'ISBN',
    作品ID:           '作品番号',
    SKU自動:          '管理コード',
    コードステータス: 'コードステータス',
    登録状況:         '登録状況',
    配送パターン:     '配送パターン'
  }
};

/* ============================================================
 * 韓国書籍 メニューラッパー
 * ============================================================ */
function 韓国書籍_確定発行()       { _kyoutuu.確定発行を実行(設定_韓国書籍); }
function 韓国書籍_削除()           { _kyoutuu.削除を実行(設定_韓国書籍); }
function 韓国書籍_一括更新()       { _kyoutuu.一括更新を実行(設定_韓国書籍); }
function 韓国書籍_プルダウン更新() { _kyoutuu.プルダウン更新を実行(設定_韓国書籍); }

/* ============================================================
 * 韓国書籍シート作成
 * ============================================================ */
function 韓国書籍シートを作成() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const シート名 = '韓国書籍';
  if (ss.getSheetByName(シート名)) { ui.alert(`「${シート名}」シートは既に存在します`); return; }

  const sh = ss.insertSheet(シート名);
  const ヘッダー = [
    '登録状況', '管理コード', 'タイトル(自動)', '作品番号', '言語コード', '版種コード', 'カテゴリコード',
    '商品名(日本語)', '商品名(原題)', '付録情報', '言語', 'カテゴリ', '版種名',
    '単巻数', 'セット巻数開始', 'セット巻数終了', '特典メモ',
    '著者', '出版社', '発売日', 'ISBN',
    'アラジン商品コード', 'アラジンURL',
    'yes24商品コード', 'yes24URL',
    'Kyobo商品コード', 'KyoboURL',
    'メイン画像URL', '追加画像URL',
    '原価', '売価', '配送パターン',
    '重複チェック用正規化タイトル', '登録日',
    '登録者', '備考'
  ];

  sh.getRange(1, 1, 1, ヘッダー.length).setValues([ヘッダー]);

  const 自動列 = ['管理コード', 'タイトル(自動)', '作品番号', '言語コード', '版種コード', 'カテゴリコード', '重複チェック用正規化タイトル', '登録日'];
  const API列  = ['商品名(原題)', '著者', '出版社', '発売日', 'ISBN', '原価',
                  'アラジン商品コード', 'アラジンURL', 'yes24商品コード', 'yes24URL',
                  'Kyobo商品コード', 'KyoboURL', 'メイン画像URL', '追加画像URL'];

  for (let i = 0; i < ヘッダー.length; i++) {
    const cell = sh.getRange(1, i + 1);
    let 色 = '#4a86e8';
    if (自動列.includes(ヘッダー[i])) 色 = '#999999';
    if (API列.includes(ヘッダー[i]))  色 = '#e69138';
    cell.setBackground(色).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  }

  sh.setRowHeight(1, 40);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);

  const 最終行 = 1000;
  sh.getRange(2, 1, 最終行 - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['未登録', '登録済み'], true).build()
  );

  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2="未登録"')
      .setBackground('#ffff00')
      .setRanges([sh.getRange(2, 1, 最終行 - 1, ヘッダー.length)])
      .build()
  ]);

  ui.alert('✅ 韓国書籍シートを作成しました\n\nメニュー「プルダウン更新」を実行してください');
}