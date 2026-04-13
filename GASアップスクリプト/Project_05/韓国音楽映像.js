/**
 * 韓国音楽映像.gs
 * 韓国音楽映像シート専用の設定とメニューラッパー
 * ※ 共通関数は _kyoutuu ライブラリを使用
 */

const 設定_韓国音楽映像 = {
  マスターシート名: '韓国音楽映像',
  作品シート名: 'Works（韓国音楽映像）',
  言語マスター名: '言語マスター',
  カテゴリマスター名: 'カテゴリマスター',
  形態マスター名: '形態マスターシート',
  マスターファイルID: '1ZAy5wtQCq1ixl47MMEIGnrVWeae-F_NfN34lSOUsN5M',

  作品ヘッダー: ['WorksKey', '作品ID', '日本語タイトル', 'アーティスト', '原題タイトル', '登録済み数', '最新数', '更新日時', '最新数(予約込み)', '予約更新日時'],
  作品列数: 10,

  色パレット: ['#b7e1cd', '#fce8b2', '#f4c7c3', '#c9daf8', '#d9d2e9', '#fce5cd', '#fff2cc', '#d9ead3', '#cfe2f3', '#f4cccc'],

  監視列: [
    'アーティスト名', '日本語タイトル', '商品名(原題)',
    '言語', 'カテゴリ', '特典メモ', 'JANコード'
  ],

  列名: {
    発行チェック:     '発番発行',
    商品コード:       '商品コード(SKU)',
    タイトル:         'タイトル',
    作者:             'アーティスト名',
    日本語タイトル:   '日本語タイトル',
    原題:             '商品名(原題)',
    形態:             '',
    言語:             '言語',
    カテゴリ:         'カテゴリ',
    単巻数:           '',
    セット開始:       '',
    セット終了:       '',
    特典メモ:         '特典メモ',
    ISBN:             'JANコード',
    作品ID:           '作品ID(W)(自動)',
    SKU自動:          'SKU(自動)',
    コードステータス: '商品コードステータス',
    登録状況:         '登録状況',
    配送パターン:     '配送パターン',
    博客來商品コード: 'アラジン商品コード',
  }
};

/* ============================================================
 * 韓国音楽映像 メニューラッパー
 * ============================================================ */
function 韓国音楽映像_確定発行()       { _kyoutuu.確定発行を実行(設定_韓国音楽映像); }
function 韓国音楽映像_削除()           { _kyoutuu.削除を実行(設定_韓国音楽映像); }
function 韓国音楽映像_一括更新()       { _kyoutuu.一括更新を実行(設定_韓国音楽映像); }
function 韓国音楽映像_プルダウン更新() { _kyoutuu.プルダウン更新を実行(設定_韓国音楽映像); }

/* ============================================================
 * 韓国音楽映像シート作成
 * ============================================================ */
function 韓国音楽映像シートを作成() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const シート名 = '韓国音楽映像';
  if (ss.getSheetByName(シート名)) { ui.alert(`「${シート名}」シートは既に存在します`); return; }

  const sh = ss.insertSheet(シート名);
  const ヘッダー = [
    '発番発行', '登録状況', '商品コード(SKU)', 'タイトル',
    'アーティスト名', '日本語タイトル', '商品名(原題)', '特典メモ',
    '言語', 'カテゴリ', '売価', '原価', '粗利益率', '配送パターン',
    '商品説明', 'JANコード', '発売日',
    'アラジン商品コード', 'アラジンURL',
    'yes24商品コード', 'yes24URL',
    'Kyobo商品コード', 'KyoboURL',
    'メイン画像URL', '追加画像URL',
    '作品ID(W)(自動)', 'SKU(自動)', '商品コードステータス',
    '重複チェックキー', '登録日', '登録者', '備考'
  ];

  sh.getRange(1, 1, 1, ヘッダー.length).setValues([ヘッダー]);

  const 自動列 = ['タイトル', '作品ID(W)(自動)', 'SKU(自動)', '商品コードステータス', '重複チェックキー', '登録日'];
  const API列  = ['商品名(原題)', '発売日', '原価', 'JANコード',
                  'アラジン商品コード', 'アラジンURL',
                  'yes24商品コード', 'yes24URL',
                  'Kyobo商品コード', 'KyoboURL',
                  'メイン画像URL', '追加画像URL'];

  for (let i = 0; i < ヘッダー.length; i++) {
    let 色 = '#4a86e8';
    if (自動列.includes(ヘッダー[i])) 色 = '#999999';
    if (API列.includes(ヘッダー[i]))  色 = '#e69138';
    sh.getRange(1, i + 1).setBackground(色).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  }

  sh.setRowHeight(1, 40);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(3);

  const 最終行 = 1000;
  sh.getRange(2, 1, 最終行 - 1, 1).insertCheckboxes();
  sh.getRange(2, 2, 最終行 - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['未登録', '登録済み'], true).build()
  );

  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$B2="未登録"')
      .setBackground('#ffff00')
      .setRanges([sh.getRange(2, 1, 最終行 - 1, ヘッダー.length)])
      .build()
  ]);

  ui.alert('✅ 韓国音楽映像シートを作成しました\n\nメニュー「プルダウン更新」を実行してください');
}