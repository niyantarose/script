/**
 * 韓国マンガ.gs
 * 韓国マンガシート専用の設定とメニューラッパー
 */

const 設定_韓国マンガ = {
  マスターシート名: '韓国マンガ',
  作品シート名: 'Works（韓国マンガ）',
  言語マスター名: '言語マスター',
  カテゴリマスター名: 'カテゴリマスター',
  形態マスター名: '形態マスターシート',
  マスターファイルID: '1ZAy5wtQCq1ixl47MMEIGnrVWeae-F_NfN34lSOUsN5M', // ← 追加

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
    登録状況:         '登録状況'
  }
};

/* ============================================================
 * onEdit トリガー
 * ============================================================ */
function onEdit(e) {
  _kyoutuu.メインonEdit(e);
}

/* ============================================================
 * 韓国マンガ メニューラッパー
 * ============================================================ */
function 韓国マンガ_確定発行()       { _kyoutuu.確定発行を実行(設定_韓国マンガ); }
function 韓国マンガ_削除()           { _kyoutuu.削除を実行(設定_韓国マンガ); }
function 韓国マンガ_一括更新()       { _kyoutuu.一括更新を実行(設定_韓国マンガ); }
function 韓国マンガ_プルダウン更新() { _kyoutuu.プルダウン更新を実行(設定_韓国マンガ); }

function 韓国マンガ_Works初期化() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('警告', 'Worksを全削除します。続行？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(設定_韓国マンガ.作品シート名);
  if (sh) { const last = sh.getLastRow(); if (last > 1) sh.deleteRows(2, last - 1); }
  else { sh = ss.insertSheet(設定_韓国マンガ.作品シート名); }
  sh.getRange(1, 1, 1, 設定_韓国マンガ.作品列数).setValues([設定_韓国マンガ.作品ヘッダー]);
  ui.alert('✅ Works初期化完了（韓国マンガ）');
}

function 韓国マンガ_WorksKey再正規化() {
  const ui = SpreadsheetApp.getUi();
  const 作品シート = SpreadsheetApp.getActive().getSheetByName(設定_韓国マンガ.作品シート名);
  if (!作品シート || 作品シート.getLastRow() < 2) { ui.alert('Worksにデータがありません'); return; }
  if (ui.alert('確認', '全WorksKeyを再計算し重複を統合します。続行？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try {
    const r = _kyoutuu.WorksKey再正規化を実行(作品シート, 設定_韓国マンガ);
    ui.alert(`✅ 完了\nキー更新: ${r.キー更新数}\n重複統合: ${r.統合数}\n削除: ${r.削除行数}`);
  } finally { lock.releaseLock(); }
}

function 韓国マンガ_重複統合() {
  const ui = SpreadsheetApp.getUi();
  const 作品シート = SpreadsheetApp.getActive().getSheetByName(設定_韓国マンガ.作品シート名);
  if (!作品シート || 作品シート.getLastRow() < 2) { ui.alert('Worksにデータがありません'); return; }
  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try {
    const 重複リスト = _kyoutuu.Works重複を検出(作品シート, 設定_韓国マンガ);
    if (重複リスト.length === 0) { ui.alert('✅ 重複はありません！'); return; }
    let レポート = `🔍 重複検出: ${重複リスト.length}件\n\n`;
    for (const dup of 重複リスト) {
      レポート += `【${dup.行配列[0].データ[2]}】 著：${dup.行配列[0].データ[3]}\n`;
      for (const item of dup.行配列) レポート += `  ID:${item.データ[1]} (行${item.行})\n`;
      レポート += '\n';
    }
    if (ui.alert('重複チェック結果', レポート + '\n自動統合しますか？', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
    const r = _kyoutuu.WorksKey再正規化を実行(作品シート, 設定_韓国マンガ);
    ui.alert(`✅ 統合完了\n統合: ${r.統合数}件, 削除: ${r.削除行数}行`);
  } finally { lock.releaseLock(); }
}

function 韓国マンガ_ID振り直し() {
  const ui = SpreadsheetApp.getUi();
  const 作品シート = SpreadsheetApp.getActive().getSheetByName(設定_韓国マンガ.作品シート名);
  if (!作品シート || 作品シート.getLastRow() < 2) { ui.alert('Worksにデータがありません'); return; }
  if (ui.alert('確認', 'Works IDを1から連番に振り直します。続行？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try {
    const r = _kyoutuu.WorksID振り直しを実行(作品シート, 設定_韓国マンガ);
    ui.alert(`✅ ID振り直し完了\n変更: ${r.変更数}件`);
  } finally { lock.releaseLock(); }
}

function 韓国マンガ_孤立削除() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const 作品シート = ss.getSheetByName(設定_韓国マンガ.作品シート名);
  const マスターシート = ss.getSheetByName(設定_韓国マンガ.マスターシート名);
  if (!作品シート || 作品シート.getLastRow() < 2) { ui.alert('Worksにデータがありません'); return; }
  if (!マスターシート) { ui.alert('マスターシートがありません'); return; }
  if (ui.alert('確認', '孤立エントリーを削除します。続行？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try {
    const 列マップ = _kyoutuu.列番号を取得(マスターシート);
    const r = _kyoutuu.Works孤立エントリーを削除(作品シート, マスターシート, 列マップ, 設定_韓国マンガ);
    ui.alert(r.削除数 === 0 ? '✅ 孤立エントリーはありません！' : `✅ 削除完了: ${r.削除数}件`);
  } finally { lock.releaseLock(); }
}

/* ============================================================
 * 韓国マンガシート作成
 * ============================================================ */
function 韓国マンガシートを作成() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const シート名 = 設定_韓国マンガ.マスターシート名;
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
  const API列 = ['商品名(原題)', '著者', '出版社', '発売日', 'ISBN', '原価',
                 'アラジン商品コード', 'アラジンURL', 'yes24商品コード', 'yes24URL',
                 'Kyobo商品コード', 'KyoboURL', 'メイン画像URL', '追加画像URL'];

  for (let i = 0; i < ヘッダー.length; i++) {
    const cell = sh.getRange(1, i + 1);
    let 色 = '#4a86e8';
    if (自動列.includes(ヘッダー[i])) 色 = '#999999';
    if (API列.includes(ヘッダー[i])) 色 = '#e69138';
    cell.setBackground(色);
    cell.setFontColor('#ffffff');
    cell.setFontWeight('bold');
    cell.setHorizontalAlignment('center');
  }

  sh.setRowHeight(1, 40);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);

  const 最終行 = 1000;
  sh.getRange(2, 1, 最終行 - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['未登録', '登録済み'], true).build()
  );

  const 全体範囲 = sh.getRange(2, 1, 最終行 - 1, ヘッダー.length);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2="未登録"')
      .setBackground('#ffff00')
      .setRanges([全体範囲])
      .build()
  ]);

  ui.alert('✅ 韓国マンガシートを作成しました\n\n韓国マンガシートをアクティブにした状態で\nメニュー「⑥プルダウン更新」を実行してください');
}

function 韓国マンガヘッダー色更新() {
  const sh = SpreadsheetApp.getActive().getSheetByName('韓国マンガ');
  if (!sh) return;
  const ヘッダー = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const 自動列 = ['管理コード', '言語コード', '版種コード', '作品番号', 'カテゴリコード', '重複チェック用正規化タイトル', '登録日'];
  const API列 = ['商品名(原題)', '著者', '出版社', '発売日', 'ISBN', '原価',
                 'アラジン商品コード', 'アラジンURL', 'yes24商品コード', 'yes24URL',
                 'Kyobo商品コード', 'KyoboURL', 'メイン画像URL', '追加画像URL'];
  for (let i = 0; i < ヘッダー.length; i++) {
    const cell = sh.getRange(1, i + 1);
    let 色 = '#4a86e8';
    if (自動列.includes(ヘッダー[i])) 色 = '#999999';
    if (API列.includes(ヘッダー[i])) 色 = '#e69138';
    cell.setBackground(色);
    cell.setFontColor('#ffffff');
    cell.setFontWeight('bold');
  }
  Logger.log('✅ ヘッダー色を更新しました');
}