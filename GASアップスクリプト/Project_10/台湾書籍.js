/**
 * 台湾書籍.gs
 * 台湾_書籍（コミック/小説/設定集）シート専用の設定とメニューラッパー
 */

const 設定_台湾書籍 = {
  マスターシート名: '台湾_書籍（コミック/小説/設定集）',
  作品シート名: 'Works（書籍専用）',
  言語マスター名: '言語マスター',
  カテゴリマスター名: 'カテゴリマスター',
  形態マスター名: '形態マスターシート',

  作品ヘッダー: ['WorksKey', '作品ID', '日本語タイトル', '作者', '原題タイトル', '登録済み巻', '最新巻', '更新日時', '最新巻(予約込み)', '予約更新日時'],
  作品列数: 10,

  色パレット: ['#b7e1cd', '#fce8b2', '#f4c7c3', '#c9daf8', '#d9d2e9', '#fce5cd', '#fff2cc', '#d9ead3', '#cfe2f3', '#f4cccc'],

  // 監視列: onEditで変更を検知する列名リスト
  監視列: [
    '作者', '日本語タイトル', '原題タイトル',
    '言語', 'カテゴリ', '形態(通常/初回限定/特装)',
    '単巻数', 'セット巻数開始番号', 'セット巻数終了番号', '特典メモ',
    'ISBN'
  ],

  列名: {
    発行チェック:     '発番発行',
    商品コード:       '商品コード(SKU)',
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
    登録状況:         '登録状況'
  }
};

/* ============================================================
 * 台湾書籍 メニューラッパー
 * ============================================================ */
function 台湾_確定発行()   { 確定発行を実行(設定_台湾書籍); }
function 台湾_削除()       { 削除を実行(設定_台湾書籍); }
function 台湾_一括更新()   { 一括更新を実行(設定_台湾書籍); }
function 台湾_プルダウン更新() { プルダウン更新を実行(設定_台湾書籍); }

function 台湾_Works初期化() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('警告', 'Worksを全削除します。続行？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(設定_台湾書籍.作品シート名);
  if (sh) { const last = sh.getLastRow(); if (last > 1) sh.deleteRows(2, last - 1); }
  else { sh = ss.insertSheet(設定_台湾書籍.作品シート名); }
  sh.getRange(1, 1, 1, 設定_台湾書籍.作品列数).setValues([設定_台湾書籍.作品ヘッダー]);
  ui.alert('✅ Works初期化完了（台湾書籍）');
}

function 台湾_WorksKey再正規化() {
  const ui = SpreadsheetApp.getUi();
  const 作品シート = SpreadsheetApp.getActive().getSheetByName(設定_台湾書籍.作品シート名);
  if (!作品シート || 作品シート.getLastRow() < 2) { ui.alert('Worksにデータがありません'); return; }
  if (ui.alert('確認', '全WorksKeyを再計算し重複を統合します。続行？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try {
    const r = WorksKey再正規化を実行(作品シート, 設定_台湾書籍);
    ui.alert(`✅ 完了\nキー更新: ${r.キー更新数}\n重複統合: ${r.統合数}\n削除: ${r.削除行数}`);
  } finally { lock.releaseLock(); }
}

function 台湾_重複統合() {
  const ui = SpreadsheetApp.getUi();
  const 作品シート = SpreadsheetApp.getActive().getSheetByName(設定_台湾書籍.作品シート名);
  if (!作品シート || 作品シート.getLastRow() < 2) { ui.alert('Worksにデータがありません'); return; }
  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try {
    const 重複リスト = Works重複を検出(作品シート, 設定_台湾書籍);
    if (重複リスト.length === 0) { ui.alert('✅ 重複はありません！'); return; }
    let レポート = `🔍 重複検出: ${重複リスト.length}件\n\n`;
    for (const dup of 重複リスト) {
      レポート += `【${dup.行配列[0].データ[2]}】 著：${dup.行配列[0].データ[3]}\n`;
      for (const item of dup.行配列) レポート += `  ID:${item.データ[1]} (行${item.行})\n`;
      レポート += '\n';
    }
    if (ui.alert('重複チェック結果', レポート + '\n自動統合しますか？', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
    const r = WorksKey再正規化を実行(作品シート, 設定_台湾書籍);
    ui.alert(`✅ 統合完了\n統合: ${r.統合数}件, 削除: ${r.削除行数}行`);
  } finally { lock.releaseLock(); }
}

function 台湾_ID振り直し() {
  const ui = SpreadsheetApp.getUi();
  const 作品シート = SpreadsheetApp.getActive().getSheetByName(設定_台湾書籍.作品シート名);
  if (!作品シート || 作品シート.getLastRow() < 2) { ui.alert('Worksにデータがありません'); return; }
  if (ui.alert('確認', 'Works IDを1から連番に振り直します。続行？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try {
    const r = WorksID振り直しを実行(作品シート, 設定_台湾書籍);
    ui.alert(`✅ ID振り直し完了\n変更: ${r.変更数}件`);
  } finally { lock.releaseLock(); }
}

function 台湾_孤立削除() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const 作品シート = ss.getSheetByName(設定_台湾書籍.作品シート名);
  const マスターシート = ss.getSheetByName(設定_台湾書籍.マスターシート名);
  if (!作品シート || 作品シート.getLastRow() < 2) { ui.alert('Worksにデータがありません'); return; }
  if (!マスターシート) { ui.alert('マスターシートがありません'); return; }
  if (ui.alert('確認', '孤立エントリーを削除します。続行？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try {
    const 列マップ = 列番号を取得(マスターシート);
    const r = Works孤立エントリーを削除(作品シート, マスターシート, 列マップ, 設定_台湾書籍);
    ui.alert(r.削除数 === 0 ? '✅ 孤立エントリーはありません！' : `✅ 削除完了: ${r.削除数}件`);
  } finally { lock.releaseLock(); }
}
