/**
 * Onopen.gs（台湾ファイル専用）
 * メニュー作成・シート設定マップ・onOpen・onEdit
 */

function シート設定を取得(シート名) {
  const map = {
    '台湾_書籍（コミック/小説/設定集）': 設定_台湾書籍,
  };
  return map[シート名] || null;
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('🌏 台湾商品管理');

  menu
    .addSubMenu(
      ui.createMenu('台湾グッズ')
        .addItem('① 確定発行',       '台湾グッズ_確定発行')
        .addItem('② 重複チェック',   '台湾グッズ_重複チェック')
        .addItem('③ Works更新',      '台湾グッズ_Worksを更新')
        .addItem('⑥ プルダウン更新', '台湾グッズ_プルダウン更新')
    )
    .addSubMenu(
      ui.createMenu('台湾書籍')
        .addItem('① 確定発行',         '台湾書籍_確定発行')
        .addItem('② チェック行を削除', '台湾書籍_削除')
        .addItem('③ 一括更新',         '台湾書籍_一括更新')
        .addItem('⑥ プルダウン更新',   '台湾書籍_プルダウン更新')
        .addSeparator()
        .addItem('🔍 Works重複チェック・統合', '台湾書籍_重複統合')
        .addItem('🔄 WorksKey再正規化',         '台湾書籍_WorksKey再正規化')
        .addItem('🔢 Works ID振り直し',         '台湾書籍_ID振り直し')
        .addItem('🧹 Works孤立エントリー削除', '台湾書籍_孤立削除')
        .addItem('⚙️ Works初期化',             '台湾書籍_Works初期化')
    )
    .addSubMenu(
  ui.createMenu('台湾雑誌')
    .addItem('① 確定発行',                     '台湾雑誌_確定発行')
    .addItem('⑥ プルダウン更新',               '台湾雑誌_プルダウン更新')
    .addSeparator()
    .addItem('📋 候補の件数を確認',             '台湾雑誌_候補件数を確認')
    .addItem('✅ 候補を正式マスターへ反映',     '台湾雑誌_候補を正式マスターへ反映')
    .addSeparator()
    .addItem('🔄 現在行を再計算',               '台湾雑誌_現在行を再計算')
    .addSeparator()
    .addItem('🔧 コード自動生成セットアップ',   '台湾雑誌_雑誌コード自動生成セットアップ')
)
    .addSeparator()
    .addSubMenu(
      ui.createMenu('シート作成・初期設定')
        .addItem('台湾グッズシートを作成',           '台湾グッズシートを作成')
        .addItem('Works（台湾グッズ）シートを作成',  '台湾グッズWorksシートを作成')
        .addItem('作品略称マスター（台湾）を作成',   '台湾作品略称マスターを作成')
        .addSeparator()
        .addItem('台湾書籍シートを作成',             '台湾書籍シートを作成')
        .addItem('Works（書籍専用）シートを作成',    '台湾書籍Worksシートを作成')
        .addSeparator()
        .addItem('雑誌マスター（台湾）を作成',       '台湾雑誌マスターを作成')
        .addItem('台湾雑誌シートを作成',             '台湾雑誌シートを作成')
    )
    .addToUi();
}

/* ============================================================
 * 台湾書籍 メニューラッパー
 * ============================================================ */
function 台湾書籍_確定発行()       { _kyoutuu.確定発行を実行(設定_台湾書籍); }
function 台湾書籍_削除()           { _kyoutuu.削除を実行(設定_台湾書籍); }
function 台湾書籍_一括更新()       { _kyoutuu.一括更新を実行(設定_台湾書籍); }
function 台湾書籍_プルダウン更新() { _kyoutuu.プルダウン更新を実行(設定_台湾書籍); }

function 台湾書籍_Works初期化() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('警告', 'Worksを全削除します。続行？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(設定_台湾書籍.作品シート名);
  if (sh) { const last = sh.getLastRow(); if (last > 1) sh.deleteRows(2, last - 1); }
  else { sh = ss.insertSheet(設定_台湾書籍.作品シート名); }
  sh.getRange(1, 1, 1, 設定_台湾書籍.作品列数).setValues([設定_台湾書籍.作品ヘッダー]);
  ui.alert('✅ Works初期化完了（台湾書籍）');
}

function 台湾書籍_WorksKey再正規化() {
  const ui = SpreadsheetApp.getUi();
  const 作品シート = SpreadsheetApp.getActive().getSheetByName(設定_台湾書籍.作品シート名);
  if (!作品シート || 作品シート.getLastRow() < 2) { ui.alert('Worksにデータがありません'); return; }
  if (ui.alert('確認', '全WorksKeyを再計算し重複を統合します。続行？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try {
    const r = _kyoutuu.WorksKey再正規化を実行(作品シート, 設定_台湾書籍);
    ui.alert(`✅ 完了\nキー更新: ${r.キー更新数}\n重複統合: ${r.統合数}\n削除: ${r.削除行数}`);
  } finally { lock.releaseLock(); }
}

function 台湾書籍_重複統合() {
  const ui = SpreadsheetApp.getUi();
  const 作品シート = SpreadsheetApp.getActive().getSheetByName(設定_台湾書籍.作品シート名);
  if (!作品シート || 作品シート.getLastRow() < 2) { ui.alert('Worksにデータがありません'); return; }
  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try {
    const 重複リスト = _kyoutuu.Works重複を検出(作品シート, 設定_台湾書籍);
    if (重複リスト.length === 0) { ui.alert('✅ 重複はありません！'); return; }
    let レポート = `🔍 重複検出: ${重複リスト.length}件\n\n`;
    for (const dup of 重複リスト) {
      レポート += `【${dup.行配列[0].データ[2]}】 著：${dup.行配列[0].データ[3]}\n`;
      for (const item of dup.行配列) レポート += `  ID:${item.データ[1]} (行${item.行})\n`;
      レポート += '\n';
    }
    if (ui.alert('重複チェック結果', レポート + '\n自動統合しますか？', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
    const r = _kyoutuu.WorksKey再正規化を実行(作品シート, 設定_台湾書籍);
    ui.alert(`✅ 統合完了\n統合: ${r.統合数}件, 削除: ${r.削除行数}行`);
  } finally { lock.releaseLock(); }
}

function 台湾書籍_ID振り直し() {
  const ui = SpreadsheetApp.getUi();
  const 作品シート = SpreadsheetApp.getActive().getSheetByName(設定_台湾書籍.作品シート名);
  if (!作品シート || 作品シート.getLastRow() < 2) { ui.alert('Worksにデータがありません'); return; }
  if (ui.alert('確認', 'Works IDを1から連番に振り直します。続行？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try {
    const r = _kyoutuu.WorksID振り直しを実行(作品シート, 設定_台湾書籍);
    ui.alert(`✅ ID振り直し完了\n変更: ${r.変更数}件`);
  } finally { lock.releaseLock(); }
}

function 台湾書籍_孤立削除() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const 作品シート = ss.getSheetByName(設定_台湾書籍.作品シート名);
  const マスターシート = ss.getSheetByName(設定_台湾書籍.マスターシート名);
  if (!作品シート || 作品シート.getLastRow() < 2) { ui.alert('Worksにデータがありません'); return; }
  if (!マスターシート) { ui.alert('マスターシートがありません'); return; }
  if (ui.alert('確認', '孤立エントリーを削除します。続行？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try {
    const 列マップ = _kyoutuu.列番号を取得(マスターシート);
    const r = _kyoutuu.Works孤立エントリーを削除(作品シート, マスターシート, 列マップ, 設定_台湾書籍);
    ui.alert(r.削除数 === 0 ? '✅ 孤立エントリーはありません！' : `✅ 削除完了: ${r.削除数}件`);
  } finally { lock.releaseLock(); }
}

/* ============================================================
 * Works（書籍専用）シート作成
 * ============================================================ */
function 台湾書籍Worksシートを作成() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const シート名 = 設定_台湾書籍.作品シート名;
  if (ss.getSheetByName(シート名)) { ui.alert(`「${シート名}」シートは既に存在します`); return; }
  const sh = ss.insertSheet(シート名);
  sh.getRange(1, 1, 1, 設定_台湾書籍.作品列数).setValues([設定_台湾書籍.作品ヘッダー]);
  sh.getRange(1, 1, 1, 設定_台湾書籍.作品列数)
    .setBackground('#4a86e8').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  ui.alert('✅ Works（書籍専用）シートを作成しました');
}

/* ============================================================
 * 台湾書籍シート作成
 * ============================================================ */
function 台湾書籍シートを作成() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const シート名 = '台湾_書籍（コミック/小説/設定集）';
  if (ss.getSheetByName(シート名)) { ui.alert(`「${シート名}」シートは既に存在します`); return; }

  const sh = ss.insertSheet(シート名);
  const ヘッダー = [
    '発番発行', '登録状況', '商品コード(SKU)', 'タイトル',
    '作者', '日本語タイトル', '原題タイトル',
    '形態(通常/初回限定/特装)', '言語', 'カテゴリ',
    '単巻数', 'セット巻数開始番号', 'セット巻数終了番号', '特典メモ',
    'ISBN',
    '作品ID(W)(自動)', 'SKU(自動)', '商品コードステータス', '登録状況',
    '売価', '原価', '粗利益率', '配送パターン', '登録者', '備考',
    '博客來商品コード', '博客來URL', 'メイン画像URL', '追加画像URL', '発売日',
    '重複チェックキー', '登録日'
  ];

  sh.getRange(1, 1, 1, ヘッダー.length).setValues([ヘッダー]);

  const 自動列 = ['タイトル', '作品ID(W)(自動)', 'SKU(自動)', '商品コードステータス', '重複チェックキー', '登録日'];
  const API列  = ['原題タイトル', '博客來商品コード', '博客來URL', 'メイン画像URL', '追加画像URL', '発売日', '原価'];

  for (let i = 0; i < ヘッダー.length; i++) {
    const cell = sh.getRange(1, i + 1);
    let 色 = '#4a86e8';
    if (自動列.includes(ヘッダー[i])) 色 = '#999999';
    if (API列.includes(ヘッダー[i]))  色 = '#e69138';
    cell.setBackground(色).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  }

  sh.setRowHeight(1, 40);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(3);

  const 最終行 = 1000;
  sh.getRange(2, ヘッダー.indexOf('発番発行') + 1, 最終行 - 1, 1).insertCheckboxes();
  sh.getRange(2, ヘッダー.indexOf('登録状況') + 1, 最終行 - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['未登録', '登録済み'], true).build()
  );

  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$B2="未登録"')
      .setBackground('#ffff00')
      .setRanges([sh.getRange(2, 1, 最終行 - 1, ヘッダー.length)])
      .build()
  ]);

  const 列幅マップ = {
    '発番発行': 60, '登録状況': 80, '商品コード(SKU)': 180, 'タイトル': 300,
    '作者': 150, '日本語タイトル': 200, '原題タイトル': 200,
    '形態(通常/初回限定/特装)': 160, '言語': 60, 'カテゴリ': 80,
    '単巻数': 60, 'セット巻数開始番号': 100, 'セット巻数終了番号': 100, '特典メモ': 200,
    'ISBN': 140, '作品ID(W)(自動)': 120, 'SKU(自動)': 180, '商品コードステータス': 160,
    '売価': 80, '原価': 80, '粗利益率': 90, '配送パターン': 100, '登録者': 80, '備考': 150,
    '博客來商品コード': 140, '博客來URL': 220, 'メイン画像URL': 200, '追加画像URL': 200, '発売日': 100,
    '重複チェックキー': 150, '登録日': 120
  };
  ヘッダー.forEach((h, i) => { if (列幅マップ[h]) sh.setColumnWidth(i + 1, 列幅マップ[h]); });

  ui.alert('✅ 台湾書籍シートを作成しました\n\nメニュー「プルダウン更新」を実行してください');
}

/* ============================================================
 * onEdit トリガー
 * ============================================================ */
function onEdit(e) {
  if (!e || !e.range) return;

  const シート名 = String(e.range.getSheet().getName() || '').trim();

  if (シート名 === '台湾グッズ') {
    台湾グッズ_onEdit(e);
    return;
  }

  if (シート名 === '台湾雑誌') {
    台湾雑誌_onEdit(e);
    return;
  }

  _kyoutuu.メインonEdit(e);
}