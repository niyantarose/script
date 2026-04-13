/**
 * 呼び出しonOpen.gs
 * メニュー登録とシート設定マップ
 *
 * ★ 新しいシートを追加したら:
 *   1. シートファイル（例: 韓国書籍.gs）を作成
 *   2. 設定_XXX オブジェクトを定義
 *   3. シート設定マップ に追加
 *   4. onOpen にメニュー項目を追加
 */

/* ============================================================
 * シート名 → 設定オブジェクト マップ
 * ============================================================ */
function シート設定を取得(シート名) {
  const マップ = {};

  if (typeof 設定_台湾書籍 !== 'undefined') {
    マップ['台湾書籍（コミック/小説/設定集）'] = 設定_台湾書籍;
  }
  if (typeof 設定_韓国マンガ !== 'undefined') {
    マップ['韓国マンガ'] = 設定_韓国マンガ;
  }
  if (typeof 設定_韓国書籍 !== 'undefined') {
    マップ['韓国書籍'] = 設定_韓国書籍;
  }
  if (typeof 設定_韓国音楽映像 !== 'undefined') {
    マップ['韓国音楽映像'] = 設定_韓国音楽映像;
  }
  if (typeof 設定_韓国グッズ !== 'undefined') {
    マップ['韓国グッズ'] = 設定_韓国グッズ;
  }
  if (typeof 設定_台湾グッズ !== 'undefined') {
    マップ['台湾グッズ'] = 設定_台湾グッズ;
  }
  if (typeof 設定_台湾雑誌 !== 'undefined') {
    マップ['台湾雑誌'] = 設定_台湾雑誌;
  }
  if (typeof 設定_韓国雑誌 !== 'undefined') {
    マップ['韓国雑誌'] = 設定_韓国雑誌;
  }

  return マップ[シート名] || null;
}

/* ============================================================
 * onOpen
 * ============================================================ */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('商品コード管理')

      // ── 共通操作（アクティブシートに作用）──
      .addItem('① 商品コード確定発行', 'メニュー_確定発行')
      .addItem('② 商品コード削除',     'メニュー_削除')
      .addSeparator()
      .addItem('③ 既存データ一括更新', 'メニュー_一括更新')
      .addSeparator()
      .addItem('⑤ マスター作成',       'メニュー_マスター作成')
      .addItem('⑥ プルダウン更新',     'メニュー_プルダウン更新')
      .addItem('🔴 アラジンTTBキー設定',    'メニュー_アラジンTTBキー設定')
　　  .addItem('🔴 アラジン一括取得',        'メニュー_アラジン一括取得')
　　　.addItem('🔴 アラジン列追加', 'メニュー_アラジン列追加')
　　　.addItem('🔴 アラジントリガー列に色付け', 'メニュー_アラジントリガー列に色付け')

      .addSeparator()

      // ── シート作成 ──
      .addItem('📋 韓国マンガシート作成', 'メニュー_韓国マンガシート作成')
      .addItem('📋 韓国書籍シート作成', 'メニュー_韓国書籍シート作成')
      .addItem('📋 韓国音楽映像シート作成', 'メニュー_韓国音楽映像シート作成')
      .addItem('📋 韓国グッズシート作成',    'メニュー_韓国グッズシート作成')
      .addItem('📋 韓国グッズWorksシート作成', 'メニュー_韓国グッズWorksシート作成')
      .addItem('📋 台湾グッズシート作成',        'メニュー_台湾グッズシート作成')
   　　.addItem('📋 台湾グッズWorksシート作成',    'メニュー_台湾グッズWorksシート作成')
  　　 .addItem('📋 作品略称マスター（台湾）作成', 'メニュー_台湾作品略称マスター作成')
　　　　.addItem('🏪 ショップマスター作成',    'メニュー_ショップマスター作成')
　　　　.addItem('💱 為替エイリアステーブル生成', '為替エイリアステーブルを生成')
　　　　.addItem('🎨 グッズ作品マスター作成',  'メニュー_グッズ作品マスター作成')
　　　　.addItem('🎨 グッズジャンルマスター作成', 'メニュー_グッズジャンルマスター作成')
      .addSeparator()

      // ── Works管理（アクティブシートの Works に作用）──
      .addItem('🔍 Works重複チェック・統合', 'メニュー_重複統合')
      .addItem('🔄 WorksKey再正規化',         'メニュー_WorksKey再正規化')
      .addItem('🔢 Works ID振り直し',         'メニュー_ID振り直し')
      .addItem('🧹 Works孤立エントリー削除', 'メニュー_孤立削除')
      .addItem('⚙️ Works初期化',             'メニュー_Works初期化')
　　　　.addSeparator()

      // ── 共通雑誌マスター ──
      .addItem('📰 共通雑誌マスターを初期化', '共通雑誌マスターを初期化')
      .addItem('📰 雑誌名を国名なしへ統合', '共通雑誌マスターを国名なしへ統合')
      .addItem('📰 候補を正式マスターへ反映', '共通雑誌候補を正式反映')
      .addItem('📰 候補シートへテスト追加', '共通雑誌候補テスト追加')
      
      .addToUi();
  } catch (e) {
    Logger.log('onOpenはスプレッドシートから実行してください: ' + e.message);
  }
}

/* ============================================================
 * メニューラッパー（アクティブシートの設定を自動取得）
 * ============================================================ */
function メニュー_確定発行() {
  const sh = SpreadsheetApp.getActive().getActiveSheet();

  // 台湾グッズは専用の確定発行
  if (sh.getName() === '台湾グッズ') {
    台湾グッズ_確定発行();
    return;
  }

  // ★ 追加
  if (sh.getName() === '韓国グッズ') {
    韓国グッズ_確定発行();
    return;
  }

  const cfg = アクティブシートの設定を取得_();
  if (!cfg) return;
  確定発行を実行(cfg);
}


function メニュー_削除() {
  const cfg = アクティブシートの設定を取得_();
  if (!cfg) return;
  削除を実行(cfg);
}

function メニュー_一括更新() {
  const cfg = アクティブシートの設定を取得_();
  if (!cfg) return;
  一括更新を実行(cfg);
}

function メニュー_プルダウン更新() {
  const sh = SpreadsheetApp.getActive().getActiveSheet();

  if (sh.getName() === '台湾グッズ') {
    台湾グッズ_プルダウン更新();
    return;
  }

  if (sh.getName() === '韓国グッズ') {
    韓国グッズ_プルダウン更新();
    return;
  }

  const cfg = アクティブシートの設定を取得_();
  if (!cfg) return;
  プルダウン更新を実行(cfg);
}



function メニュー_Works初期化() {
  const cfg = アクティブシートの設定を取得_();
  if (!cfg) return;
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('警告', 'Worksを全削除します。続行？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(cfg.作品シート名);
  if (sh) { const last = sh.getLastRow(); if (last > 1) sh.deleteRows(2, last - 1); }
  else { sh = ss.insertSheet(cfg.作品シート名); }
  sh.getRange(1, 1, 1, cfg.作品列数).setValues([cfg.作品ヘッダー]);
  ui.alert('✅ Works初期化完了');
}

function メニュー_WorksKey再正規化() {
  const cfg = アクティブシートの設定を取得_();
  if (!cfg) return;
  const ui = SpreadsheetApp.getUi();
  const 作品シート = SpreadsheetApp.getActive().getSheetByName(cfg.作品シート名);
  if (!作品シート || 作品シート.getLastRow() < 2) { ui.alert('Worksにデータがありません'); return; }
  if (ui.alert('確認', '全WorksKeyを再計算し重複を統合します。続行？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try {
    const r = WorksKey再正規化を実行(作品シート, cfg);
    ui.alert(`✅ 完了\nキー更新: ${r.キー更新数}\n重複統合: ${r.統合数}\n削除: ${r.削除行数}`);
  } finally { lock.releaseLock(); }
}

function メニュー_重複統合() {
  const cfg = アクティブシートの設定を取得_();
  if (!cfg) return;
  const ui = SpreadsheetApp.getUi();
  const 作品シート = SpreadsheetApp.getActive().getSheetByName(cfg.作品シート名);
  if (!作品シート || 作品シート.getLastRow() < 2) { ui.alert('Worksにデータがありません'); return; }
  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try {
    const 重複リスト = Works重複を検出(作品シート, cfg);
    if (重複リスト.length === 0) { ui.alert('✅ 重複はありません！'); return; }
    let レポート = `🔍 重複検出: ${重複リスト.length}件\n\n`;
    for (const dup of 重複リスト) {
      レポート += `【${dup.行配列[0].データ[2]}】 著：${dup.行配列[0].データ[3]}\n`;
      for (const item of dup.行配列) レポート += `  ID:${item.データ[1]} (行${item.行})\n`;
      レポート += '\n';
    }
    if (ui.alert('重複チェック結果', レポート + '\n自動統合しますか？', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
    const r = WorksKey再正規化を実行(作品シート, cfg);
    ui.alert(`✅ 統合完了\n統合: ${r.統合数}件, 削除: ${r.削除行数}行`);
  } finally { lock.releaseLock(); }
}

function メニュー_ID振り直し() {
  const cfg = アクティブシートの設定を取得_();
  if (!cfg) return;
  const ui = SpreadsheetApp.getUi();
  const 作品シート = SpreadsheetApp.getActive().getSheetByName(cfg.作品シート名);
  if (!作品シート || 作品シート.getLastRow() < 2) { ui.alert('Worksにデータがありません'); return; }
  if (ui.alert('確認', 'Works IDを1から連番に振り直します。続行？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try {
    const r = WorksID振り直しを実行(作品シート, cfg);
    ui.alert(`✅ ID振り直し完了\n変更: ${r.変更数}件`);
  } finally { lock.releaseLock(); }
}

function メニュー_孤立削除() {
  const cfg = アクティブシートの設定を取得_();
  if (!cfg) return;
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const 作品シート = ss.getSheetByName(cfg.作品シート名);
  const マスターシート = ss.getSheetByName(cfg.マスターシート名);
  if (!作品シート || 作品シート.getLastRow() < 2) { ui.alert('Worksにデータがありません'); return; }
  if (!マスターシート) { ui.alert('マスターシートがありません'); return; }
  if (ui.alert('確認', '孤立エントリーを削除します。続行？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try {
    const 列マップ = 列番号を取得(マスターシート);
    const r = Works孤立エントリーを削除(作品シート, マスターシート, 列マップ, cfg);
    ui.alert(r.削除数 === 0 ? '✅ 孤立エントリーはありません！' : `✅ 削除完了: ${r.削除数}件\n\n${r.削除リスト.slice(0, 20).join('\n')}`);
  } finally { lock.releaseLock(); }
}

function メニュー_マスター作成() {
  // マスターシートは全シート共通のため台湾書籍設定を使用
  const cfg = 設定_台湾書籍;
  const ss = SpreadsheetApp.getActive();
  const created = [];
  const 言語初期値 = [
    ['台湾', 'TW', '#b7e1cd'], ['韓国', 'KR', '#fce8b2'], ['日本', 'JP', '#f4c7c3'],
    ['中国', 'CN', '#c9daf8'], ['タイ', 'TH', '#d9d2e9'], ['英語', 'US', '#fce5cd']
  ];
  const カテゴリ初期値 = [
    ['コミック', 'CM', '#fff2cc'], ['まんが', 'CM', '#fff2cc'], ['小説', 'NV', '#d9ead3'],
    ['グッズ', 'GD', '#cfe2f3'], ['設定集', 'ART', '#f4cccc'], ['アートブック', 'ART', '#f4cccc'],
    ['雑誌', 'MZ', '#d9d2e9']
  ];

  if (!ss.getSheetByName(cfg.言語マスター名)) {
    const sh = ss.insertSheet(cfg.言語マスター名);
    sh.getRange(1, 1, 1, 3).setValues([['言語', 'コード', '色']]);
    sh.getRange(2, 1, 言語初期値.length, 3).setValues(言語初期値);
    created.push('言語マスター');
  }
  if (!ss.getSheetByName(cfg.カテゴリマスター名)) {
    const sh = ss.insertSheet(cfg.カテゴリマスター名);
    sh.getRange(1, 1, 1, 3).setValues([['カテゴリ', 'コード', '色']]);
    sh.getRange(2, 1, カテゴリ初期値.length, 3).setValues(カテゴリ初期値);
    created.push('カテゴリマスター');
  }

  if (created.length > 0) SpreadsheetApp.getUi().alert(`✅ 作成: ${created.join(', ')}`);
  else SpreadsheetApp.getUi().alert('既に全て存在します');
}

function メニュー_韓国マンガシート作成() {
  韓国マンガシートを作成(); // 韓国マンガ.gs に定義
}

/* ============================================================
 * 内部ヘルパー
 * ============================================================ */
function アクティブシートの設定を取得_() {
  const sh = SpreadsheetApp.getActive().getActiveSheet();
  if (!sh) return null;
  const cfg = シート設定を取得(sh.getName());
  if (!cfg) {
    SpreadsheetApp.getUi().alert(`「${sh.getName()}」はこの操作の対象外です。\n対象シートをアクティブにしてから実行してください。`);
    return null;
  }
  return cfg;
}

function メニュー_韓国書籍シート作成() {
  韓国書籍シートを作成(); // 韓国書籍シート作成.gs に定義
}

function メニュー_韓国音楽映像シート作成() {
  韓国音楽映像シートを作成();
}

function メニュー_韓国グッズシート作成()      { 韓国グッズシートを作成(); }
function メニュー_韓国グッズWorksシート作成()  { 韓国グッズWorksシートを作成(); }
function メニュー_韓国グッズ重複チェック()     { 韓国グッズ_重複チェック(); }
function メニュー_韓国グッズWorks初期化()      { 韓国グッズ_Works初期化(); }
function メニュー_グッズジャンルマスター作成() { グッズジャンルマスターを作成(); }


function メニュー_ショップマスター作成() {
  ショップマスターを作成();
}
function メニュー_グッズ作品マスター作成() {
  グッズ作品マスターを作成();
}

function メニュー_台湾グッズシート作成()        { 台湾グッズシートを作成(); }
function メニュー_台湾グッズWorksシート作成()    { 台湾グッズWorksシートを作成(); }
function メニュー_台湾作品略称マスター作成()     { 台湾作品略称マスターを作成(); }
function メニュー_台湾グッズ確定発行()           { 台湾グッズ_確定発行(); }
function メニュー_台湾グッズ重複チェック()       { 台湾グッズ_重複チェック(); }
function メニュー_台湾グッズWorks初期化()        { 台湾グッズ_Works初期化(); }

// --- 追加するラッパー関数 ---

function メニュー_アラジンTTBキー設定() {
  アラジンTTBキーを設定();
}

function メニュー_アラジン一括取得() {
  アラジン一括取得();
}

function メニュー_アラジン列追加() {
  アラジン列追加();
}

function メニュー_アラジントリガー列に色付け() {
  アラジントリガー列に色付け();
}

function 設定取得テスト() {
  const cfg = シート設定を取得('韓国雑誌');
  Logger.log(cfg);
}