/**
 * Onopen.gs（台湾ファイル専用）
 * メニュー作成・シート設定マップ・onOpen・onEdit
 */

const DEBUG_MODE = false; // ← ここ！

function 台湾まんが_ONEDIT_LOG_(label, payload) {
  const sheet = payload && payload.sheet;
  if (sheet && sheet !== '台湾まんが') return;

  let body = '';
  try {
    body = JSON.stringify(payload || {});
  } catch (err) {
    body = String(payload || '');
  }

  const line = `TW_MANGA_ONEDIT_DEBUG ${label}: ${body}`;
  try { console.log(line); } catch (err) {}
  try { Logger.log(line); } catch (err) {}
}

/* ============================================================
 * シート設定取得
 * ============================================================ */
function シート設定を取得(シート名) {
  const map = {
    '台湾まんが': 設定_台湾まんが,
    '台湾書籍その他': 設定_台湾書籍その他,
    '台湾雑誌': typeof 設定_台湾雑誌 !== 'undefined' ? 設定_台湾雑誌 : null,
    '台湾グッズ': typeof 設定_台湾グッズ !== 'undefined' ? 設定_台湾グッズ : null
  };
  return map[シート名] || null;
}

/* ============================================================
 * onOpen メニュー
 * ============================================================ */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('🌏 台湾商品管理');

  menu
    .addSubMenu(
      ui.createMenu('☑ チェック操作')
        .addItem('全チェック', 'チェック_全チェック')
        .addItem('全チェック解除', 'チェック_全解除')
        .addItem('未登録のみチェック', 'チェック_未登録のみ')
    )
    .addSeparator()
    .addItem('📁 フォルダをSKUにリネーム', '全シート_フォルダをSKUにリネーム')
    .addItem('🔍 次の空き作品IDを確認', '作品ID_次の空き番号を確認')
    .addItem('⚡ トリガー再設定（onEdit+onChange）', 'トリガーを設定')

    .addSubMenu(
      ui.createMenu('台湾グッズ')
        .addItem('① 確定発行', '台湾グッズ_確定発行')
        .addItem('② 一括更新', '台湾グッズ_一括更新')
        .addItem('③ 重複チェック', '台湾グッズ_重複チェック')
        .addItem('④ Works更新', '台湾グッズ_Worksを更新')
        .addItem('⑤ プルダウン更新', '台湾グッズ_プルダウン更新')
    )

    .addSubMenu(
      ui.createMenu('台湾まんが')
        .addItem('① 確定発行', '台湾まんが_確定発行')
        .addItem('② チェック行を削除', '台湾まんが_削除')
        .addItem('③ 一括更新', '台湾まんが_一括更新')
        .addItem('④ 選択行を再生成', '台湾まんが_現在行を再計算')
        .addItem('⑥ プルダウン更新', '台湾まんが_プルダウン更新')
        .addSeparator()
        .addItem('🛠 WorksKeyに媒体コード付与（削除なし）', '台湾書籍系_WorksKey媒体付与_削除なし')
        .addItem('🔎 重複作品を検出（ドライラン）', '台湾書籍系_重複作品_検出ドライラン')
        .addItem('🎨 重複検出の色をクリア', '台湾書籍系_重複作品_色をクリア')
        .addItem('🔗 重複作品を統合（実行）', '台湾書籍系_重複作品_統合実行')
        .addItem('🔄 Works最新巻を再計算', '台湾書籍系_Works最新巻を再計算メニュー_')
        .addItem('🔒 Works危険操作は停止中', '台湾書籍系_Works危険操作メニュー説明')
        
    )

    .addSubMenu(
      ui.createMenu('台湾書籍その他')
        .addItem('① 確定発行', '台湾書籍その他_確定発行')
        .addItem('② チェック行を削除', '台湾書籍その他_削除')
        .addItem('③ 一括更新', '台湾書籍その他_一括更新')
        .addItem('④ 選択行を再生成', '台湾書籍その他_現在行を再計算')
        .addItem('⑥ プルダウン更新', '台湾書籍その他_プルダウン更新')
        .addSeparator()
        .addItem('🛠 WorksKeyに媒体コード付与（削除なし）', '台湾書籍系_WorksKey媒体付与_削除なし')
        .addItem('🔎 重複作品を検出（ドライラン）', '台湾書籍系_重複作品_検出ドライラン')
        .addItem('🎨 重複検出の色をクリア', '台湾書籍系_重複作品_色をクリア')
        .addItem('🔗 重複作品を統合（実行）', '台湾書籍系_重複作品_統合実行')
        .addItem('🔄 Works最新巻を再計算', '台湾書籍系_Works最新巻を再計算メニュー_')
        .addItem('🔒 Works危険操作は停止中', '台湾書籍系_Works危険操作メニュー説明')
        
    )

    .addSubMenu(
      ui.createMenu('台湾雑誌')
        .addItem('① 確定発行', '台湾雑誌_確定発行')
        .addItem('⑥ プルダウン更新', '台湾雑誌_プルダウン更新')
        .addSeparator()
        .addItem('📋 候補の件数を確認', '台湾雑誌_候補件数を確認')
        .addItem('✅ 候補を正式マスターへ反映', '台湾雑誌_候補を正式マスターへ反映')
        .addSeparator()
        .addItem('🔄 現在行を再計算', '台湾雑誌_現在行を再計算')
        .addSeparator()
        .addItem('🔧 コード自動生成セットアップ', '台湾雑誌_雑誌コード自動生成セットアップ')
    )

    .addSeparator()

    .addSubMenu(
  ui.createMenu('シート作成・初期設定')
    .addItem('台湾グッズシートを作成', '台湾グッズシートを作成')
    .addItem('Works（台湾グッズ）シートを作成', '台湾グッズWorksシートを作成')
    .addItem('作品略称マスター（台湾）を作成', '台湾グッズ作品マスターを作成')
    .addSeparator()
    .addItem('台湾まんがシートを作成', '台湾まんがシートを作成')
    .addItem('台湾書籍その他シートを作成', '台湾書籍その他シートを作成')
    .addItem('Works（書籍専用）シートを作成', '台湾書籍Worksシートを作成')
    .addSeparator()
    .addItem('🔢 作品IDを4桁に統一', '台湾書籍系_作品IDを4桁に統一_')
    .addSeparator()
    .addItem('雑誌マスター（台湾）を作成', '台湾雑誌マスターを作成')
    .addItem('台湾雑誌シートを作成', '台湾雑誌シートを作成')
    .addItem('⚠️ 作品ID衝突チェック', '台湾書籍系_作品ID衝突チェック_')
    .addItem('🔢 Works作品IDを4桁文字列に統一', '台湾書籍系_Works作品IDを4桁文字列に統一_')
)
    .addSeparator()
    .addItem('🔧 トリガーを再設定（重複削除）', 'トリガーを設定')

    .addSeparator()
    .addItem('📤 Yahooテンプレへ送信（チェック行）', '台湾_Yahooテンプレへ送信')
    .addItem('🔎 Yahoo送信 構造を確認', '台湾_Yahoo送信_構造を確認')

    .addToUi();

  ui.createMenu('💴 計算機')
    .addItem('🧮 電卓を開く', '計算機サイドバーを開く')
    .addItem('✨ きれいな計算機シートを作成', '計算機_きれいなシートを作成')
    .addToUi();
}

/* ============================================================
 * onEdit トリガー
 * ============================================================ */
function onEditInstallable_(e) {
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
  let シート名 = '';

  try {
    if (!e || !e.range) return;

    シート名 = String(e.range.getSheet().getName() || '').trim();

    // ★ ログシートは完全に無視
    if (シート名 === '_DEBUG_LOG') return;

    台湾まんが_ONEDIT_LOG_('onEditInstallable_:entry', {
      sheet: シート名,
      row: e.range.getRow(),
      column: e.range.getColumn(),
      a1: e.range.getA1Notation(),
      value: e.value !== undefined ? e.value : null,
      oldValue: e.oldValue !== undefined ? e.oldValue : null
    });

    デバッグログ出力_('onEditInstallable_start', {
      ts,
      hasE: !!e,
      hasRange: !!(e && e.range),
      hasSource: !!(e && e.source)
    });

    let ss = null;
    try {
      ss = (e && e.source) ? e.source : SpreadsheetApp.getActiveSpreadsheet();
    } catch (err) {
      デバッグログ出力_('spreadsheet_get_error', { message: String(err) });
    }

    if (DEBUG_MODE) {
  try {
    if (ss) ss.toast(`onEdit発火 ${ts}`, 'トリガー確認', 3);
  } catch (err) {
    デバッグログ出力_('toast_error', { message: String(err) });
  }
}

    デバッグログ出力_('onEditInstallable_router', {
      sheet: シート名,
      row: e.range.getRow(),
      col: e.range.getColumn(),
      value: e.value !== undefined ? e.value : null,
      oldValue: e.oldValue !== undefined ? e.oldValue : null
    });

    if (シート名 === '台湾グッズ') {
      デバッグログ出力_('go_台湾グッズ_onEdit', {});
      台湾グッズ_onEdit(e);
      return;
    }

    if (シート名 === '台湾まんが') {
      デバッグログ出力_('go_台湾まんが_onEdit', {});
      台湾まんが_onEdit(e);
      return;
    }

    if (シート名 === '台湾書籍その他') {
      デバッグログ出力_('go_台湾書籍その他_onEdit', {});
      台湾書籍その他_onEdit(e);
      return;
    }

    if (シート名 === '台湾雑誌') {
      デバッグログ出力_('go_台湾雑誌_onEdit', {});
      台湾雑誌_onEdit(e);
      return;
    }

    デバッグログ出力_('onEditInstallable_no_match', { sheet: シート名 });

  } catch (err) {
    // ★ ログシート中は catch 側も書かない
    if (シート名 !== '_DEBUG_LOG') {
      デバッグログ出力_('onEditInstallable_error', {
        message: String(err),
        stack: err && err.stack ? String(err.stack) : ''
      });
    }
    throw err;
  }
}

/* ============================================================
 * 台湾まんが メニューラッパー
 * ============================================================ */
function 台湾まんが_確定発行() {
  台湾書籍系_チェック行を事前補完_('台湾まんが', 設定_台湾まんが);
  _kyoutuu.確定発行を実行(設定_台湾まんが);
}
function 台湾まんが_削除() { _kyoutuu.削除を実行(設定_台湾まんが); }

// ★ここを _kyoutuu ではなくローカル一括更新に変更
function 台湾まんが_一括更新() {
  台湾書籍系_ローカル一括更新_('台湾まんが', 設定_台湾まんが);
}

function 台湾まんが_プルダウン更新() { _kyoutuu.プルダウン更新を実行(設定_台湾まんが); }
function 台湾まんが_Works初期化() { 台湾書籍系_Works危険操作を停止_('Works初期化'); }
function 台湾まんが_WorksKey再正規化() { 台湾書籍系_Works危険操作を停止_('WorksKey再正規化'); }
function 台湾まんが_重複統合() { 台湾書籍系_Works危険操作を停止_('Works重複チェック・統合'); }
function 台湾まんが_ID振り直し() { 台湾書籍系_Works危険操作を停止_('Works ID振り直し'); }
function 台湾まんが_孤立削除() { 台湾書籍系_Works危険操作を停止_('Works孤立エントリー削除'); }

/* ============================================================
 * 台湾書籍その他 メニューラッパー
 * ============================================================ */
function 台湾書籍その他_確定発行() {
  台湾書籍系_チェック行を事前補完_('台湾書籍その他', 設定_台湾書籍その他);
  _kyoutuu.確定発行を実行(設定_台湾書籍その他);
}
function 台湾書籍その他_削除() { _kyoutuu.削除を実行(設定_台湾書籍その他); }

// ★ここを _kyoutuu ではなくローカル一括更新に変更
function 台湾書籍その他_一括更新() {
  台湾書籍系_ローカル一括更新_('台湾書籍その他', 設定_台湾書籍その他);
}

function 台湾書籍その他_プルダウン更新() { _kyoutuu.プルダウン更新を実行(設定_台湾書籍その他); }
function 台湾書籍その他_Works初期化() { 台湾書籍系_Works危険操作を停止_('Works初期化'); }
function 台湾書籍その他_WorksKey再正規化() { 台湾書籍系_Works危険操作を停止_('WorksKey再正規化'); }
function 台湾書籍その他_重複統合() { 台湾書籍系_Works危険操作を停止_('Works重複チェック・統合'); }
function 台湾書籍その他_ID振り直し() { 台湾書籍系_Works危険操作を停止_('Works ID振り直し'); }
function 台湾書籍その他_孤立削除() { 台湾書籍系_Works危険操作を停止_('Works孤立エントリー削除'); }

function 台湾書籍系_Works危険操作メニュー説明() {
  台湾書籍系_Works危険操作を停止_('Works危険操作');
}

function 台湾書籍系_Works危険操作を停止_(操作名) {
  SpreadsheetApp.getUi().alert(
    '停止中',
    `${操作名} は現在停止しています。\n\n` +
    'この操作はWorks行の統合・削除・ID変更を行う可能性があり、CM/NVを別作品として扱う現在の運用では危険です。\n\n' +
    '必要な場合は、削除なしの専用修復メニューを作ってから実行してください。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/* ============================================================
 * 書籍 Works 補助
 * ============================================================ */
function _kyoutuu書籍Works初期化_(cfg) {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('警告', 'Worksを全削除します。続行？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(cfg.作品シート名);

  if (sh) {
    const last = sh.getLastRow();
    if (last > 1) sh.deleteRows(2, last - 1);
  } else {
    sh = ss.insertSheet(cfg.作品シート名);
  }

  sh.getRange(1, 1, 1, cfg.作品列数).setValues([cfg.作品ヘッダー]);
  ui.alert('✅ Works初期化完了');
}

function _kyoutuu書籍WorksKey再正規化_(cfg) {
  const ui = SpreadsheetApp.getUi();
  const 作品シート = SpreadsheetApp.getActive().getSheetByName(cfg.作品シート名);
  if (!作品シート || 作品シート.getLastRow() < 2) {
    ui.alert('Worksにデータがありません');
    return;
  }

  if (ui.alert('確認', '全WorksKeyを再計算し重複を統合します。続行？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const r = _kyoutuu.WorksKey再正規化を実行(作品シート, cfg);
    ui.alert(`✅ 完了\nキー更新: ${r.キー更新数}\n重複統合: ${r.統合数}\n削除: ${r.削除行数}`);
  } finally {
    lock.releaseLock();
  }
}

function _kyoutuu書籍重複統合_(cfg) {
  const ui = SpreadsheetApp.getUi();
  const 作品シート = SpreadsheetApp.getActive().getSheetByName(cfg.作品シート名);
  if (!作品シート || 作品シート.getLastRow() < 2) {
    ui.alert('Worksにデータがありません');
    return;
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const 重複リスト = _kyoutuu.Works重複を検出(作品シート, cfg);
    if (重複リスト.length === 0) {
      ui.alert('✅ 重複はありません！');
      return;
    }

    let レポート = `🔍 重複検出: ${重複リスト.length}件\n\n`;
    for (const dup of 重複リスト) {
      レポート += `【${dup.行配列[0].データ[2]}】 著：${dup.行配列[0].データ[3]}\n`;
      for (const item of dup.行配列) {
        レポート += `  ID:${item.データ[1]} (行${item.行})\n`;
      }
      レポート += '\n';
    }

    if (ui.alert('重複チェック結果', レポート + '\n自動統合しますか？', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

    const r = _kyoutuu.WorksKey再正規化を実行(作品シート, cfg);
    ui.alert(`✅ 統合完了\n統合: ${r.統合数}件, 削除: ${r.削除行数}行`);
  } finally {
    lock.releaseLock();
  }
}

function _kyoutuu書籍ID振り直し_(cfg) {
  const ui = SpreadsheetApp.getUi();
  const 作品シート = SpreadsheetApp.getActive().getSheetByName(cfg.作品シート名);
  if (!作品シート || 作品シート.getLastRow() < 2) {
    ui.alert('Worksにデータがありません');
    return;
  }

  if (ui.alert('確認', 'Works IDを1から連番に振り直します。続行？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const r = _kyoutuu.WorksID振り直しを実行(作品シート, cfg);
    ui.alert(`✅ ID振り直し完了\n変更: ${r.変更数}件`);
  } finally {
    lock.releaseLock();
  }
}

function _kyoutuu書籍孤立削除_(cfg) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const 作品シート = ss.getSheetByName(cfg.作品シート名);
  const マスターシート = ss.getSheetByName(cfg.マスターシート名);

  if (!作品シート || 作品シート.getLastRow() < 2) {
    ui.alert('Worksにデータがありません');
    return;
  }
  if (!マスターシート) {
    ui.alert('マスターシートがありません');
    return;
  }

  if (ui.alert('確認', '孤立エントリーを削除します。続行？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const 列マップ = _kyoutuu.列番号を取得(マスターシート);
    const r = _kyoutuu.Works孤立エントリーを削除(作品シート, マスターシート, 列マップ, cfg);
    ui.alert(r.削除数 === 0 ? '✅ 孤立エントリーはありません！' : `✅ 削除完了: ${r.削除数}件`);
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
 * チェックボックス一括操作
 * ============================================================ */
function チェック_全チェック() { _チェック操作(true, 'all'); }
function チェック_全解除() { _チェック操作(false, 'all'); }
function チェック_未登録のみ() { _チェック操作(true, 'unregistered'); }

function _チェック操作(チェック値, モード) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const 列マップ = _kyoutuu.列番号を取得(sh);
  const チェック列 = 列マップ['発番発行'];
  const 登録状況列 = 列マップ['登録状況'];

  if (!チェック列 || !登録状況列) {
    SpreadsheetApp.getUi().alert('このシートはチェック操作に対応していません');
    return;
  }

  const 全行 = sh.getRange(2, 登録状況列, sh.getLastRow() - 1, 1).getValues();
  let 最終行 = 1;
  for (let i = 全行.length - 1; i >= 0; i--) {
    if (全行[i][0] !== '') {
      最終行 = i + 2;
      break;
    }
  }
  if (最終行 < 2) return;

  const 登録状況 = sh.getRange(2, 登録状況列, 最終行 - 1, 1).getValues();
  const updates = 登録状況.map(([v]) => {
    const 未登録 = !String(v || '').trim().startsWith('登録済');
    if (モード === 'all') return [チェック値];
    return [未登録 ? true : false];
  });

  sh.getRange(2, チェック列, 最終行 - 1, 1).setValues(updates);
}

/* ============================================================
 * トリガー設定
 * ============================================================ */
function トリガーを設定() {
  const ss = SpreadsheetApp.getActive();
  const handlers = [
    'onEdit',
    'onEditInstallable_',
    'onEditInstallable_test_',
    'onChangeInstallable_',
    'fixRowHeightOnEdit'
  ];
  let 削除数 = 0;

  ScriptApp.getProjectTriggers()
    .filter(t => handlers.includes(t.getHandlerFunction()))
    .forEach(t => {
      ScriptApp.deleteTrigger(t);
      削除数++;
    });

  ScriptApp.newTrigger('onEditInstallable_')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  // 「上から同じ値を入れる（フィルダウン/貼り付け）」は値が変わらず onEdit が発火しないため、
  // onChange でも変更範囲を拾って再生成する。
  ScriptApp.newTrigger('onChangeInstallable_')
    .forSpreadsheet(ss)
    .onChange()
    .create();

  if (typeof fixRowHeightOnEdit === 'function') {
    ScriptApp.newTrigger('fixRowHeightOnEdit')
      .forSpreadsheet(ss)
      .onEdit()
      .create();
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `トリガー設定完了: onEdit=1 / onChange=1 / 行高さ=1 / 削除=${削除数}`,
    '完了',
    5
  );
}

/* ============================================================
 * onChange トリガー
 * フィルダウン/貼り付けなど、値が変わらず onEdit が発火しないケースを拾う
 * （onChange は Apps Script 自身の変更では発火しないため、ループにはならない）
 * ============================================================ */
function onChangeInstallable_(e) {
  try {
    const changeType = e ? String(e.changeType || '') : '';
    // 値が入りうる変更だけ対象（書式変更・行挿入などは無視）
    if (['EDIT', 'PASTE', 'OTHER'].indexOf(changeType) === -1) return;

    const ss = (e && e.source) ? e.source : SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;

    const sh = ss.getActiveSheet();
    const シート名 = sh ? String(sh.getName() || '').trim() : '';
    if (シート名 === '_DEBUG_LOG') return;

    const 設定マップ = 台湾書籍系_保留設定マップ_();
    const 設定 = 設定マップ[シート名];
    if (!設定) return; // 台湾まんが / 台湾書籍その他 以外は無視

    const rng = ss.getActiveRange();
    if (!rng) return;

    let 開始行 = rng.getRow();
    let 行数 = rng.getNumRows();
    if (開始行 < 2) {
      行数 -= (2 - 開始行);
      開始行 = 2;
    }
    if (行数 <= 0) return;

    if (typeof 台湾まんが_ONEDIT_LOG_ === 'function') {
      台湾まんが_ONEDIT_LOG_('onChange:enter', {
        sheet: シート名,
        changeType,
        startRow: 開始行,
        rows: 行数
      });
    }

    // 大量範囲は保留キューに積んで時間トリガーで補完（タイムアウト/取りこぼし対策）。
    if (行数 >= 10) {
      台湾書籍系_保留に追加_(シート名, 開始行, 行数);
      台湾書籍系_保留補完トリガーを確保_();
      return;
    }

    // 少量範囲はその場で再生成。失敗時は保留キューにフォールバック。
    try {
      台湾書籍系_新取込後補完_共通_(シート名, 開始行, 行数, 設定);
    } catch (procErr) {
      台湾書籍系_保留に追加_(シート名, 開始行, 行数);
      台湾書籍系_保留補完トリガーを確保_();
      デバッグログ出力_('onChangeInstallable_処理エラー', { message: String(procErr) });
    }
  } catch (err) {
    デバッグログ出力_('onChangeInstallable_エラー', { message: String(err) });
  }
}

/* ============================================================
 * 再計算補助
 * ============================================================ */
function 台湾まんが_再計算(開始行, 行数) {
  const sh = SpreadsheetApp.getActive().getSheetByName('台湾まんが');
  if (!sh || 行数 <= 0) return;

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  // ループ前に数式を確定し、必要な情報と行データを一括取得
  SpreadsheetApp.flush();

  const 列 = 台湾書籍系_列マップを取得_(sh);
  const lastCol = sh.getLastColumn();
  const allRowValues = sh.getRange(開始行, 1, 行数, lastCol).getValues();

  try {
    for (let r = 開始行; r < 開始行 + 行数; r++) {
      if (r < 2) continue;
      const rowValues = allRowValues[r - 開始行];
      if (!rowValues) continue;

      台湾書籍系_1行補完_共通_(sh, r, 設定_台湾まんが, {
        skipFlush: true,
        列マップ: 列,
        lastCol: lastCol,
        rowValues: rowValues
      });
    }
  } finally {
    lock.releaseLock();
  }
}

function 台湾書籍その他_再計算(開始行, 行数) {
  const sh = SpreadsheetApp.getActive().getSheetByName('台湾書籍その他');
  if (!sh || 行数 <= 0) return;

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  // ループ前に数式を確定し、必要な情報と行データを一括取得
  SpreadsheetApp.flush();

  const 列 = 台湾書籍系_列マップを取得_(sh);
  const lastCol = sh.getLastColumn();
  const allRowValues = sh.getRange(開始行, 1, 行数, lastCol).getValues();

  try {
    for (let r = 開始行; r < 開始行 + 行数; r++) {
      if (r < 2) continue;
      const rowValues = allRowValues[r - 開始行];
      if (!rowValues) continue;

      台湾書籍系_1行補完_共通_(sh, r, 設定_台湾書籍その他, {
        skipFlush: true,
        列マップ: 列,
        lastCol: lastCol,
        rowValues: rowValues
      });
    }
  } finally {
    lock.releaseLock();
  }
}


/* ============================================================
 * 台湾書籍系 ローカル一括更新
 * _kyoutuu.一括更新ではなく、
 * 台湾書籍系_1行補完_共通_() を直接通す
 * ============================================================ */
function 台湾書籍系_ローカル一括更新_(シート名, 設定) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(シート名);

  if (!sh || sh.getLastRow() < 2) {
    ui.alert(`${シート名} にデータがありません`);
    return;
  }

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
    .getDisplayValues()[0]
    .map(v => String(v || '').trim());

  const 登録状況列 = headers.indexOf('登録状況') + 1;

  if (!登録状況列) {
    ui.alert('登録状況列が見つかりません');
    return;
  }

  const lastRow = sh.getLastRow();
  const statuses = sh.getRange(2, 登録状況列, lastRow - 1, 1).getDisplayValues();

  const 対象行 = [];

  statuses.forEach((r, i) => {
    const 状態 = String(r[0] || '').trim();

    // 現場運用では登録済みは触らない
    if (状態 === '未登録') {
      対象行.push(i + 2);
    }
  });

  if (対象行.length === 0) {
    ui.alert('未登録の対象行がありません');
    return;
  }

  if (
    ui.alert(
      '確認',
      `${シート名} の未登録 ${対象行.length} 行を一括更新します。\n\n` +
      '・作品IDの4桁統一\n' +
      '・古い作品IDの誤引き継ぎチェック\n' +
      '・Works取得または新規作成\n' +
      '・親コード / SKU / タイトル再生成\n\n' +
      '続行しますか？',
      ui.ButtonSet.OK_CANCEL
    ) !== ui.Button.OK
  ) {
    return;
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  let 成功 = 0;
  let 失敗 = 0;
  const エラー = [];

  // ループ前に数式を確定し、必要な情報とシートデータを一括取得
  SpreadsheetApp.flush();

  const 列 = 台湾書籍系_列マップを取得_(sh);
  const lastCol = sh.getLastColumn();
  const allSheetValues = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  try {
    // 先にWorksと商品シート側の作品ID表記を4桁に寄せる
    if (typeof 台湾書籍系_作品IDを4桁に統一_ === 'function') {
      台湾書籍系_作品IDを4桁に統一_();
    }


    対象行.forEach(row => {
      try {
        デバッグログ出力_('一括更新_行開始', {
          シート名,
          row
        });

        const rowValues = allSheetValues[row - 2];

        台湾書籍系_1行補完_共通_(sh, row, 設定, {
          skipFlush: true,
          列マップ: 列,
          lastCol: lastCol,
          rowValues: rowValues
        });

        デバッグログ出力_('一括更新_行完了', {
          シート名,
          row
        });

        成功++;
      } catch (err) {
        失敗++;
        エラー.push(`${row}行目: ${err && err.message ? err.message : err}`);

        デバッグログ出力_('一括更新_行エラー', {
          シート名,
          row,
          message: err && err.message ? err.message : String(err),
          stack: err && err.stack ? String(err.stack) : ''
        });
      }
    });

    SpreadsheetApp.flush();

    let msg =
      `✅ ${シート名}：未登録 ${対象行.length} 行を一括更新しました\n\n` +
      `成功: ${成功}件\n` +
      `失敗: ${失敗}件`;

    if (エラー.length) {
      msg += '\n\nエラー:\n' + エラー.slice(0, 10).join('\n');
      if (エラー.length > 10) msg += `\nほか ${エラー.length - 10}件`;
    }

    ui.alert(msg);

  } finally {
    lock.releaseLock();
  }
}
function 台湾まんが_現在行を再計算() {
  台湾書籍系_現在行をローカル再計算_('台湾まんが', 設定_台湾まんが);
}

function 台湾書籍その他_現在行を再計算() {
  台湾書籍系_現在行をローカル再計算_('台湾書籍その他', 設定_台湾書籍その他);
}

function 台湾書籍系_現在行をローカル再計算_(シート名, 設定) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();

  if (!sh || sh.getName() !== シート名) {
    SpreadsheetApp.getUi().alert(`${シート名} シートで実行してください`);
    return;
  }

  const range = sh.getActiveRange();
  const startRow = range ? range.getRow() : sh.getActiveCell().getRow();
  const rowCount = range ? range.getNumRows() : 1;
  const endRow = Math.min(sh.getLastRow(), startRow + rowCount - 1);

  if (endRow < 2) {
    SpreadsheetApp.getUi().alert('2行目以降で実行してください');
    return;
  }

  const actualStart = Math.max(2, startRow);
  const actualRows = endRow - actualStart + 1;
  const lastCol = sh.getLastColumn();
  const values = sh.getRange(actualStart, 1, actualRows, lastCol).getValues();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    if (typeof 台湾書籍系_使用済みIDキャッシュ無効化_ === 'function') {
      台湾書籍系_使用済みIDキャッシュ無効化_();
    }

    const 列 = 台湾書籍系_列マップを取得_(sh);
    for (let i = 0; i < actualRows; i++) {
      台湾書籍系_1行補完_共通_(sh, actualStart + i, 設定, {
        Works新規作成: true,
        skipFlush: true,
        列マップ: 列,
        lastCol: lastCol,
        rowValues: values[i]
      });
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  const label = actualRows === 1
    ? `${actualStart}行目`
    : `${actualStart}-${endRow}行目`;
  ss.toast(`✅ ${シート名} ${label}を再生成しました`, '完了', 4);
}

function 台湾書籍系_作品IDを4桁に統一_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const 対象 = [
    {
      シート名: 'Works（書籍専用）',
      候補列名: ['作品ID']
    },
    {
      シート名: '台湾まんが',
      候補列名: ['作品ID(W)（自動）', '作品ID(W)(自動)', '作品(W)（自動）']
    },
    {
      シート名: '台湾書籍その他',
      候補列名: ['作品ID(W)(自動)', '作品ID(W)（自動）', '作品(W)（自動）']
    }
  ];

  let 更新数 = 0;

  対象.forEach(info => {
    const sh = ss.getSheetByName(info.シート名);
    if (!sh || sh.getLastRow() < 2) return;

    const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
      .getDisplayValues()[0]
      .map(v => String(v || '').trim());

    let col = 0;
    for (const name of info.候補列名) {
      const idx = headers.indexOf(name);
      if (idx >= 0) {
        col = idx + 1;
        break;
      }
    }

    if (!col) return;

    const lastRow = sh.getLastRow();
    const range = sh.getRange(2, col, lastRow - 1, 1);
    const values = range.getDisplayValues();

    const out = values.map(([v]) => {
      const id = 台湾書籍系_作品ID4桁文字列_(v);
      if (!id) return [''];
      if (String(v || '').trim() !== id) 更新数++;
      return [id];
    });

    // 文字列として固定する
    range.setNumberFormat('@');
    range.setValues(out);
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `✅ 作品IDを4桁に統一しました：${更新数}件`,
    '完了',
    5
  );
}

function 台湾書籍系_作品ID4桁文字列_(v) {
  const s = String(v == null ? '' : v).trim();
  if (!s) return '';

  const m = s.match(/\d+/);
  if (!m) return '';

  const n = parseInt(m[0], 10);
  if (isNaN(n)) return '';

  return String(n).padStart(4, '0');
}

function デバッグログ出力_(tag, obj) {
  const DEBUG_TRACE = false; // 追跡中だけ true。終わったら false にする。

  if (!DEBUG_TRACE) return;

  try {
    const json = obj ? JSON.stringify(obj) : '';
    console.log(`[${tag}] ${json}`);
  } catch (err) {
    console.log(`[${tag}] ログ出力エラー: ${err}`);
  }
}

function 台湾書籍系_作品ID使用場所チェック_0088() {
  台湾書籍系_作品ID使用場所チェック_('0088');
}

function 台湾書籍系_作品ID使用場所チェック_(targetId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const target = String(targetId || '').replace(/[^\d]/g, '').padStart(4, '0');

  const result = [];

  const コードからID = (v) => {
    const s = String(v == null ? '' : v).trim().toUpperCase();
    if (!s) return '';

    // 例: TW0088-CM-01 / TWS0088-CM-0304
    const m = s.match(/^[A-Z]{2}[A-Z]*?(\d{4})-/);
    return m ? m[1] : '';
  };

  const 正規ID = (v) => {
    const s = String(v == null ? '' : v).trim();
    if (!s) return '';

    const m = s.match(/\d+/);
    if (!m) return '';

    return String(parseInt(m[0], 10)).padStart(4, '0');
  };

  const checkSheet = (sheetName, cols) => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return;

    const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
      .getDisplayValues()[0]
      .map(v => String(v || '').trim());

    cols.forEach(colName => {
      const col = headers.indexOf(colName) + 1;
      if (!col) return;

      const values = sh.getRange(2, col, sh.getLastRow() - 1, 1).getDisplayValues();

      values.forEach(([v], i) => {
        const row = i + 2;
        const id1 = 正規ID(v);
        const id2 = コードからID(v);

        if (id1 === target || id2 === target) {
          result.push({
            シート: sheetName,
            行: row,
            列: colName,
            値: v
          });
        }
      });
    });
  };

  checkSheet('Works（書籍専用）', [
    '作品ID'
  ]);

  checkSheet('台湾まんが', [
    '作品ID(W)（自動）',
    '作品ID(W)(自動)',
    '作品(W)（自動）',
    '親コード',
    '商品コード(SKU)',
    'SKU（自動）',
    'SKU(自動)'
  ]);

  checkSheet('台湾書籍その他', [
    '作品ID(W)（自動）',
    '作品ID(W)(自動)',
    '作品(W)（自動）',
    '親コード',
    '商品コード(SKU)',
    'SKU（自動）',
    'SKU(自動)'
  ]);

  if (result.length === 0) {
    console.log(`✅ ${target} は見つかりませんでした`);
    SpreadsheetApp.getUi().alert(`${target} は Works / 台湾まんが / 台湾書籍その他 では見つかりませんでした`);
    return;
  }

  console.log(`⚠️ ${target} の使用場所`);
  result.forEach(r => console.log(JSON.stringify(r)));

  SpreadsheetApp.getUi().alert(
    `${target} の使用場所が ${result.length} 件あります。\n\nApps Script の「実行数」ログを確認してください。`
  );
}

function 作品ID0088使用場所チェック() {
  台湾書籍系_作品ID使用場所チェック_('0088');
}

function 列名確認() {
  const sh = SpreadsheetApp.getActive().getSheetByName('台湾まんが');
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  console.log(JSON.stringify(headers));
}
