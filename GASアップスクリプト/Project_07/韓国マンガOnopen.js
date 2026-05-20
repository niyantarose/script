/**
 * Onopen.gs（韓国マンガファイル専用）
 * メニュー作成・シート設定マップ・インストール型onEdit
 */

/* ============================================================
 * シート名 → 設定オブジェクト マップ
 * ============================================================ */
function シート設定を取得(シート名) {
  const マップ = {
    '韓国マンガ': 設定_韓国マンガ,
  };

  return マップ[シート名] || null;
}

/* ============================================================
 * onOpen
 * ============================================================ */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('📚 韓国マンガ管理')

    .addSubMenu(
      ui.createMenu('☑ チェック操作')
        .addItem('全チェック',         'チェック_全チェック')
        .addItem('全チェック解除',     'チェック_全解除')
        .addItem('未登録のみチェック', 'チェック_未登録のみ')
    )

    .addSeparator()

    .addSubMenu(
      ui.createMenu('韓国マンガ')
        .addItem('① 商品コード確定発行', 'メニュー_確定発行')
        .addItem('② チェック行を削除',   'メニュー_削除')
        .addItem('③ 既存データ一括更新', 'メニュー_一括更新')
        .addItem('⑥ プルダウン更新',     'メニュー_プルダウン更新')
        .addSeparator()
        .addItem('🔍 列名チェック',       '韓国マンガ_列名チェック')
        .addItem('🔁 選択行を再生成',     '韓国マンガ_選択行を再生成')
        .addItem('🧹 自己更新フラグ解除', '共通_自己更新フラグ解除')
    )

    .addSubMenu(
      ui.createMenu('アラジン')
        .addItem('🔴 アラジンTTBキー設定',       'メニュー_アラジンTTBキー設定')
        .addItem('🔴 アラジン一括取得',           'メニュー_アラジン一括取得')
        .addItem('🔴 アラジン列追加',             'メニュー_アラジン列追加')
        .addItem('🔴 アラジントリガー列に色付け', 'メニュー_アラジントリガー列に色付け')
    )

    .addSubMenu(
      ui.createMenu('Works管理')
        .addItem('🔍 Works重複チェック・統合', 'メニュー_重複統合')
        .addItem('🔄 WorksKey再正規化',        'メニュー_WorksKey再正規化')
        .addItem('🔢 Works ID振り直し',        'メニュー_ID振り直し')
        .addItem('🧹 Works孤立エントリー削除', 'メニュー_孤立削除')
        .addItem('⚙️ Works初期化',             'メニュー_Works初期化')
    )

    .addSeparator()

    .addSubMenu(
      ui.createMenu('シート作成・初期設定')
        .addItem('📋 韓国マンガシート作成', 'メニュー_韓国マンガシート作成')
        .addItem('⚙️ onEditトリガー設定',   'トリガーを設定')
    )

    .addToUi();
}

/* ============================================================
 * onEdit トリガー
 * ※ Apps Scriptのトリガー設定で
 *    onEdit_インストール型 を「編集時」に設定する
 * ============================================================ */
function onEdit_インストール型(e) {
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  const シート名 = String(sh.getName() || '').trim();

  if (シート名 !== '韓国マンガ') return;

  const 開始行 = e.range.getRow();
  const 行数 = e.range.getNumRows();

  if (開始行 + 行数 - 1 < 2) return;

  const cfg = シート設定を取得(シート名);
  if (!cfg) return;

  const lock = LockService.getDocumentLock();

  try {
    if (!lock.tryLock(5000)) return;

    const 列マップ = _kyoutuu.列番号を取得(sh);

    // 監視列だけ反応させる
    const 編集開始列 = e.range.getColumn();
    const 編集終了列 = e.range.getLastColumn();

    const 監視列番号 = (cfg.監視列 || [])
      .map(name => 列マップ[String(name || '').trim()])
      .filter(Boolean);

    const 対象列が含まれる = 監視列番号.some(col =>
      col >= 編集開始列 && col <= 編集終了列
    );

    if (!対象列が含まれる) return;

    _kyoutuu.onEdit処理を実行(
      e,
      sh,
      cfg,
      列マップ,
      開始行,
      行数
    );

  } catch (err) {
    console.error('韓国マンガ onEdit生成エラー: ' + err.message);
    throw err;
  } finally {
    try {
      lock.releaseLock();
    } catch (_) {}
  }
}

/* ============================================================
 * トリガー設定
 * ============================================================ */
function トリガーを設定() {
  const ss = SpreadsheetApp.getActive();

  ScriptApp.getProjectTriggers()
    .filter(t => [
      'onEdit',
      'onEdit_インストール型',
      'onEditInstallable_'
    ].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('onEdit_インストール型')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `✅ onEditトリガー設定完了: ${ss.getName()}`,
    '完了',
    5
  );
}

/* ============================================================
 * メニューラッパー
 * ============================================================ */
function メニュー_確定発行()       { 韓国マンガ_確定発行(); }
function メニュー_削除()           { 韓国マンガ_削除(); }
function メニュー_一括更新()       { 韓国マンガ_一括更新(); }
function メニュー_プルダウン更新() { 韓国マンガ_プルダウン更新(); }

function メニュー_Works初期化()       { 韓国マンガ_Works初期化(); }
function メニュー_重複統合()          { 韓国マンガ_重複統合(); }
function メニュー_WorksKey再正規化()  { 韓国マンガ_WorksKey再正規化(); }
function メニュー_ID振り直し()        { 韓国マンガ_ID振り直し(); }
function メニュー_孤立削除()          { 韓国マンガ_孤立削除(); }

function メニュー_韓国マンガシート作成() {
  韓国マンガシートを作成();
}

/* ============================================================
 * アラジン メニューラッパー
 * 関数が存在しない場合でも落ちないようにする
 * ============================================================ */
function メニュー_アラジンTTBキー設定() {
  if (typeof アラジンTTBキーを設定 === 'function') {
    アラジンTTBキーを設定();
  } else {
    SpreadsheetApp.getUi().alert('アラジンTTBキーを設定 関数が見つかりません');
  }
}

function メニュー_アラジン一括取得() {
  if (typeof アラジン一括取得 === 'function') {
    アラジン一括取得();
  } else {
    SpreadsheetApp.getUi().alert('アラジン一括取得 関数が見つかりません');
  }
}

function メニュー_アラジン列追加() {
  if (typeof アラジン列追加 === 'function') {
    アラジン列追加();
  } else {
    SpreadsheetApp.getUi().alert('アラジン列追加 関数が見つかりません');
  }
}

function メニュー_アラジントリガー列に色付け() {
  if (typeof アラジントリガー列に色付け === 'function') {
    アラジントリガー列に色付け();
  } else {
    SpreadsheetApp.getUi().alert('アラジントリガー列に色付け 関数が見つかりません');
  }
}

/* ============================================================
 * チェックボックス一括操作
 * ============================================================ */
function チェック_全チェック() {
  _チェック操作(true, 'all');
}

function チェック_全解除() {
  _チェック操作(false, 'all');
}

function チェック_未登録のみ() {
  _チェック操作(true, 'unregistered');
}

function _チェック操作(チェック値, モード) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const ヘッダー = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const 列 = {};
  ヘッダー.forEach((h, i) => {
    if (h) 列[String(h).trim()] = i + 1;
  });

  const チェック列 = 列['発番発行'];
  const 登録状況列 = 列['登録状況'];

  if (!チェック列 || !登録状況列) {
    SpreadsheetApp.getUi().alert('このシートはチェック操作に対応していません');
    return;
  }

  if (sh.getLastRow() < 2) return;

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

    if (モード === 'all') {
      return [チェック値];
    }

    return [未登録 ? true : false];
  });

  sh.getRange(2, チェック列, 最終行 - 1, 1).setValues(updates);
}

/* ============================================================
 * 列名チェック
 * ============================================================ */
function 韓国マンガ_列名チェック() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(設定_韓国マンガ.マスターシート名);

  if (!sh) {
    SpreadsheetApp.getUi().alert('韓国マンガシートが見つかりません');
    return;
  }

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
    .getValues()[0]
    .map(v => String(v || '').trim());

  const 必要列 = [
    ...(設定_韓国マンガ.監視列 || []),
    ...Object.values(設定_韓国マンガ.列名 || {}),
    ...(設定_韓国マンガ.商品重複キー列優先順位 || [])
  ].filter(Boolean);

  const 不足列 = [...new Set(必要列)]
    .filter(name => !headers.includes(name));

  if (不足列.length) {
    SpreadsheetApp.getUi().alert(
      '❌ 不足している列があります\n\n' + 不足列.join('\n')
    );
  } else {
    SpreadsheetApp.getUi().alert('✅ 韓国マンガ：列名は設定と一致しています');
  }
}

/* ============================================================
 * 選択行を再生成
 * ============================================================ */
function 韓国マンガ_選択行を再生成() {
  const sh = SpreadsheetApp.getActiveSheet();

  if (sh.getName() !== 設定_韓国マンガ.マスターシート名) {
    SpreadsheetApp.getUi().alert('韓国マンガシートで実行してください');
    return;
  }

  const range = sh.getActiveRange();
  const 開始行 = range.getRow();
  const 行数 = range.getNumRows();

  if (開始行 < 2) {
    SpreadsheetApp.getUi().alert('ヘッダー行は再生成できません');
    return;
  }

  _kyoutuu.指定行を再生成(
    設定_韓国マンガ,
    開始行,
    行数
  );

  SpreadsheetApp.getUi().alert(`✅ ${行数}行を再生成しました`);
}

/* ============================================================
 * 自己更新フラグ解除
 * ============================================================ */
function 共通_自己更新フラグ解除() {
  PropertiesService.getScriptProperties().deleteProperty('KYOUTUU_UPDATING');
  SpreadsheetApp.getUi().alert('✅ 自己更新フラグを解除しました');
}