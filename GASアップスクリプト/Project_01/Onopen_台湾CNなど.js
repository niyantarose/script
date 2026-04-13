/**
 * Onopen.gs（台湾ファイル専用）
 * メニュー作成・シート設定マップ・onOpen・onEdit
 */

/* ============================================================
 * 共通ヘルパー
 * ============================================================ */

function シート設定を取得(シート名) {
  const map = {};

  if (typeof 設定_台湾グッズ !== 'undefined') map['台湾グッズ'] = 設定_台湾グッズ;
  if (typeof 設定_台湾まんが !== 'undefined') map['台湾まんが'] = 設定_台湾まんが;
  if (typeof 設定_台湾書籍その他 !== 'undefined') map['台湾書籍その他'] = 設定_台湾書籍その他;
  if (typeof 設定_台湾雑誌 !== 'undefined') map['台湾雑誌'] = 設定_台湾雑誌;

  return map[シート名] || null;
}

function OnopenGS_ヘッダーMap_(sh) {
  const vals = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  const map = {};
  vals.forEach((h, i) => {
    const key = String(h || '').trim();
    if (key) map[key] = i + 1;
  });
  return map;
}

function OnopenGS_関数を取得_(name) {
  try {
    const fn = globalThis[name];
    return typeof fn === 'function' ? fn : null;
  } catch (_) {
    return null;
  }
}

function OnopenGS_関数を呼ぶ_(name, ...args) {
  const fn = OnopenGS_関数を取得_(name);
  if (!fn) return false;
  fn(...args);
  return true;
}

/* ============================================================
 * onOpen
 * ============================================================ */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('🌏 台湾商品管理');

  menu
    .addSubMenu(
      ui.createMenu('☑ チェック操作')
        .addItem('全チェック',         'チェック_全チェック')
        .addItem('全チェック解除',     'チェック_全解除')
        .addItem('未登録のみチェック', 'チェック_未登録のみ')
    )
    .addSeparator()
    .addItem('📁 フォルダをSKUにリネーム', '全シート_フォルダをSKUにリネーム')

    .addSubMenu(
      ui.createMenu('台湾グッズ')
        .addItem('① 確定発行',       '台湾グッズ_確定発行')
        .addItem('② 一括更新',       '台湾グッズ_一括更新')
        .addItem('③ 重複チェック',   '台湾グッズ_重複チェック')
        .addItem('④ Works更新',      '台湾グッズ_Works更新')
        .addItem('⑤ プルダウン更新', '台湾グッズ_プルダウン更新')
    )

    .addSubMenu(
      ui.createMenu('台湾まんが')
        .addItem('① 確定発行',         '台湾まんが_確定発行')
        .addItem('② チェック行を削除', '台湾まんが_削除')
        .addItem('③ 一括更新',         '台湾まんが_一括更新')
        .addItem('⑥ プルダウン更新',   '台湾まんが_プルダウン更新')
        .addSeparator()
        .addItem('🔍 Works重複チェック・統合', '台湾まんが_重複統合')
        .addItem('🔄 WorksKey再正規化',         '台湾まんが_WorksKey再正規化')
        .addItem('🔢 Works ID振り直し',         '台湾まんが_ID振り直し')
        .addItem('🧹 Works孤立エントリー削除', '台湾まんが_孤立削除')
        .addItem('⚙️ Works初期化',             '台湾まんが_Works初期化')
    )

    .addSubMenu(
      ui.createMenu('台湾書籍その他')
        .addItem('① 確定発行',         '台湾書籍その他_確定発行')
        .addItem('② チェック行を削除', '台湾書籍その他_削除')
        .addItem('③ 一括更新',         '台湾書籍その他_一括更新')
        .addItem('⑥ プルダウン更新',   '台湾書籍その他_プルダウン更新')
        .addSeparator()
        .addItem('🔍 Works重複チェック・統合', '台湾書籍その他_重複統合')
        .addItem('🔄 WorksKey再正規化',         '台湾書籍その他_WorksKey再正規化')
        .addItem('🔢 Works ID振り直し',         '台湾書籍その他_ID振り直し')
        .addItem('🧹 Works孤立エントリー削除', '台湾書籍その他_孤立削除')
        .addItem('⚙️ Works初期化',             '台湾書籍その他_Works初期化')
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
        .addItem('台湾まんがシートを作成',           '台湾まんがシートを作成')
        .addItem('台湾書籍その他シートを作成',       '台湾書籍その他シートを作成')
        .addItem('Works（書籍専用）シートを作成',    '台湾書籍Worksシートを作成')
        .addSeparator()
        .addItem('雑誌マスター（台湾）を作成',       '台湾雑誌マスターを作成')
        .addItem('台湾雑誌シートを作成',             '台湾雑誌シートを作成')
    )
    .addToUi();
}

/* ============================================================
 * onEdit トリガー
 * ============================================================ */

function onEditInstallable_(e) {
  console.log('--- onEditInstallable_ start ---');
  console.log(JSON.stringify({
    hasE: !!e,
    hasRange: !!(e && e.range),
    sheet: e && e.range ? e.range.getSheet().getName() : null,
    row: e && e.range ? e.range.getRow() : null,
    col: e && e.range ? e.range.getColumn() : null,
    value: e && e.value !== undefined ? e.value : null,
    oldValue: e && e.oldValue !== undefined ? e.oldValue : null
  }));

  if (!e || !e.range) {
    console.log('return: no e or no range');
    return;
  }

  const シート名 = String(e.range.getSheet().getName() || '').trim();
  console.log('sheet name = [' + シート名 + ']');

  try {
    if (シート名 === '台湾グッズ') {
      console.log('go: 台湾グッズ_onEdit');
      if (!OnopenGS_関数を呼ぶ_('台湾グッズ_onEdit', e)) {
        console.log('handler not found: 台湾グッズ_onEdit');
      }
      return;
    }

    if (シート名 === '台湾まんが') {
      console.log('go: 台湾まんが_onEdit');
      if (!OnopenGS_関数を呼ぶ_('台湾まんが_onEdit', e)) {
        console.log('handler not found: 台湾まんが_onEdit');
      }
      return;
    }

    if (シート名 === '台湾書籍その他') {
      console.log('go: 台湾書籍その他_onEdit');
      if (!OnopenGS_関数を呼ぶ_('台湾書籍その他_onEdit', e)) {
        // 旧互換
        if (!OnopenGS_関数を呼ぶ_('台湾書籍その他_onEdit_処理', e)) {
          console.log('handler not found: 台湾書籍その他_onEdit / 台湾書籍その他_onEdit_処理');
        }
      }
      return;
    }

    if (シート名 === '台湾雑誌') {
      console.log('go: 台湾雑誌_onEdit');
      if (!OnopenGS_関数を呼ぶ_('台湾雑誌_onEdit', e)) {
        console.log('handler not found: 台湾雑誌_onEdit');
      }
      return;
    }

    console.log('return: no matched sheet');
  } catch (err) {
    console.log('onEditInstallable_ error: ' + err);
    throw err;
  }
}

/* ============================================================
 * 旧互換（必要なら残す）
 * ============================================================ */

function 台湾書籍その他_onEdit_処理(e) {
  const cfg = 設定_台湾書籍その他;
  const sh = e.range.getSheet();
  const 開始行 = e.range.getRow();
  const 行数 = e.range.getNumRows();
  if (開始行 + 行数 - 1 < 2) return;

  const 列マップ = _kyoutuu.列番号を取得(sh);
  const 監視列番号 = cfg.監視列.map(h => 列マップ[h]).filter(Boolean);
  const 編集開始列 = e.range.getColumn();
  const 編集終了列 = e.range.getLastColumn();
  if (!監視列番号.some(c => c >= 編集開始列 && c <= 編集終了列)) return;

  const lock = LockService.getDocumentLock();
  try {
    if (!lock.tryLock(5000)) return;
  } catch (_) {
    return;
  }

  try {
    _kyoutuu.onEdit処理を実行(e, sh, cfg, 列マップ, 開始行, 行数);
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
 * 台湾グッズ メニュー補助ラッパー
 * ============================================================ */

function 台湾作品略称マスターを作成() {
  if (OnopenGS_関数を呼ぶ_('台湾グッズ作品マスターを作成')) return;
  SpreadsheetApp.getUi().alert('台湾グッズ作品マスターを作成 関数が見つかりません');
}

function 台湾グッズ_Works更新() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('台湾グッズ');
  if (!sh || sh.getLastRow() < 2) {
    ui.alert('台湾グッズシートにデータがありません');
    return;
  }

  const ensureFn = OnopenGS_関数を取得_('台湾グッズ_Worksシートを確保_');
  const updateFn = OnopenGS_関数を取得_('台湾グッズ_Worksを更新_');
  if (!ensureFn || !updateFn) {
    ui.alert('台湾グッズ Works 更新に必要な関数が見つかりません');
    return;
  }

  const 列 = OnopenGS_ヘッダーMap_(sh);
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();

  const Works更新Map = {};
  data.forEach(row => {
    const 原題 = 列['原題タイトル'] ? String(row[列['原題タイトル'] - 1] || '').trim() : '';
    const 日本語 = 列['日本語タイトル'] ? String(row[列['日本語タイトル'] - 1] || '').trim() : '';
    if (!原題) return;
    if (!Works更新Map[原題]) Works更新Map[原題] = { 日本語, 件数: 0 };
    Works更新Map[原題].件数++;
    if (!Works更新Map[原題].日本語 && 日本語) Works更新Map[原題].日本語 = 日本語;
  });

  const Worksシート = ensureFn(ss);
  updateFn(Worksシート, Works更新Map);

  ui.alert(`✅ 台湾グッズ Works更新完了: ${Object.keys(Works更新Map).length}作品`);
}

function 台湾グッズ_重複チェック() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActive().getSheetByName('台湾グッズ');
  if (!sh || sh.getLastRow() < 2) {
    ui.alert('台湾グッズシートにデータがありません');
    return;
  }

  const 列 = OnopenGS_ヘッダーMap_(sh);
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();

  const maps = {
    サイト商品コード: {},
    原題商品タイトル: {},
    商品コードSKU: {}
  };

  data.forEach((row, i) => {
    const rowNum = i + 2;

    const サイト商品コード = 列['サイト商品コード'] ? String(row[列['サイト商品コード'] - 1] || '').trim() : '';
    const 原題商品タイトル = 列['原題商品タイトル'] ? String(row[列['原題商品タイトル'] - 1] || '').trim() : '';
    const 商品コードSKU = 列['商品コード（SKU）'] ? String(row[列['商品コード（SKU）'] - 1] || '').trim() : '';

    if (サイト商品コード) {
      if (!maps.サイト商品コード[サイト商品コード]) maps.サイト商品コード[サイト商品コード] = [];
      maps.サイト商品コード[サイト商品コード].push(rowNum);
    }
    if (原題商品タイトル) {
      if (!maps.原題商品タイトル[原題商品タイトル]) maps.原題商品タイトル[原題商品タイトル] = [];
      maps.原題商品タイトル[原題商品タイトル].push(rowNum);
    }
    if (商品コードSKU) {
      if (!maps.商品コードSKU[商品コードSKU]) maps.商品コードSKU[商品コードSKU] = [];
      maps.商品コードSKU[商品コードSKU].push(rowNum);
    }
  });

  const lines = [];

  Object.entries(maps.サイト商品コード).forEach(([key, rows]) => {
    if (rows.length > 1) lines.push(`サイト商品コード重複: ${key} → ${rows.join(', ')}`);
  });
  Object.entries(maps.原題商品タイトル).forEach(([key, rows]) => {
    if (rows.length > 1) lines.push(`原題商品タイトル重複: ${key} → ${rows.join(', ')}`);
  });
  Object.entries(maps.商品コードSKU).forEach(([key, rows]) => {
    if (rows.length > 1) lines.push(`商品コード（SKU）重複: ${key} → ${rows.join(', ')}`);
  });

  if (lines.length === 0) {
    ui.alert('✅ 重複は見つかりませんでした');
    return;
  }

  ui.alert(`⚠️ 重複が見つかりました\n\n${lines.join('\n')}`);
}

/* ============================================================
 * シート作成ラッパー
 * ============================================================ */

function 台湾書籍その他シートを作成() {
  if (OnopenGS_関数を呼ぶ_('台湾書籍系シートを作成する_共通', '台湾書籍その他')) return;
  SpreadsheetApp.getUi().alert('台湾書籍系シートを作成する_共通 関数が見つかりません');
}

function 台湾書籍Worksシートを作成() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();

  let cfg = null;
  if (typeof 設定_台湾まんが !== 'undefined') cfg = 設定_台湾まんが;
  if (!cfg && typeof 設定_台湾書籍その他 !== 'undefined') cfg = 設定_台湾書籍その他;

  if (!cfg) {
    ui.alert('書籍Works用の設定が見つかりません');
    return;
  }

  const シート名 = cfg.作品シート名;
  if (ss.getSheetByName(シート名)) {
    ui.alert(`「${シート名}」シートは既に存在します`);
    return;
  }

  const sh = ss.insertSheet(シート名);
  sh.getRange(1, 1, 1, cfg.作品列数).setValues([cfg.作品ヘッダー]);
  sh.getRange(1, 1, 1, cfg.作品列数)
    .setBackground('#cc0000')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sh.setFrozenRows(1);

  ui.alert(`✅ ${シート名} シートを作成しました`);
}

/* ============================================================
 * 台湾まんが メニューラッパー
 * ============================================================ */

function 台湾まんが_確定発行()         { _kyoutuu.確定発行を実行(設定_台湾まんが); }
function 台湾まんが_削除()             { _kyoutuu.削除を実行(設定_台湾まんが); }
function 台湾まんが_一括更新()         { _kyoutuu.一括更新を実行(設定_台湾まんが); }
function 台湾まんが_プルダウン更新()   { _kyoutuu.プルダウン更新を実行(設定_台湾まんが); }
function 台湾まんが_Works初期化()      { _kyoutuu書籍Works初期化_(設定_台湾まんが); }
function 台湾まんが_WorksKey再正規化() { _kyoutuu書籍WorksKey再正規化_(設定_台湾まんが); }
function 台湾まんが_重複統合()         { _kyoutuu書籍重複統合_(設定_台湾まんが); }
function 台湾まんが_ID振り直し()       { _kyoutuu書籍ID振り直し_(設定_台湾まんが); }
function 台湾まんが_孤立削除()         { _kyoutuu書籍孤立削除_(設定_台湾まんが); }

/* ============================================================
 * 台湾書籍その他 メニューラッパー
 * ============================================================ */

function 台湾書籍その他_確定発行()         { _kyoutuu.確定発行を実行(設定_台湾書籍その他); }
function 台湾書籍その他_削除()             { _kyoutuu.削除を実行(設定_台湾書籍その他); }
function 台湾書籍その他_一括更新()         { _kyoutuu.一括更新を実行(設定_台湾書籍その他); }
function 台湾書籍その他_プルダウン更新()   { _kyoutuu.プルダウン更新を実行(設定_台湾書籍その他); }
function 台湾書籍その他_Works初期化()      { _kyoutuu書籍Works初期化_(設定_台湾書籍その他); }
function 台湾書籍その他_WorksKey再正規化() { _kyoutuu書籍WorksKey再正規化_(設定_台湾書籍その他); }
function 台湾書籍その他_重複統合()         { _kyoutuu書籍重複統合_(設定_台湾書籍その他); }
function 台湾書籍その他_ID振り直し()       { _kyoutuu書籍ID振り直し_(設定_台湾書籍その他); }
function 台湾書籍その他_孤立削除()         { _kyoutuu書籍孤立削除_(設定_台湾書籍その他); }

/* ============================================================
 * 書籍Works系 共通ラッパー
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
      for (const item of dup.行配列) レポート += `  ID:${item.データ[1]} (行${item.行})\n`;
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

function チェック_全チェック()   { _チェック操作(true,  'all'); }
function チェック_全解除()       { _チェック操作(false, 'all'); }
function チェック_未登録のみ()   { _チェック操作(true,  'unregistered'); }

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

  ScriptApp.getProjectTriggers()
    .filter(t => ['onEdit', 'onEditInstallable_'].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('onEditInstallable_')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast('トリガー設定完了', '完了', 3);
}

/* ============================================================
 * 再計算補助
 * ============================================================ */

function 台湾まんが_再計算(開始行, 行数) {
  const sh = SpreadsheetApp.getActive().getSheetByName('台湾まんが');
  if (!sh) return;

  const cfg = 設定_台湾まんが;
  const 列マップ = _kyoutuu.列番号を取得(sh);

  const ダミーe = {
    range: sh.getRange(開始行, 1, 行数, sh.getLastColumn())
  };

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    _kyoutuu.onEdit処理を実行(ダミーe, sh, cfg, 列マップ, 開始行, 行数);
  } finally {
    lock.releaseLock();
  }
}