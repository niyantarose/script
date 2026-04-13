/**
 * Onopen.gs（韓国書籍・映像ファイル専用）
 * メニュー作成・シート設定マップ・onOpen・onEdit
 */

/* ============================================================
 * シート設定マップ
 * 韓国書籍・韓国音楽映像 → 共通Worksフレームワーク
 * 韓国雑誌 → 独自onEdit（マップには含めない）
 * ============================================================ */
function シート設定を取得(シート名) {
  const map = {
    '韓国書籍':     設定_韓国書籍,
    '韓国音楽映像': 設定_韓国音楽映像,
  };
  return map[シート名] || null;
}

/* ============================================================
 * onOpen
 * ============================================================ */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('📚 韓国書籍・映像管理');

  menu

  .addSubMenu(
  ui.createMenu('☑ チェック操作')
    .addItem('全チェック',         'チェック_全チェック')
    .addItem('全チェック解除',     'チェック_全解除')
    .addItem('未登録のみチェック', 'チェック_未登録のみ')
)
    .addSubMenu(
      ui.createMenu('韓国書籍')
        .addItem('① 確定発行',         '韓国書籍_確定発行')
        .addItem('② チェック行を削除', '韓国書籍_削除')
        .addItem('③ 一括更新',         '韓国書籍_一括更新')
        .addItem('⑥ プルダウン更新',   '韓国書籍_プルダウン更新')
    )
    .addSubMenu(
      ui.createMenu('韓国音楽映像')
        .addItem('① 確定発行',         '韓国音楽映像_確定発行')
        .addItem('② チェック行を削除', '韓国音楽映像_削除')
        .addItem('③ 一括更新',         '韓国音楽映像_一括更新')
        .addItem('⑥ プルダウン更新',   '韓国音楽映像_プルダウン更新')
    )
    .addSubMenu(
  ui.createMenu('韓国雑誌')
    .addItem('① 確定発行',                 '韓国雑誌_確定発行')
    .addItem('⑥ プルダウン更新',           '韓国雑誌_プルダウン更新')
    .addSeparator()
    .addItem('📋 候補の件数を確認',         '韓国雑誌_候補件数を確認')
    .addItem('✅ 候補を正式マスターへ反映', '韓国雑誌_候補を正式マスターへ反映')
)
    .addSeparator()
    .addSubMenu(
      ui.createMenu('シート作成・初期設定')
        .addItem('韓国書籍シートを作成',     '韓国書籍シートを作成')
        .addItem('韓国音楽映像シートを作成', '韓国音楽映像シートを作成')
        .addSeparator()
        .addItem('韓国雑誌マスターを作成',   '韓国雑誌マスターを作成')
        .addItem('韓国雑誌シートを作成',     '韓国雑誌シートを作成')
    )
    .addToUi();
}

/* ============================================================
 * onEdit トリガー
 * 韓国書籍・韓国音楽映像 → 共通フレームワーク（_kyoutuu.メインonEdit）
 * 韓国雑誌 → 独自処理（韓国雑誌_onEdit）
 * ============================================================ */
function onEdit(e) {
  const シート名 = e.range.getSheet().getName();
  if (シート名 === '韓国雑誌') {
    韓国雑誌_onEdit(e);
    return;
  }
  メインonEdit(e);  // _kyoutuu. を削除
}

function チェック_全チェック()   { _チェック操作(true, 'all'); }
function チェック_全解除()       { _チェック操作(false, 'all'); }
function チェック_未登録のみ()   { _チェック操作(true, 'unregistered'); }

function _チェック操作(チェック値, モード) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ヘッダー = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const 列 = {};
  ヘッダー.forEach((h, i) => { if (h) 列[String(h).trim()] = i + 1; });
  const チェック列 = 列['発番発行'];
  const 登録状況列 = 列['登録状況'];
  if (!チェック列 || !登録状況列) {
    SpreadsheetApp.getUi().alert('このシートはチェック操作に対応していません');
    return;
  }
  const 全行 = sh.getRange(2, 登録状況列, sh.getLastRow() - 1, 1).getValues();
  let 最終行 = 1;
  for (let i = 全行.length - 1; i >= 0; i--) {
    if (全行[i][0] !== '') { 最終行 = i + 2; break; }
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