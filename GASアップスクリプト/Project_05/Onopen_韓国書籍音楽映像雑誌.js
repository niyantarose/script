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

  ui.createMenu('📚 韓国書籍・映像管理')

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
        .addSeparator()
        .addItem('🔍 列名チェック', '韓国書籍_列名チェック')
    )

    .addSubMenu(
      ui.createMenu('韓国音楽映像')
        .addItem('① 確定発行',         '韓国音楽映像_確定発行')
        .addItem('② チェック行を削除', '韓国音楽映像_削除')
        .addItem('③ 一括更新',         '韓国音楽映像_一括更新')
        .addItem('⑥ プルダウン更新',   '韓国音楽映像_プルダウン更新')
        .addSeparator()
        .addItem('🎬 クレジット種別列を追加',       '韓国音楽映像_クレジット種別列を追加')
        .addItem('🎬 クレジット種別プルダウン作成', '韓国音楽映像_クレジット種別プルダウン作成')
        .addSeparator()
        .addItem('🔍 列名チェック',       '韓国音楽映像_列名チェック')
        .addItem('🔁 選択行を再生成',     '韓国音楽映像_選択行を再生成')
        .addItem('🧹 自己更新フラグ解除', '韓国音楽映像_自己更新フラグ解除')
    )

    .addSubMenu(
      ui.createMenu('韓国雑誌')
        .addItem('🚚 ダニエル商品コード取得',   '韓国雑誌_ダニエル商品コード取得')
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
function onEdit_インストール型(e) {
  console.log('=== onEdit START ===');

  if (!e || !e.range) {
    console.log('e または e.range がありません');
    return;
  }

  const sh = e.range.getSheet();
  const シート名 = String(sh.getName() || '').trim();
  const 開始行 = e.range.getRow();
  const 行数 = e.range.getNumRows();

  console.log('シート名: ' + シート名);
  console.log('編集セル: row=' + 開始行 + ', col=' + e.range.getColumn());

  if (シート名 === '韓国雑誌') {
    韓国雑誌_onEdit(e);
    return;
  }

  const cfg = シート設定を取得(シート名);
  console.log('cfg取得: ' + (cfg ? 'OK' : 'NG'));

  if (!cfg) return;
  if (開始行 + 行数 - 1 < 2) return;

  const lock = LockService.getDocumentLock();

  try {
    if (!lock.tryLock(5000)) {
      console.log('ロック取得できず終了');
      return;
    }

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

    if (!対象列が含まれる) {
      console.log('監視対象外なので終了');
      return;
    }

    console.log('onEdit処理を実行 前');

    _kyoutuu.onEdit処理を実行(
      e,
      sh,
      cfg,
      列マップ,
      開始行,
      行数
    );

    console.log('onEdit処理を実行 後');

  } catch (err) {
    console.error('onEdit生成エラー: ' + err.message);
    throw err;
  } finally {
    try {
      lock.releaseLock();
    } catch (_) {}
  }
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

if (sh.getLastRow() < 2) return;

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

function 共通_自己更新フラグ解除() {
  PropertiesService.getScriptProperties().deleteProperty('KYOUTUU_UPDATING');
  SpreadsheetApp.getUi().alert('✅ 自己更新フラグを解除しました');
}

function 韓国書籍_列名チェック() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(設定_韓国書籍.マスターシート名);

  if (!sh) {
    SpreadsheetApp.getUi().alert('韓国書籍シートが見つかりません');
    return;
  }

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
    .getValues()[0]
    .map(v => String(v || '').trim());

  const 必要列 = [
    ...(設定_韓国書籍.監視列 || []),
    ...Object.values(設定_韓国書籍.列名 || {}),
    ...(設定_韓国書籍.商品重複キー列優先順位 || [])
  ].filter(Boolean);

  const 不足列 = [...new Set(必要列)]
    .filter(name => !headers.includes(name));

  if (不足列.length) {
    SpreadsheetApp.getUi().alert(
      '❌ 不足している列があります\n\n' + 不足列.join('\n')
    );
  } else {
    SpreadsheetApp.getUi().alert('✅ 韓国書籍：列名は設定と一致しています');
  }
}