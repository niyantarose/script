/*************************************************
 * 日本語タイトル辞書：プルダウン設定 フル版
 *
 * できること
 * 1. 「日本語タイトル辞書」シートの
 *    A列「言語」 → 「言語マスター」A列からプルダウン
 *    B列「カテゴリ」 → 「カテゴリマスター」A列からプルダウン
 *
 * 2. スクリプト画面から
 *    「日本語タイトル辞書_プルダウン設定」を直接実行できる
 *
 * 3. シートを開いたときにメニューからも実行できる
 *************************************************/


/** ===== 設定 ===== */
const TITLE_DICT_CONFIG = {
  辞書シート名: '日本語タイトル辞書',
  言語マスター名: '言語マスター',
  カテゴリマスター名: 'カテゴリマスター',

  // 辞書シートの列
  言語列: 1,      // A列
  カテゴリ列: 2,  // B列

  // 何行目から入力欄か
  開始行: 2
};


/**
 * スクリプト画面から直接実行する本体
 * これを実行すればOK
 */
function 日本語タイトル辞書_プルダウン設定() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  日本語タイトル辞書_プルダウン設定_実行_(ss);
}


/**
 * メニュー追加
 */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('日本語タイトル辞書')
      .addItem('プルダウン設定', '日本語タイトル辞書_プルダウン設定')
      .addToUi();
  } catch (err) {
    // UIがない状況では無視
    Logger.log(err);
  }
}


/**
 * 実処理
 */
function 日本語タイトル辞書_プルダウン設定_実行_(ss) {
  const cfg = TITLE_DICT_CONFIG;

  const 辞書シート = ss.getSheetByName(cfg.辞書シート名);
  const 言語マスター = ss.getSheetByName(cfg.言語マスター名);
  const カテゴリマスター = ss.getSheetByName(cfg.カテゴリマスター名);

  if (!辞書シート) {
    throw new Error(`シート「${cfg.辞書シート名}」が見つかりません`);
  }
  if (!言語マスター) {
    throw new Error(`シート「${cfg.言語マスター名}」が見つかりません`);
  }
  if (!カテゴリマスター) {
    throw new Error(`シート「${cfg.カテゴリマスター名}」が見つかりません`);
  }

  // A1は見出し想定
  const 言語最終行 = 言語マスター.getLastRow();
  const カテゴリ最終行 = カテゴリマスター.getLastRow();

  if (言語最終行 < 2) {
    throw new Error(`「${cfg.言語マスター名}」に候補データがありません`);
  }
  if (カテゴリ最終行 < 2) {
    throw new Error(`「${cfg.カテゴリマスター名}」に候補データがありません`);
  }

  // プルダウン元範囲
  const 言語候補範囲 = 言語マスター.getRange(2, 1, 言語最終行 - 1, 1);       // A2:A
  const カテゴリ候補範囲 = カテゴリマスター.getRange(2, 1, カテゴリ最終行 - 1, 1); // A2:A

  // 辞書シートの適用先
  const 最大行 = 辞書シート.getMaxRows();
  const 行数 = 最大行 - cfg.開始行 + 1;

  if (行数 <= 0) {
    throw new Error(`「${cfg.辞書シート名}」の行数が不足しています`);
  }

  const 言語入力範囲 = 辞書シート.getRange(cfg.開始行, cfg.言語列, 行数, 1);
  const カテゴリ入力範囲 = 辞書シート.getRange(cfg.開始行, cfg.カテゴリ列, 行数, 1);

  // 入力規則
  const 言語ルール = SpreadsheetApp.newDataValidation()
    .requireValueInRange(言語候補範囲, true)
    .setAllowInvalid(false)
    .setHelpText('言語マスターから選択してください')
    .build();

  const カテゴリルール = SpreadsheetApp.newDataValidation()
    .requireValueInRange(カテゴリ候補範囲, true)
    .setAllowInvalid(false)
    .setHelpText('カテゴリマスターから選択してください')
    .build();

  // 設定
  言語入力範囲.setDataValidation(言語ルール);
  カテゴリ入力範囲.setDataValidation(カテゴリルール);

  // 任意：1行目見出しを太字にする
  辞書シート.getRange(1, 1, 1, 7).setFontWeight('bold');

  // 通知
  安全にToast_(
    ss,
    `「${cfg.辞書シート名}」にプルダウンを設定しました`,
    '完了'
  );

  Logger.log(`完了: ${cfg.辞書シート名} にプルダウン設定`);
}


/**
 * Toastはスクリプト画面実行時に環境によって出ないことがあるので安全化
 */
function 安全にToast_(ss, message, title) {
  try {
    ss.toast(message, title || '通知', 5);
  } catch (err) {
    Logger.log(`${title || '通知'}: ${message}`);
  }
}