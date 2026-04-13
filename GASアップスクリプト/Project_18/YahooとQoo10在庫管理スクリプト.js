// ────────────────────────────────────────────
// Apps Script：Yahoo & Qoo10 在庫差分自動化（テンポラリシート経由でIMPORTQUERY）
// ────────────────────────────────────────────

// ■ 定数
const SHEET_YAHOO = '即納抽出';        // Yahoo 在庫取得シート名
const RAW_SHEET   = 'Qoo10在庫_raw';   // Qoo10 データ貼付＆差分計算シート名
const SRC_SS_ID   = '18ZyInb5wjP6n9nEGV2Sl2Lp4ayhFyXuC'; // Qoo10 元スプレッドシート ID
const SRC_SHEET   = 'Qoo10在庫_raw';   // Qoo10 元シート名
const TMP_SHEET   = '__IMPORT_TMP__';  // 一時インポート用シート名

/**
 * importQoo10ViaTmp()
 * テンポラリシートを作成して IMPORTRANGE+QUERY を実行
 * 取り込まれた値だけ RAW_SHEET に貼り付け、テンポラリを削除します。
 */
function importQoo10ViaTmp() {
  const ss = SpreadsheetApp.getActive();
  // 1) 既存のテンポラリシートを削除
  const old = ss.getSheetByName(TMP_SHEET);
  if (old) ss.deleteSheet(old);
  // 2) テンポラリシートを新規作成
  const tmp = ss.insertSheet(TMP_SHEET);

  // 3) IMPORTRANGE+QUERY 式を A1 にセット
  const formula =
    '=QUERY(' +
      'IMPORTRANGE("' + SRC_SS_ID + '","' + SRC_SHEET + '!B:M"),' +
      '"select Col1,Col12 where Col1 is not null",1' +
    ')';
  tmp.getRange('A1').setFormula(formula);

  // 4) Google に計算させてから値を取得
  SpreadsheetApp.flush();
  Utilities.sleep(5000);  // データ量によっては増減してください

  const lastRow = tmp.getLastRow();
  if (lastRow < 1) {
    ss.deleteSheet(tmp);
    return;
  }
  const values = tmp.getRange(1, 1, lastRow, 2).getValues();

  // 5) RAW_SHEET に貼り付け
  const raw = ss.getSheetByName(RAW_SHEET);
  raw.clearContents();
  raw.getRange(1, 1, values.length, 2).setValues(values);

  // 6) テンポラリシートを削除
  ss.deleteSheet(tmp);
}

/**
 * diffQoo10()
 * RAW_SHEET の B 列 (在庫数) を読み、C 列に差分を書き込みます。
 */
function diffQoo10() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(RAW_SHEET);
  if (!sheet) return;
  const last = sheet.getLastRow();
  if (last < 2) return;

  const rows   = sheet.getRange(1, 1, last, 2).getValues();
  const cache  = CacheService.getScriptCache();
  const oldMap = JSON.parse(cache.get('qoo10_inv') || '{}');
  const diffs  = [['差分']];
  const newMap = {};

  for (let i = 1; i < rows.length; i++) {
    const code = String(rows[i][0]);
    const now  = Number(rows[i][1]) || 0;
    const prev = (oldMap[code] != null) ? Number(oldMap[code]) : now;
    diffs.push([ now - prev ]);
    newMap[code] = now;
  }
  sheet.getRange(1, 3, diffs.length, 1).setValues(diffs);
  cache.put('qoo10_inv', JSON.stringify(newMap), 6*60*60);
}

/**
 * diffYahoo()
 * SHEET_YAHOO の B 列（現在在庫）を読み、C 列に差分を書き込みます。
 */
function diffYahoo() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_YAHOO);
  if (!sheet) return;
  const last = sheet.getLastRow();
  if (last < 2) return;

  const rows   = sheet.getRange(1, 1, last, 2).getValues();
  const cache  = CacheService.getScriptCache();
  const oldMap = JSON.parse(cache.get('yahoo_inv') || '{}');
  const diffs  = [['差分']];
  const newMap = {};

  for (let i = 1; i < rows.length; i++) {
    const code = String(rows[i][0]);
    const now  = Number(rows[i][1]) || 0;
    const prev = (oldMap[code] != null) ? Number(oldMap[code]) : now;
    diffs.push([ now - prev ]);
    newMap[code] = now;
  }
  sheet.getRange(1, 3, diffs.length, 1).setValues(diffs);
  cache.put('yahoo_inv', JSON.stringify(newMap), 6*60*60);
}

/**
 * updateInventoryAndDiff()
 * Qoo10 インポート → Qoo10 差分 → Yahoo 差分 の順に実行します。
 */
function updateInventoryAndDiff() {
  importQoo10ViaTmp();
  diffQoo10();
  diffYahoo();
}

/**
 * onOpen()
 * スプレッドシート起動時に「在庫管理」メニューを追加します。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('在庫管理')
    .addItem('インポート＆差分実行', 'updateInventoryAndDiff')
    .addToUi();
}
