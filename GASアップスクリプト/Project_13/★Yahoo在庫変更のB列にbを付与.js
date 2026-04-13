/******************************************************
 * Yahoo在庫変更：末尾に "a" / "b" 付与（全体 + 選択範囲）
 ******************************************************/

/** ====== シート全体：B列 ====== */
function Yahoo在庫変更_B列_末尾にb付与() {
  Yahoo在庫変更_B列_末尾に指定文字付与_('b');
}
function Yahoo在庫変更_B列_末尾にa付与() {
  Yahoo在庫変更_B列_末尾に指定文字付与_('a');
}

/** ====== 選択範囲：選択セル全部 ====== */
function 選択範囲_末尾にb付与() {
  選択範囲_末尾に指定文字付与_('b');
}
function 選択範囲_末尾にa付与() {
  選択範囲_末尾に指定文字付与_('a');
}

/**
 * 内部共通：Yahoo在庫変更シートのB列だけ、ヘッダー除外で末尾に a/b 付与
 * @param {"a"|"b"} suffix
 */
function Yahoo在庫変更_B列_末尾に指定文字付与_(suffix) {
  const SHEET_NAME = 'Yahoo在庫変更';
  const HEADER_ROWS = 1;
  const TARGET_COL = 2; // B列
  const SUFFIX = String(suffix || '').toLowerCase();

  if (SUFFIX !== 'a' && SUFFIX !== 'b') {
    throw new Error('suffix は "a" または "b" を指定してください。');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`シート「${SHEET_NAME}」が見つかりません。`);

  const lastRow = sh.getLastRow();
  if (lastRow <= HEADER_ROWS) {
    SpreadsheetApp.getUi().alert('処理対象の行がありません。');
    return;
  }

  const startRow = HEADER_ROWS + 1;
  const numRows = lastRow - HEADER_ROWS;

  const range = sh.getRange(startRow, TARGET_COL, numRows, 1);
  const values = range.getValues();
  const formulas = range.getFormulas();

  const result = 末尾付与ロジック_(values, formulas, SUFFIX);

  range.setValues(result.values);

  const msg = `完了：B列 末尾「${SUFFIX}」付与 ${result.changed}件（空欄:${result.skippedEmpty} / a,b済:${result.skippedHasSuffix} / 数式:${result.skippedFormula} / 安全:${result.skippedUnsafe}）`;
  SpreadsheetApp.getActive().toast(msg, 'Yahoo在庫変更', 8);
  Logger.log(msg);
}

/**
 * 内部共通：選択範囲（アクティブレンジ）のセルに末尾 a/b 付与
 * - 選択がない/単一セルでもOK
 * - 数式セルはスキップ（式破壊防止）
 * @param {"a"|"b"} suffix
 */
function 選択範囲_末尾に指定文字付与_(suffix) {
  const SUFFIX = String(suffix || '').toLowerCase();
  if (SUFFIX !== 'a' && SUFFIX !== 'b') {
    throw new Error('suffix は "a" または "b" を指定してください。');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  const range = sh.getActiveRange();

  if (!range) {
    SpreadsheetApp.getUi().alert('選択範囲が取得できません。セル範囲を選択してから実行してください。');
    return;
  }

  const values = range.getValues();
  const formulas = range.getFormulas();

  const result = 末尾付与ロジック_(values, formulas, SUFFIX);

  range.setValues(result.values);

  const msg = `完了：選択範囲 末尾「${SUFFIX}」付与 ${result.changed}件（空欄:${result.skippedEmpty} / a,b済:${result.skippedHasSuffix} / 数式:${result.skippedFormula} / 安全:${result.skippedUnsafe}）`;
  SpreadsheetApp.getActive().toast(msg, '選択範囲', 8);
  Logger.log(msg);
}

/**
 * 共通ロジック：2次元配列 values を走査し、末尾に suffix を付ける
 * - formulas があるセルはスキップ
 * - 末尾が a/b はスキップ（事故防止）
 * - 末尾が英数字で終わらないものはスキップ（事故防止）
 *
 * @param {any[][]} values
 * @param {string[][]} formulas
 * @param {"a"|"b"} suffix
 * @return {{values:any[][], changed:number, skippedEmpty:number, skippedHasSuffix:number, skippedFormula:number, skippedUnsafe:number}}
 */
function 末尾付与ロジック_(values, formulas, suffix) {
  let changed = 0;
  let skippedEmpty = 0;
  let skippedHasSuffix = 0;
  let skippedFormula = 0;
  let skippedUnsafe = 0;

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (formulas && formulas[r] && formulas[r][c]) {
        skippedFormula++;
        continue;
      }

      const v = values[r][c];
      if (v === '' || v === null || typeof v === 'undefined') {
        skippedEmpty++;
        continue;
      }

      const s = String(v).trim();
      if (!s) {
        skippedEmpty++;
        continue;
      }

      // すでに末尾a/bなら触らない
      if (/[ab]$/i.test(s)) {
        skippedHasSuffix++;
        continue;
      }

      // 末尾が英数字で終わってないものは触らない
      if (!/[0-9A-Za-z]$/.test(s)) {
        skippedUnsafe++;
        continue;
      }

      values[r][c] = s + suffix;
      changed++;
    }
  }

  return { values, changed, skippedEmpty, skippedHasSuffix, skippedFormula, skippedUnsafe };
}
