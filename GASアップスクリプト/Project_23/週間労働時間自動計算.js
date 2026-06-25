/**
 * 労務管理_2026 - 週間労働時間（雇用保険チェック）
 *
 * 各月シート（1月〜12月）の一番右に「週間労働時間」列を追加し、
 * その行が属する週（月曜〜日曜）の実働時間合計を「時間（小数1桁）」で表示する。
 *   ・週20時間以上 … 赤（雇用保険の加入ライン）
 *   ・週18時間以上 … 黄（要注意）
 *
 * 列番号は固定せず、見出し行（4行目）の文字から「日付」列・「実働」列を自動検出する。
 * これにより、休憩列の有無などシートごとのレイアウト差があっても正しく動作する。
 *
 * 実行: メニュー「週間労働時間」→「週合計を計算（雇用保険チェック）」
 *       もしくは関数 setupWeeklyWorkHours を直接実行。
 *
 * 注意: 月をまたぐ週（例 6/29月〜7/5日）は各月シートに分かれて集計される。
 */
function setupWeeklyWorkHours() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const HEADER_ROW = 4;
  const FIRST_DATA_ROW = 5;
  const YEAR = 2026;
  const WEEK_LIMIT = 20; // 雇用保険ライン（時間/週）
  const WARN_LIMIT = 18; // 早めの警告ライン
  const HEADER_TEXT = '週間労働時間';

  const sheets = ss.getSheets().filter(function (s) {
    return /^\d{1,2}月$/.test(s.getName());
  });

  const done = [];

  sheets.forEach(function (sh) {
    const m = parseInt(sh.getName(), 10);
    const ndays = new Date(YEAR, m, 0).getDate(); // その月の日数
    const lastDataRow = HEADER_ROW + ndays; // 5 .. (4 + ndays)

    const headers = sh.getRange(HEADER_ROW, 1, 1, sh.getMaxColumns()).getValues()[0];

    let dateCol = 0;
    let workCol = 0;
    let existCol = 0;
    let lastHeaderCol = 0;
    for (let c = 0; c < headers.length; c++) {
      const h = String(headers[c]).trim();
      if (h !== '') lastHeaderCol = c + 1;
      if (h.indexOf('日付') >= 0) dateCol = c + 1;
      if (h.indexOf('実働') >= 0) workCol = c + 1;
      if (h.indexOf(HEADER_TEXT) >= 0) existCol = c + 1;
    }
    if (dateCol === 0 || workCol === 0) return; // 見出しが見つからないシートはスキップ

    const outCol = existCol || lastHeaderCol + 1;
    const dL = whColLetter_(dateCol);
    const wL = whColLetter_(workCol);

    // 見出し（右隣に新設するときは、左の見出しの書式をコピーして体裁を合わせる）
    const headerCell = sh.getRange(HEADER_ROW, outCol);
    if (lastHeaderCol >= 1 && outCol > lastHeaderCol) {
      sh.getRange(HEADER_ROW, lastHeaderCol).copyTo(headerCell, { formatOnly: true });
    }
    headerCell.setValue(HEADER_TEXT).setFontWeight('bold').setHorizontalAlignment('center');

    // 数式: その行の週（月〜日）の実働合計を時間(小数)で。空欄・非日付行は空欄。
    const rng = '$' + dL + '$' + FIRST_DATA_ROW + ':$' + dL + '$' + lastDataRow;
    const wrng = '$' + wL + '$' + FIRST_DATA_ROW + ':$' + wL + '$' + lastDataRow;
    const formulas = [];
    for (let r = FIRST_DATA_ROW; r <= lastDataRow; r++) {
      const mon = '($' + dL + r + '-WEEKDAY($' + dL + r + ',3))'; // その週の月曜
      formulas.push([
        '=IF(ISNUMBER($' + dL + r + '),' +
          'ROUND(SUMIFS(' + wrng + ',' +
          rng + ',">="&' + mon + ',' +
          rng + ',"<="&' + mon + '+6)*24,1),"")',
      ]);
    }
    const dataRange = sh.getRange(FIRST_DATA_ROW, outCol, ndays, 1);
    dataRange.setFormulas(formulas).setNumberFormat('0.0"h"').setHorizontalAlignment('center');

    // 条件付き書式: このM列のみ。既存の同列ルールは置き換え、他シート/他列のルールは保持。
    const over = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(WEEK_LIMIT)
      .setBackground('#F4C7C3')
      .setFontColor('#CC0000')
      .setBold(true)
      .setRanges([dataRange])
      .build();
    const warn = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(WARN_LIMIT, WEEK_LIMIT)
      .setBackground('#FCE8B2')
      .setRanges([dataRange])
      .build();
    const kept = sh.getConditionalFormatRules().filter(function (rule) {
      return !rule.getRanges().some(function (g) {
        return g.getColumn() === outCol && g.getNumColumns() === 1;
      });
    });
    sh.setConditionalFormatRules(kept.concat([over, warn]));

    done.push(sh.getName());
  });

  SpreadsheetApp.getUi().alert(
    '週間労働時間の列を更新しました（' + done.length + 'シート）。\n\n' +
      '・各行 = その週(月〜日)の実働合計（時間）\n' +
      '・' + WEEK_LIMIT + 'h以上 … 赤（雇用保険ライン）\n' +
      '・' + WARN_LIMIT + 'h以上 … 黄（要注意）\n\n' +
      '※月をまたぐ週は各月シートに分かれて集計されます。'
  );
}

/** 列番号(1始まり)を列記号(A, B, ... , AA)に変換 */
function whColLetter_(col) {
  let s = '';
  while (col > 0) {
    const mod = (col - 1) % 26;
    s = String.fromCharCode(65 + mod) + s;
    col = Math.floor((col - mod - 1) / 26);
  }
  return s;
}
