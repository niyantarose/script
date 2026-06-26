/**
 * 労務管理_2026 - 休憩時間 自動計算セットアップ
 *
 * セットアップ手順（初回のみ）:
 * 1. スプレッドシートを再読み込み
 * 2. メニュー「休憩時間」→「自動計算をセットアップ」を実行
 */

const MONTH_SHEETS = ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月'];
const DATA_START_ROW = 5;
const DATA_END_ROW = 34;

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('休憩時間')
    .addItem('自動計算をセットアップ', 'setupBreakTimeAutoCalc')
    .addToUi();
  ui.createMenu('週間労働時間')
    .addItem('週合計を計算（雇用保険チェック）', 'setupWeeklyWorkHours')
    .addToUi();
}

function setupBreakTimeAutoCalc() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let count = 0;

  MONTH_SHEETS.forEach(function (sheetName) {
    if (setupSheet_(ss.getSheetByName(sheetName))) {
      count++;
    }
  });

  SpreadsheetApp.getUi().alert(
    '設定完了（' + count + 'シート）\n\n' +
      '【使い方】\n' +
      'E列: 休憩開始（例 13:00）\n' +
      'F列: 休憩終了（例 14:00）\n' +
      'G列: 休憩(分) が自動表示（例 60）\n\n' +
      '実働時間・給与の数式は列挿入により自動調整されます。'
  );
}

function setupSheet_(sheet) {
  if (!sheet) return false;
  if (sheet.getRange('E4').getValue() === '休憩開始') return false;

  const numRows = DATA_END_ROW - DATA_START_ROW + 1;
  const oldBreakMins = sheet.getRange(DATA_START_ROW, 5, numRows, 1).getValues();

  sheet.insertColumnsBefore(5, 2);

  sheet.getRange('E4').setValue('休憩開始');
  sheet.getRange('F4').setValue('休憩終了');

  sheet.getRange(DATA_START_ROW, 5, numRows, 2).setNumberFormat('h:mm');

  for (let i = 0; i < oldBreakMins.length; i++) {
    const mins = parseBreakMinutes_(oldBreakMins[i][0]);
    if (mins > 0) {
      const row = DATA_START_ROW + i;
      sheet.getRange(row, 5).setValue('13:00');
      sheet.getRange(row, 6).setValue('14:00');
    }
  }

  const formula =
    '=IF(AND(E' + DATA_START_ROW + '<>"",F' + DATA_START_ROW + '<>""),' +
    'ROUND((F' + DATA_START_ROW + '-E' + DATA_START_ROW + ')*1440,0),"")';

  const formulaCell = sheet.getRange('G' + DATA_START_ROW);
  formulaCell.setFormula(formula);
  formulaCell.autoFill(sheet.getRange('G' + DATA_START_ROW + ':G' + DATA_END_ROW));

  return true;
}

function parseBreakMinutes_(value) {
  if (typeof value === 'number' && !isNaN(value)) return value;
  if (typeof value === 'string' && value.trim() !== '') {
    const n = Number(value);
    return isNaN(n) ? 0 : n;
  }
  return 0;
}
