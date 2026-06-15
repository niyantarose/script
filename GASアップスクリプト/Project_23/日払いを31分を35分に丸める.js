/**
 * 実働時間を「5分単位・切り上げ」に変更するスクリプト
 *
 * 各月シートの F列（実働時間）の数式を、
 *   (退勤 - 出勤) - 休憩(分) → 分に換算 → 5分単位で切り上げ
 * に書き換える。給与・残業はこの丸め後の実働時間で計算される。
 *
 * 例）8:11 → 8:15 ／ 8:00 → 8:00 ／ 8:16 → 8:20
 *
 * 【実行方法】Apps Script に貼り付け → 関数「applyFiveMinRoundUp」を実行。
 */
function applyFiveMinRoundUp() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var YEAR = 2026;

  var monthSheets = ss.getSheets().filter(function(s) {
    return /^\d+月$/.test(s.getName());
  });

  monthSheets.forEach(function(ws) {
    var m = parseInt(ws.getName());
    var ndays = new Date(YEAR, m, 0).getDate();

    var formulas = [];
    for (var d = 1; d <= ndays; d++) {
      var r = 4 + d;
      // (D-C)*1440 = 拘束分, -E(休憩分) = 実働分 → CEILINGで5分切り上げ → /1440で時刻シリアルへ
      formulas.push([
        '=IF(OR(C' + r + '="",D' + r + '=""),"",' +
        'CEILING((D' + r + '-C' + r + ')*1440-E' + r + ',5)/1440)'
      ]);
    }
    // F列（6列目）の5行目から ndays 行分を一括更新
    ws.getRange(5, 6, ndays, 1).setFormulas(formulas);

    Logger.log(ws.getName() + ' → F列を5分切り上げに更新（' + ndays + '行）');
  });

  SpreadsheetApp.getUi().alert('完了！実働時間を「5分単位・切り上げ」に更新しました。');
}