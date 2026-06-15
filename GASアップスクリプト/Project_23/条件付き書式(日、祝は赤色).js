/**
 * タイミー労務管理 - 条件付き書式 修正スクリプト（他シート参照なし版）
 *
 * Google Sheets の条件付き書式は他シート参照を一切受け付けない（名前付き範囲経由でも
 * API では弾かれる）ため、各月の祝日を数式に DATE() で直接埋め込む方式にした。
 *
 * 【実行方法】
 * ① 「拡張機能 > Apps Script」のコードを全部消してこれを貼り付け
 * ② 関数セレクタを「fixConditionalFormatting」にして「▶ 実行」
 * ③ 権限を許可 → 完了アラートが出ればOK
 */

function fixConditionalFormatting() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var YEAR = 2026;

  // 月ごとの祝日（日にちだけ）。振替休日・国民の休日込み
  var holidaysByMonth = {
    1:  [1, 12],        // 元日・成人の日
    2:  [11, 23],       // 建国記念の日・天皇誕生日
    3:  [20],           // 春分の日
    4:  [29],           // 昭和の日
    5:  [3, 4, 5, 6],   // 憲法記念日・みどりの日・こどもの日・振替休日
    6:  [],
    7:  [20],           // 海の日
    8:  [11],           // 山の日
    9:  [21, 22, 23],   // 敬老の日・国民の休日・秋分の日
    10: [12],           // スポーツの日
    11: [3, 23],        // 文化の日・勤労感謝の日
    12: []
  };

  var monthSheets = ss.getSheets().filter(function(s) {
    return /^\d+月$/.test(s.getName());
  });

  monthSheets.forEach(function(ws) {
    var m = parseInt(ws.getName());
    var ndays = new Date(YEAR, m, 0).getDate(); // その月の日数
    var firstRow = 5;
    var range = ws.getRange(firstRow, 1, ndays, 11); // A5:K(4+ndays)

    // 赤ルールの条件：日曜 or その月の各祝日（DATE直書き）
    var conds = ['WEEKDAY($A' + firstRow + ',1)=1'];
    holidaysByMonth[m].forEach(function(day) {
      conds.push('$A' + firstRow + '=DATE(' + YEAR + ',' + m + ',' + day + ')');
    });
    var redFormula = '=OR(' + conds.join(',') + ')';

    // 日祝ルール（赤）
    var sunHolRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(redFormula)
      .setBackground('#FCE4E4')
      .setFontColor('#CC0000')
      .setRanges([range])
      .build();

    // 土曜ルール（青）
    var satRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=WEEKDAY($A' + firstRow + ',1)=7')
      .setBackground('#DDEBF7')
      .setFontColor('#0000CC')
      .setRanges([range])
      .build();

    // 既存ルールを削除して再設定（赤を先 = 優先）
    ws.clearConditionalFormatRules();
    ws.setConditionalFormatRules([sunHolRule, satRule]);

    Logger.log(ws.getName() + ' → ' + ndays + '日分を更新 / ' + redFormula);
  });

  SpreadsheetApp.getUi().alert('完了！全月シートの条件付き書式を更新しました（他シート参照なし）。');
}