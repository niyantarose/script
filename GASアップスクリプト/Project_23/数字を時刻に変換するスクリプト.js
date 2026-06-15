/**
 * 出退勤 時刻入力まわり（修正版）
 *  - 952 / 9.52 / 9:52 / 9：52 → すべて 9:52 に変換
 *  - 既に「9.52」等の"数値"で入ってしまったセルも修復できる
 *
 * 【使い方】
 *  ① このコードを貼り付けて保存
 *  ② まず「fixTimeNumbers」を1回実行 → 既存の壊れたセルを修復
 *  ③ 以降は onEdit が自動で動くので、出退勤は数字打ちでOK
 */

// ───────── 自動変換（入力時） ─────────
function onEdit(e) {
  if (!e || !e.range || e.value === undefined) return;
  var sheet = e.range.getSheet();
  if (!/^\d+月$/.test(sheet.getName())) return;
  var col = e.range.getColumn();
  if (col !== 3 && col !== 4) return;       // C=出勤, D=退勤
  if (e.range.getRow() < 5) return;

  var t = parseTimeInput(String(e.value).trim());
  if (t === null) return;
  e.range.setValue((t.h * 60 + t.m) / 1440);
  e.range.setNumberFormat('h:mm');
}

// 入力文字列を {h, m} に。変換不可なら null
function parseTimeInput(v) {
  var h, m;
  var sep = v.match(/^(\d{1,2})[.:：](\d{1,2})$/);  // 9.52 / 9:52 / 9：52
  if (sep) {
    h = parseInt(sep[1], 10);
    m = parseInt(sep[2], 10);
  } else if (/^\d{1,4}$/.test(v)) {                 // 952 / 1730 / 9
    var n = parseInt(v, 10);
    if (n < 100) { h = n; m = 0; }
    else { h = Math.floor(n / 100); m = n % 100; }
  } else {
    return null;
  }
  if (h > 23 || m > 59) return null;
  return { h: h, m: m };
}

// ───────── 既存セルの一括修復（1回だけ実行） ─────────
function fixTimeNumbers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var YEAR = 2026, fixed = 0;

  ss.getSheets().forEach(function(ws) {
    if (!/^\d+月$/.test(ws.getName())) return;
    var m = parseInt(ws.getName());
    var ndays = new Date(YEAR, m, 0).getDate();

    [3, 4].forEach(function(col) {                  // C列・D列
      var vals = ws.getRange(5, col, ndays, 1).getValues();
      for (var i = 0; i < vals.length; i++) {
        var v = vals[i][0];
        // 数値で、かつ 1以上（=時刻シリアル0〜1ではない）＝誤入力とみなす
        if (typeof v !== 'number' || v < 1) continue;

        var h, mm;
        if (Number.isInteger(v)) {                  // 952 / 1732 形式
          if (v < 100) { h = v; mm = 0; }
          else { h = Math.floor(v / 100); mm = v % 100; }
        } else {                                    // 9.52 / 17.32 形式
          h = Math.floor(v);
          mm = Math.round((v - h) * 100);
        }
        if (h > 23 || mm > 59) continue;            // 判定不能は触らない

        var cell = ws.getRange(4 + i + 1, col);
        cell.setValue((h * 60 + mm) / 1440);
        cell.setNumberFormat('h:mm');
        fixed++;
      }
    });
  });

  SpreadsheetApp.getUi().alert('完了！数値で入っていた出退勤 ' + fixed + ' 件を時刻に変換しました。');
}