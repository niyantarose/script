/**
 * ①商品入力シート(C=親コード, D=種類数) から
 * ②オプション入力シート(A:D) を毎回作り直す（種類数=1も作成）
 * v2: getUi() 安全化
 *
 * 見やすさ:
 * - 親コードが変わったら空行を1行入れる
 * - さらに「お取り寄せ」行の下に太罫線
 */

// ====================================================================
// ★ 安全なUI通知（UIが無くても落ちない）
// ====================================================================

function OPT_uiSafeAlert_(msg) {
  try {
    SpreadsheetApp.getUi().alert(String(msg));
    return;
  } catch (e) {}

  // フォールバック: toast → log
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(String(msg), 'オプション生成', 10);
  } catch (e2) {}

  Logger.log('[OPT_UI_FALLBACK] ' + msg);
}


function 一から二オプション入力シートを自動生成_1も作る() {
  const ss = SpreadsheetApp.getActive();
  const sh1 = ss.getSheetByName('①商品入力シート');
  const sh2 = ss.getSheetByName('②オプション入力シート');

  if (!sh1) throw new Error('①商品入力シート が見つかりません');
  if (!sh2) throw new Error('②オプション入力シート が見つかりません');

  const HEADER1 = 2; // ①はヘッダー2行
  const lastRow = sh1.getLastRow();

  // ②を作成前にクリア（A:D のデータ部）
  const last2 = sh2.getLastRow();
  if (last2 >= 2) sh2.getRange(2, 1, last2 - 1, 4).clearContent();

  if (lastRow <= HEADER1) {
    OPT_uiSafeAlert_('①にデータがありません。②をクリアしました。');
    return;
  }

  // ①から C:D を読む（C=3列目, D=4列目）
  const numRows = lastRow - HEADER1;
  const vals = sh1.getRange(HEADER1 + 1, 3, numRows, 2).getValues(); // C:D

  const out = [];
  let prevCode = null;

  for (let i = 0; i < vals.length; i++) {
    const code = String(vals[i][0] || '').trim();   // C: 親コード
    if (!code) continue;

    let kinds = Number(String(vals[i][1] || '').trim()); // D: 種類数
    if (!isFinite(kinds) || kinds <= 0) kinds = 1;

    // 親コードが変わったら空行（見やすさ）
    if (prevCode !== null && code !== prevCode) {
      out.push(['', '', '', '']); // 完全空行
    }
    prevCode = code;

    for (let k = 1; k <= kinds; k++) {
      const subA = `${code}-${k}a`;
      const subB = `${code}-${k}b`;

      out.push([code, '即納（日本在庫）', subA, '']);              // Dは手入力
      out.push([code, 'お取り寄せ（韓国から）', subB, '']);        // ← 既存資産と整合させる
    }
  }

  if (out.length > 0) {
    sh2.getRange(2, 1, out.length, 4).setValues(out);
  }

  // お取り寄せ行の下に太罫線
  お取り寄せ行の下に罫線を引く_オプション入力();

  OPT_uiSafeAlert_(`②オプション入力シートを作成しました。\n出力行数: ${out.length}`);
}

/**
 * ②オプション入力シート
 * 「お取り寄せ」行の下に太罫線を引く
 */
function お取り寄せ行の下に罫線を引く_オプション入力() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('②オプション入力シート');
  if (!sh) return;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const lastCol = sh.getLastColumn();

  // 既存の下罫線を一旦クリア（データ部だけ）
  sh.getRange(2, 1, lastRow - 1, lastCol).setBorder(null, null, false, null, null, null);

  const values = sh.getRange(2, 1, lastRow - 1, 2).getValues(); // A:B

  values.forEach((row, i) => {
    const stockType = String(row[1] || '');
    if (stockType.includes('お取り寄せ')) {
      const targetRow = i + 2;
      sh.getRange(targetRow, 1, 1, lastCol)
        .setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  });
}