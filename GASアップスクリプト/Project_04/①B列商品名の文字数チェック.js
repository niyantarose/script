/******************************************************
 * 商品名バイト数チェック
 * B列の商品名が75文字(150バイト)以内かチェックする
 * ※Yahoo Shopping基準: 全角=2バイト, 半角=1バイト
 ******************************************************/

/**
 * Shift-JIS基準のバイト数を計算
 * 全角文字=2バイト、半角文字=1バイト
 */
function countSjisBytes_(str) {
  let bytes = 0;
  for (let i = 0; i < str.length; i++) {
    const code = str.charCodeAt(i);
    // 半角: ASCII(0x00-0x7F) + 半角カナ(0xFF61-0xFF9F)
    if ((code >= 0x00 && code <= 0x7F) || (code >= 0xFF61 && code <= 0xFF9F)) {
      bytes += 1;
    } else {
      bytes += 2;
    }
  }
  return bytes;
}

/**
 * 共通チェック処理
 */
function 商品名バイト数チェック_汎用_(sheetName, headerRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('シートが見つかりません: ' + sheetName);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRows) {
    SpreadsheetApp.getUi().alert('データ行がありません: ' + sheetName);
    return;
  }

  const data = sheet.getRange(headerRows + 1, 2, lastRow - headerRows, 1).getValues();
  const over = [];

  for (let i = 0; i < data.length; i++) {
    const name = String(data[i][0] || '').trim();
    if (!name) continue;

    const chars = name.length;
    const bytes = countSjisBytes_(name);

    if (chars > 75 || bytes > 150) {
      over.push({
        行: i + headerRows + 1,
        文字数: chars,
        バイト数: bytes,
        商品名: name.length > 40 ? name.substring(0, 40) + '...' : name
      });
    }
  }

  if (over.length === 0) {
    SpreadsheetApp.getUi().alert(`✅ ${sheetName}\n\n全商品名が75文字(150バイト)以内です！`);
    return;
  }

  // 超過行をハイライト
  const range = sheet.getRange(headerRows + 1, 2, lastRow - headerRows, 1);
  const bgs = range.getBackgrounds();
  for (const item of over) {
    const idx = item.行 - headerRows - 1;
    bgs[idx][0] = '#f4cccc';  // 赤系ハイライト
  }
  range.setBackgrounds(bgs);

  // レポート
  let msg = `⚠️ ${sheetName}\n\n超過: ${over.length}件（赤ハイライト済み）\n\n`;
  const show = over.slice(0, 15);
  for (const item of show) {
    msg += `行${item.行}: ${item.文字数}文字/${item.バイト数}byte\n  ${item.商品名}\n`;
  }
  if (over.length > 15) msg += `\n...他 ${over.length - 15}件`;

  SpreadsheetApp.getUi().alert(msg);
}

// ====== エントリポイント（修正版） ======

/**
 * 両方のシートを一括チェック
 */
function 商品名チェック_両方() {
  console.log('--- 両方一括チェック開始 ---');
  
  // 1つずつ個別に呼び出すのではなく、エラーが起きても止まらないように実行
  const targets = [
    { name: '①商品入力シート', header: 2 },
    { name: 'Yahoo商品登録', header: 1 } // ここが「Yahoo商品登録シート」なら修正してください
  ];

  let report = "";

  targets.forEach(t => {
    try {
      // 汎用関数を直接呼び出し、戻り値（あれば）を受け取る構成にするか
      // もしくは単純に順番に実行
      商品名バイト数チェック_汎用_(t.name, t.header);
      report += `✅ ${t.name}: 実行完了\n`;
    } catch (e) {
      report += `❌ ${t.name}: エラー（${e.message}）\n`;
    }
  });

  // 最後にまとめて実行結果をログに出す（デバッグ用）
  console.log(report);
}

// 単体実行用（シート名が違うと言われる場合は、ここの ' ' 内を書き換えてください）
function 商品名チェック_商品入力シート() {
  商品名バイト数チェック_汎用_('①商品入力シート', 2);
}

function 商品名チェック_Yahoo商品登録シート() {
  // もし実際のシート名が「Yahoo商品登録シート」なら下記を書き換えてください
  商品名バイト数チェック_汎用_('Yahoo商品登録シート', 1);
}

function シート名を確認する() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const names = sheets.map(s => `「${s.getName()}」`).join('\n');
  
  SpreadsheetApp.getUi().alert("現在このファイルにあるシート名一覧:\n\n" + names);
}