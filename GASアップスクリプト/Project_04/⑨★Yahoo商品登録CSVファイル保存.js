/******************************************************
 * Yahoo商品登録シートをCSVでローカルDL（図形ボタン割り当て用）
 * 使い方：
 * 1) シートの図形ボタンを右クリック →「スクリプトを割り当て」
 * 2) CSVDL_ボタン実行 を入力
 ******************************************************/

const CSV_DL_SHEET_NAME = 'Yahoo商品登録シート';

/**
 * ★図形ボタンから呼ぶ関数（引数なし）
 */
function CSVDL_ボタン実行() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CSV_DL_SHEET_NAME);

    if (!sheet) {
      SpreadsheetApp.getUi().alert('エラー', `「${CSV_DL_SHEET_NAME}」が見つかりません`, SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow < 1 || lastCol < 1) {
      SpreadsheetApp.getUi().alert('エラー', 'シートにデータがありません', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const csv = CSVDL_convertToCSV_(data);

    const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
    const fileName = `Yahoo商品登録_${timestamp}.csv`;

    CSVDL_showDownloadDialog_(csv, fileName);

  } catch (err) {
    try {
      SpreadsheetApp.getUi().alert('エラー', String(err.message || err), SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) {
      Logger.log('[CSVDL_ERROR] ' + (err && err.stack ? err.stack : err));
    }
  }
}

/**
 * 2次元配列 → CSV文字列
 */
function CSVDL_convertToCSV_(data) {
  return data.map(row => {
    return row.map(cell => {
      if (cell === null || cell === undefined) return '';

      if (cell instanceof Date) {
        return Utilities.formatDate(cell, 'Asia/Tokyo', 'yyyy/MM/dd');
      }

      let str = String(cell);

      // カンマ、改行、ダブルクォートを含む場合はCSVエスケープ
      if (str.includes(',') || str.includes('\n') || str.includes('"')) {
        str = '"' + str.replace(/"/g, '""') + '"';
      }

      return str;
    }).join(',');
  }).join('\r\n');
}

/**
 * ダウンロードダイアログ表示（ローカル保存）
 */
function CSVDL_showDownloadDialog_(csvContent, fileName) {
  // Shift_JIS（Yahoo用）
  const sjisBlob = Utilities.newBlob('', 'text/csv', fileName)
    .setDataFromString(csvContent, 'Shift_JIS');
  const base64 = Utilities.base64Encode(sjisBlob.getBytes());

  const safeFileName = CSVDL_escapeHtml_(fileName);

  const html = `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <meta charset="UTF-8">
  <style>
    body {
      font-family: sans-serif;
      padding: 18px;
      text-align: center;
      background: #f9f9f9;
      color: #333;
    }
    .filename {
      margin: 10px 0 16px;
      font-weight: bold;
      color: #2e7d32;
      word-break: break-all;
    }
    .btn {
      display: inline-block;
      padding: 12px 20px;
      background: #4CAF50;
      color: #fff;
      border: none;
      border-radius: 6px;
      font-weight: bold;
      cursor: pointer;
    }
    .btn:hover { background: #43a047; }
    .note {
      margin-top: 12px;
      font-size: 13px;
      line-height: 1.6;
    }
    .warn { color: #d32f2f; }
    .sub { color: #777; font-size: 12px; }
  </style>
</head>
<body>
  <h3>📥 CSVを保存します</h3>
  <div class="filename">${safeFileName}</div>

  <button id="dlBtn" class="btn" type="button">ダウンロードする</button>

  <div class="note">
    自動で保存されない場合は、ボタンを押してください。<br>
    <span class="warn">※ブラウザ設定によっては保存先確認が出ます</span><br><br>
    <span class="sub">※この画面は数秒後に自動で閉じます</span>
  </div>

  <script>
    const FILE_NAME = ${JSON.stringify(fileName)};
    const BASE64 = ${JSON.stringify(base64)};

    function b64ToUint8Array(base64) {
      const binary = atob(base64);
      const len = binary.length;
      const bytes = new Uint8Array(len);
      for (let i = 0; i < len; i++) {
        bytes[i] = binary.charCodeAt(i);
      }
      return bytes;
    }

    function downloadCsv() {
      try {
        const bytes = b64ToUint8Array(BASE64);
        const blob = new Blob([bytes], { type: 'text/csv;charset=shift_jis;' });
        const url = URL.createObjectURL(blob);

        const a = document.createElement('a');
        a.href = url;
        a.download = FILE_NAME;
        a.style.display = 'none';
        document.body.appendChild(a);
        a.click();
        a.remove();

        setTimeout(() => URL.revokeObjectURL(url), 10000);
      } catch (e) {
        alert('ダウンロードに失敗しました。もう一度ボタンを押してください。');
      }
    }

    document.getElementById('dlBtn').addEventListener('click', downloadCsv);

    window.onload = function() {
      // 自動ダウンロード
      downloadCsv();

      // 5秒後に閉じる
      setTimeout(function() {
        try { google.script.host.close(); } catch (e) {}
      }, 5000);
    };
  </script>
</body>
</html>
`;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(430)
    .setHeight(260);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'CSVダウンロード');
}

function CSVDL_escapeHtml_(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}