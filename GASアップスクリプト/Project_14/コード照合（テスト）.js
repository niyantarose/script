/**
 * Qoo10のcodeと、yahoo変換元ファイルのcode/postage-setを照合して
 * 「code照合結果」シートに書き出す
 */
function コード照合結果を作成() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const qooSheet   = ss.getSheetByName('Qoo10');
  const yahooSheet = ss.getSheetByName('yahoo変換元ファイル');
  const resultSheet = ss.getSheetByName('code照合結果');

  if (!qooSheet || !yahooSheet || !resultSheet) {
    throw new Error('Qoo10 / yahoo変換元ファイル / code照合結果 のどれかのシートが見つかりません。');
  }

  // ===== 1) yahoo変換元ファイル側のマスタ作成 =====
  const yahooValues = yahooSheet.getDataRange().getValues();
  if (yahooValues.length < 2) {
    throw new Error('yahoo変換元ファイル にデータがありません。');
  }

  const header = yahooValues[0];

  // ヘッダー名から「code」「postage-set」の列を自動特定する
  const YAHOO_CODE_COL = header.indexOf('code');
  const YAHOO_POSTAGE_COL = header.indexOf('postage-set');

  if (YAHOO_CODE_COL === -1 || YAHOO_POSTAGE_COL === -1) {
    throw new Error('yahoo変換元ファイル で "code" または "postage-set" 列が見つかりません。');
  }

  const yahooData = yahooValues.slice(1); // 2行目以降
  const yahooMap = {}; // { code : postage-set }

  yahooData.forEach(r => {
    const code = String(r[YAHOO_CODE_COL] || '').trim();
    if (!code) return;

    const postage = r[YAHOO_POSTAGE_COL]; // BG列（postage-set）
    // 空欄は登録しない（1,2など数字が入っているものだけ）
    if (postage !== '' && postage != null) {
      yahooMap[code] = postage;
    }
  });

  // ===== 2) Qoo10側のcode一覧を取得（B列、5行目から） =====
  const lastRowQoo = qooSheet.getLastRow();
  if (lastRowQoo < 5) {
    throw new Error('Qoo10 シートにデータ行がありません。');
  }

  const qooRange = qooSheet.getRange(5, 2, lastRowQoo - 4, 1); // B5:B(最終)
  const qooValues = qooRange.getValues();

  // ===== 3) 照合して code照合結果 シートに出力 =====
  resultSheet.clearContents();

  const headerResult = [
    ['Qoo10code', 'Yahoo側code', 'Yahoo側postage-set']
  ];
  resultSheet.getRange(1, 1, 1, headerResult[0].length).setValues(headerResult);

  const output = [];
  let hitCount = 0;

  qooValues.forEach((row, idx) => {
    const qooCode = String(row[0] || '').trim();
    if (!qooCode) return; // 空行はスキップ

    const postage = yahooMap[qooCode];

    if (postage !== undefined) {
      // 一致した場合：B列にYahooのcode（=同じ）、C列にpostage-set
      output.push([qooCode, qooCode, postage]);
      hitCount++;
    } else {
      // 見つからない場合：B列「×」、C列は空欄
      output.push([qooCode, '×', '']);
    }
  });

  if (output.length > 0) {
    resultSheet.getRange(2, 1, output.length, 3).setValues(output);
  }

  Logger.log('Qoo10コード総数 = ' + output.length);
  Logger.log('Yahoo側のcode一致 = ' + hitCount);
  Logger.log('一致しない = ' + (output.length - hitCount));
}
