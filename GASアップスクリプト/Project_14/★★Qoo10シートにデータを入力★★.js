/** HTML をざっくりテキストにする */
function stripHtml_(text) {
  return String(text || '')
    .replace(/<[^>]+>/g, ' ')  // タグ削除
    .replace(/&nbsp;/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * 発送可能日（AA列）を判定する（今はまだ使っていないが将来用）
 * 即納在庫 → "3"
 * お取り寄せ → "14"
 * 判定できなければ "" を返す
 */
function 判定発送日AA_(itemName, descriptionHtml) {
  const title = String(itemName || '');
  const desc  = stripHtml_(descriptionHtml || '');
  const text  = (title + ' ' + desc).replace(/\s+/g, '');

  const has即納 = text.indexOf('即納在庫') !== -1 || text.indexOf('即納') !== -1;
  const has取寄 = text.indexOf('お取り寄せ') !== -1 || text.indexOf('取り寄せ') !== -1;

  if (has即納 && !has取寄) return '3';
  if (has取寄 && !has即納) return '14';

  return '';
}

/**
 * 国コードを判定して返す
 * 戻り値: 'KR','TW','CN','TH','PH','AM' など / 判定不可なら ''
 */
function 判定国コード_(itemName, descriptionHtml) {
  const title = String(itemName || '');
  const desc  = stripHtml_(descriptionHtml || '');

  // ① タイトル優先
  if (title.indexOf('韓国') !== -1 || /Korea/i.test(title)) return 'KR';
  if (title.indexOf('台湾') !== -1 || /Taiwan/i.test(title)) return 'TW';
  if (title.indexOf('中国') !== -1 || /China/i.test(title))  return 'CN';
  if (title.indexOf('フィリピン') !== -1 || /Philippines?/i.test(title)) return 'PH';
  if (title.indexOf('タイ') !== -1 || /Thailand/i.test(title)) return 'TH';

  // 英語雑誌 → AM（アメリカ）
  if (title.indexOf('英語 雑誌') !== -1 ||
      title.indexOf('英文版') !== -1 ||
      /English/i.test(title)) {
    return 'AM';
  }

  // ② 説明文の「○○の雑誌」
  if (desc.indexOf('台湾の雑誌') !== -1)        return 'TW';
  if (desc.indexOf('韓国の雑誌') !== -1)        return 'KR';
  if (desc.indexOf('中国の雑誌') !== -1)        return 'CN';
  if (desc.indexOf('フィリピンの雑誌') !== -1)  return 'PH';
  if (desc.indexOf('タイの雑誌') !== -1)        return 'TH';

  // 出版社からフィリピン
  if (/Philippines?/i.test(desc)) return 'PH';

  // ③ 言語から判定
  if (desc.indexOf('韓国語で書かれています') !== -1) return 'KR';
  // 台湾の雑誌かどうかは②で見ているので、ここは中国扱いでOK
  if (desc.indexOf('中国語で書かれています') !== -1) return 'CN';
  if (desc.indexOf('タイ語で書かれています') !== -1) return 'TH';

  // 英語 → AM
  if (desc.indexOf('英語で書かれています') !== -1 ||
      desc.indexOf('英文版') !== -1 ||
      /English/i.test(desc)) {
    return 'AM';
  }

  return '';
}

/** Qoo10シートにデータを入力（国コード判定つき + URL整形） */
function Qoo10シートにデータを入力() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qoo10Sheet = ss.getSheetByName('Qoo10');
  const yahooSheet = ss.getSheetByName('yahoo変換元ファイル');

  if (!qoo10Sheet) {
    Browser.msgBox('エラー', 'Qoo10シートが見つかりません', Browser.Buttons.OK);
    return;
  }
  if (!yahooSheet) {
    Browser.msgBox('エラー', 'yahoo変換元ファイルシートが見つかりません', Browser.Buttons.OK);
    return;
  }

  try {
    // Qoo10側のデータ最終行（B列・5行目以降で判定）
    const lastRow = qoo10Sheet.getLastRow();
    let dataLastRow = 5;
    for (let i = lastRow; i >= 5; i--) {
      const v = qoo10Sheet.getRange(i, 2).getValue(); // B列
      if (v !== '' && v !== null) {
        dataLastRow = i;
        break;
      }
    }
    if (dataLastRow < 5) {
      Browser.msgBox('エラー', 'B列（5行目以降）にデータがありません', Browser.Buttons.OK);
      return;
    }

    // yahoo側
    const yahooLastRow = yahooSheet.getLastRow();
    const yahooCodeColumn = 3;  // C列
    const yahooUrlColumn  = 90; // CL列
    const yahooDataStartRow = 2;
    if (yahooLastRow < yahooDataStartRow) {
      Browser.msgBox('エラー', 'yahoo変換元ファイルにデータがありません', Browser.Buttons.OK);
      return;
    }
    const yahooCodes = yahooSheet.getRange(
      yahooDataStartRow, yahooCodeColumn,
      yahooLastRow - yahooDataStartRow + 1, 1
    ).getValues();
    const yahooUrls = yahooSheet.getRange(
      yahooDataStartRow, yahooUrlColumn,
      yahooLastRow - yahooDataStartRow + 1, 1
    ).getValues();

    // code → URL文字列 のマップを作成
    const yahooMap = {};
    for (let i = 0; i < yahooCodes.length; i++) {
      const code = yahooCodes[i][0];
      const url  = yahooUrls[i][0];
      if (code) yahooMap[code] = url;
    }

    // Qoo10側の全データ
    const dataRowCount = dataLastRow - 4; // 5行目から
    const qoo10Data = qoo10Sheet.getRange(
      5, 1, dataRowCount, qoo10Sheet.getLastColumn()
    ).getValues();

    const eColumn  = [];
    const gColumn  = [];
    const iColumn  = [];
    const adColumn = [];
    const aeColumn = [];
    const agColumn = [];
    const rColumn  = [];

    let agEmptyCount = 0;
    const agEmptyRows = [];

    for (let i = 0; i < qoo10Data.length; i++) {
      const rowData = qoo10Data[i];
      const sellerUniqueItemId = rowData[1];  // B列
      const itemName           = rowData[4];  // E列（元のタイトル）
      const descriptionHtml    = rowData[23]; // X列

      // E列クリーニング
      let cleanedItemName = '';
      if (itemName) {
        cleanedItemName = String(itemName)
          // 1. ★～★を削除
          .replace(/★[^★]*★/g, '')
          // 2. カタカナ括弧を削除（日本在庫など）
          .replace(/[（(][ァ-ヴー・\s]+[）)]/g, '')
          // 3. 英数字の後のカタカナ読み仮名を削除
          .replace(/([A-Za-z0-9+\-]+)\s+[ァ-ヴー・]+/g, '$1')
          // 4. 「特別付録」→「付録」
          .replace(/特別付録/g, '付録')
          // 5. 連続スペースを1つに
          .replace(/\s+/g, ' ')
          .trim();
      }
      eColumn.push([cleanedItemName]);

      // 固定値
      gColumn.push(['Y']);
      iColumn.push(['2050-01-01']);

      // ★★★★★ ここが追加：中古/古本 判定で AD を切り替える ★★★★★
      // 「E列の商品名に中古 or 古本が含まれたら AD=4、それ以外は AD=1」
      const checkText = (String(itemName || '') + ' ' + String(cleanedItemName || '')).toLowerCase();
      const isUsedBook =
        checkText.indexOf('中古') !== -1 ||
        checkText.indexOf('古本') !== -1 ||
        checkText.indexOf('古書') !== -1 ||
        checkText.indexOf('used') !== -1; // 英語表記が混ざる場合の保険

      adColumn.push([isUsedBook ? '4' : '1']);
      // ★★★★★ 追加ここまで ★★★★★

      aeColumn.push(['2']);

      // 国コード
      const countryCode = 判定国コード_(itemName, descriptionHtml);
      if (!countryCode) {
        agEmptyCount++;
        agEmptyRows.push(i + 5); // 行番号（5行目スタート）
      }
      agColumn.push([countryCode]);

      // URL整形：;区切り → 有効なURLだけ $$ で連結
      let convertedUrl = '';
      if (sellerUniqueItemId && yahooMap[sellerUniqueItemId]) {
        const originalUrl = String(yahooMap[sellerUniqueItemId] || '');

        const urlList = originalUrl
          .split(';')
          .map(s => s.trim())
          .filter(s => s !== '');

        convertedUrl = urlList.join('$$');
      }
      rColumn.push([convertedUrl]);
    }

    // 一括書き込み
    if (gColumn.length > 0) {
      qoo10Sheet.getRange(5, 5,  eColumn.length, 1).setValues(eColumn);    // E
      qoo10Sheet.getRange(5, 7,  gColumn.length, 1).setValues(gColumn);    // G
      qoo10Sheet.getRange(5, 9,  iColumn.length, 1).setValues(iColumn);    // I
      qoo10Sheet.getRange(5, 30, adColumn.length, 1).setValues(adColumn);  // AD
      qoo10Sheet.getRange(5, 31, aeColumn.length, 1).setValues(aeColumn);  // AE
      qoo10Sheet.getRange(5, 33, agColumn.length, 1).setValues(agColumn);  // AG
      qoo10Sheet.getRange(5, 18, rColumn.length, 1).setValues(rColumn);    // R
    }

    // AG列チェックのメッセージ
    if (agEmptyCount > 0) {
      const sample = agEmptyRows.slice(0, 10).join(', ');
      Browser.msgBox(
        '注意',
        'AG列（発送国コード）が空欄の行が ' + agEmptyCount + ' 行あります。\n' +
        '例：行 ' + sample + '\n\n手動で国を確認して入力してください。',
        Browser.Buttons.OK
      );
    } else {
      Browser.msgBox(
        '完了',
        dataRowCount + '行のデータを処理しました（AG列もすべて埋まりました）。',
        Browser.Buttons.OK
      );
    }

  } catch (e) {
    Logger.log('エラー: ' + e.message);
    Browser.msgBox('エラー', e.message + '\n' + e.stack, Browser.Buttons.OK);
  }
}
