// ============================================================
// ヘルパー：シート名の一部でシートを探す ＋ ログ出力
// ============================================================
function findSheetByKeyword_(ss, keyword) {
  const sheets = ss.getSheets();
  const allNames = sheets.map(s => s.getName());
  Logger.log(`findSheetByKeyword_: keyword="${keyword}", allSheets=${JSON.stringify(allNames)}`);

  const hit = sheets.find(sh => sh.getName().indexOf(keyword) !== -1);
  if (hit) {
    Logger.log(`  -> hit sheet: "${hit.getName()}"`);
  } else {
    Logger.log('  -> hit sheet: (none)');
  }
  return hit;
}

// ============================================================
// ヘルパー：L列テキストから配送パターンキーを作る
//   例）節約便【360円】 → "節約便_360"
//       佐川急便        → "佐川"
// ============================================================
function 抽出配送パターンキー_(text) {
  if (!text) return '';

  const t = String(text)
    .replace(/<[^>]+>/g, '') // HTMLタグ除去
    .replace(/\s+/g, ' ')
    .trim();

  // 節約便（郵便受けに配達）---【360円】 など
  if (t.indexOf('節約便') !== -1) {
    const m = t.match(/【(\d+)円】/);
    if (m) return `節約便_${m[1]}`;
    return '節約便';
  }

  // 佐川急便
  if (t.indexOf('佐川急便') !== -1 || t.indexOf('佐川') !== -1) {
    return '佐川';
  }

  return '';
}

// 数字だけかどうか
function isNumericCode_(v) {
  if (v === null || v === undefined) return false;
  const s = String(v).trim();
  return /^[0-9]+$/.test(s);
}

// ============================================================
// メイン：送料コード反映
//   ・Qoo10：4行ヘッダー＋5行目〜データ
//   ・yahoo変換元：1行ヘッダー＋2行目〜データ
// ============================================================
function 送料コード反映() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const qoo10Sheet  = findSheetByKeyword_(ss, 'Qoo10');
  const yahooSheet  = findSheetByKeyword_(ss, 'yahoo変換元ファイル');
  const groupSheet  = findSheetByKeyword_(ss, 'yahoo配送グループ');

  if (!qoo10Sheet)  throw new Error('Qoo10 シートが見つかりません');
  if (!yahooSheet)  throw new Error('yahoo変換元ファイル シートが見つかりません');
  if (!groupSheet)  throw new Error('yahoo配送グループ シートが見つかりません');

  Logger.log('=== 送料コード反映 開始 ===');

  // ----------------------------------------------------------
  // 1) yahoo配送グループ: postage-set → Qoo10送料コード
  // ----------------------------------------------------------
  const groupValues = groupSheet.getDataRange().getValues();
  if (groupValues.length <= 1) {
    Logger.log('yahoo配送グループ にデータ行がありません。処理終了。');
    return;
  }
  const groupHeader = groupValues[0];

  const COL_G_POSTAGE = groupHeader.findIndex(h =>
    typeof h === 'string' && h.indexOf('postage-set') !== -1
  );
  const COL_G_QOO10   = groupHeader.findIndex(h =>
    typeof h === 'string' && h.indexOf('Qoo10送料コード') !== -1
  );

  if (COL_G_POSTAGE < 0 || COL_G_QOO10 < 0) {
    throw new Error('yahoo配送グループ: 「postage-set」または「Qoo10送料コード」列が見つかりません');
  }

  const groupMap = {};  // postage-set → Qoo10送料コード
  groupValues.slice(1).forEach(r => {
    const key  = r[COL_G_POSTAGE];
    const code = r[COL_G_QOO10];
    if (key !== '' && key != null && code !== '' && code != null) {
      groupMap[String(key)] = code;
    }
  });
  Logger.log(`groupMap 件数 = ${Object.keys(groupMap).length}`);

  // ----------------------------------------------------------
  // 2) yahoo変換元ファイル：code → postage-set を決める
  //    （L列の配送パターンを見て、空欄推奨などは上書き）
  // ----------------------------------------------------------
  const yahooValues = yahooSheet.getDataRange().getValues();
  if (yahooValues.length <= 1) {
    Logger.log('yahoo変換元ファイル にデータ行がありません。処理終了。');
    return;
  }
  const yahooHeader = yahooValues[0];

  const COL_Y_CODE    = yahooHeader.indexOf('code');        // C列
  const COL_Y_POSTAGE = yahooHeader.indexOf('postage-set'); // BG列
  const COL_Y_LTEXT   = 11; // L列（A=0 → 11）

  if (COL_Y_CODE < 0 || COL_Y_POSTAGE < 0) {
    throw new Error('yahoo変換元ファイル: 「code」または「postage-set」列が見つかりません');
  }
  Logger.log(`Yahoo 列Index: code=${COL_Y_CODE}, postage-set=${COL_Y_POSTAGE}, L列=${COL_Y_LTEXT}`);

  // まず「Lテキスト → 正しいpostage-set」の対応表を作る
  const 配送パターンマップ = {}; // 例: 節約便_360 → 3, 佐川 → 1 など
  yahooValues.slice(1).forEach(r => {
    const ltext   = r[COL_Y_LTEXT];
    const postage = r[COL_Y_POSTAGE];
    if (postage === '' || postage == null) return;

    const key = 抽出配送パターンキー_(ltext);
    if (key) {
      配送パターンマップ[key] = postage;
    }
  });
  Logger.log('配送パターンマップ=' + JSON.stringify(配送パターンマップ));

  // 次に各 code について最終的な postage-set を決める
  const yahooMap = {};   // code → 最終postage-set
  const yahooLMap = {};  // code → Lテキスト（デバッグ用）

  yahooValues.slice(1).forEach(r => {
    const code  = String(r[COL_Y_CODE]).trim();
    if (!code) return;

    let   postage = r[COL_Y_POSTAGE];
    const ltext   = r[COL_Y_LTEXT];
    const key     = 抽出配送パターンキー_(ltext);

    const origPostage = postage;

    // 既存postage-set → Qoo10送料コードを確認
    let shipFromOrig = null;
    if (postage !== '' && postage != null && groupMap[String(postage)] !== undefined) {
      shipFromOrig = groupMap[String(postage)];
    }

    // 「空欄推奨」など数字でない送料コードの場合は L列パターンで上書き
    if (
      !postage ||                             // 空
      !shipFromOrig ||                        // グループ表にない
      !isNumericCode_(shipFromOrig)           // 数字以外（空欄推奨など）
    ) {
      if (key && 配送パターンマップ[key] != null) {
        postage = 配送パターンマップ[key];
        Logger.log(
          `postage補正: code=${code}, orig=${origPostage}, key=${key}, new=${postage}`
        );
      }
    }

    if (postage !== '' && postage != null) {
      yahooMap[code]  = postage;
      yahooLMap[code] = ltext;
    }
  });

  Logger.log(`yahooMap 件数 = ${Object.keys(yahooMap).length}`);

  // ----------------------------------------------------------
  // 3) Qoo10シート側：Y列(Shipping_number)を書き換える
  // ----------------------------------------------------------
  const qoo10Values = qoo10Sheet.getDataRange().getValues();
  if (qoo10Values.length <= 4) {
    Logger.log('Qoo10 シートにデータ行がありません。処理終了。');
    return;
  }
  const qHeader = qoo10Values[3]; // 4行目がヘッダー

  const COL_Q_SELLER = qHeader.findIndex(h =>
    typeof h === 'string' &&
    (h === 'seller_unique_item_id' || h.indexOf('販売者商品コード') !== -1)
  );
  const COL_Q_SHIPPING = qHeader.findIndex(h =>
    typeof h === 'string' &&
    (h === 'Shipping_number' || h.indexOf('送料コード') !== -1)
  );

  if (COL_Q_SELLER < 0 || COL_Q_SHIPPING < 0) {
    throw new Error('Qoo10: 「seller_unique_item_id」または「Shipping_number(送料コード)」列が見つかりません');
  }

  let updateCount = 0;

  // 5行目から最終行までがデータ
  const lastRow = qoo10Sheet.getLastRow();
  for (let row = 5; row <= lastRow; row++) {
    const idx      = row - 1;
    const rowArray = qoo10Values[idx];
    const seller   = String(rowArray[COL_Q_SELLER] || '').trim();
    if (!seller) continue;

    const postage = yahooMap[seller];
    if (!postage) continue;

    const shipCode = groupMap[String(postage)];
    if (!shipCode) continue;

    qoo10Sheet.getRange(row, COL_Q_SHIPPING + 1).setValue(shipCode);
    updateCount++;
  }

  Logger.log(`更新行数 = ${updateCount}`);

  // ----------------------------------------------------------
  // 4) 662・664 行のデバッグログ（何が入ったか確認）
  // ----------------------------------------------------------
  [662, 664].forEach(rowNo => {
    if (rowNo > lastRow) return;
    const idx = rowNo - 1;
    const seller = String(qoo10Values[idx][COL_Q_SELLER] || '').trim();
    const ship   = qoo10Sheet.getRange(rowNo, COL_Q_SHIPPING + 1).getValue();
    const yPost  = yahooMap[seller];
    Logger.log(
      `DEBUG row=${rowNo}: seller="${seller}", Shipping="${ship}", ` +
      `final_postage-set="${yPost}"`
    );
  });

  // ----------------------------------------------------------
  // 5) 送料コードチェック（空欄 & 数字以外）＋ダイアログ表示
  // ----------------------------------------------------------
  const ui = SpreadsheetApp.getUi();
  let blankCount = 0;
  let nonNumCount = 0;
  const nonNumRows = [];

  for (let row = 5; row <= lastRow; row++) {
    const val = qoo10Sheet.getRange(row, COL_Q_SHIPPING + 1).getValue();
    const s = String(val || '').trim();
    if (!s) {
      blankCount++;
    } else if (!isNumericCode_(s)) {
      nonNumCount++;
      if (nonNumRows.length < 10) {
        nonNumRows.push(row);
      }
    }
  }

  let msg =
    '空欄の行: ' + blankCount + '件\n' +
    '数字以外が入っている行: ' + nonNumCount + '件\n';
  if (nonNumRows.length > 0) {
    msg += '\n数字以外行の例: ' + nonNumRows.join(', ');
  }

  ui.alert('送料コードチェック結果', msg, ui.ButtonSet.OK);

  Logger.log('=== 送料コード反映 完了 ===');
}
