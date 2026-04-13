/**
 * ②オプション入力シートから sub-code と options を生成
 * ★台湾商品の場合は「お取り寄せ（台湾から）」と表示（sub-code / options 両方）
 */
function サブコードとオプションを生成() {
  Logger.log('=== START サブコードとオプションを生成 ===');
  Logger.log('Active SS: ' + SpreadsheetApp.getActive().getName());

  const ss = SpreadsheetApp.getActive();
  const optSheet = ss.getSheetByName('②オプション入力シート');
  const yahooSheet = ss.getSheetByName('Yahoo商品登録シート');

  if (!optSheet || !yahooSheet) {
    throw new Error('必要なシートが見つかりません。(②オプション入力シート / Yahoo商品登録シート)');
  }

  const optData = optSheet.getDataRange().getValues();
  if (optData.length <= 1) {
    SpreadsheetApp.getUi().alert('②オプション入力シートにデータがありません。');
    return;
  }

  // ヘッダーを除いたデータ行
  const dataRows = optData.slice(1);

  // 親コードごとにグループ化
  const grouped = new Map();

  dataRows.forEach(row => {
    const parentCode  = String(row[0] || '').trim();  // A列: 親コード
    const stockType   = String(row[1] || '').trim();  // B列: 在庫種別
    const subCode     = String(row[2] || '').trim();  // C列: サブコード
    const displayName = String(row[3] || '').trim();  // D列: 表示名

    if (!parentCode && !stockType && !subCode && !displayName) return;

    if (!parentCode || !subCode || !stockType || !displayName) {
      Logger.log('スキップ行(必須欠落): ' + JSON.stringify({ parentCode, stockType, subCode, displayName }));
      return;
    }

    if (!grouped.has(parentCode)) grouped.set(parentCode, []);
    grouped.get(parentCode).push({ stockType, subCode, displayName });
  });

  // Yahoo商品登録シートの全データ
  const yahooData = yahooSheet.getDataRange().getValues();
  if (yahooData.length < 2) {
    SpreadsheetApp.getUi().alert('Yahoo商品登録シートにデータがありません。');
    return;
  }

  // ===== ヘッダー行を自動検出（1行目 or 2行目）=====
  const headerRowIndex = (() => {
    const h1 = (yahooData[0] || []).map(x => String(x || '').trim().toLowerCase());
    const h2 = (yahooData[1] || []).map(x => String(x || '').trim().toLowerCase());
    const hasCode1 = h1.includes('code');
    const hasCode2 = h2.includes('code');
    if (hasCode2 && !hasCode1) return 1; // 2行目が本ヘッダー
    return 0; // 1行目が本ヘッダー（デフォ）
  })();

  Logger.log('headerRowIndex=' + headerRowIndex);

  const yahooHeader = yahooData[headerRowIndex].map(x => String(x || '').trim());
  Logger.log('yahooHeader=' + JSON.stringify(yahooHeader));

  const findCol = (cands) => {
    const lower = yahooHeader.map(h => String(h || '').trim().toLowerCase());
    for (const cand of cands) {
      const idx = lower.indexOf(String(cand).toLowerCase());
      if (idx !== -1) return idx;
    }
    return -1;
  };

  const codeColIndex    = findCol(['code']);
  const subCodeColIndex = findCol(['sub-code', 'subcode']);
  const optionsColIndex = findCol(['options', 'option']);
  const nameColIndex    = findCol(['name', '商品名']);

  Logger.log(`codeColIndex=${codeColIndex}, subCodeColIndex=${subCodeColIndex}, optionsColIndex=${optionsColIndex}, nameColIndex=${nameColIndex}`);

  if (codeColIndex === -1 || subCodeColIndex === -1 || optionsColIndex === -1) {
    throw new Error('Yahoo商品登録シートに code / sub-code / options のいずれかの列が見つかりません。');
  }
  if (nameColIndex === -1) {
    throw new Error('Yahoo商品登録シートに name（商品名）列が見つかりません。');
  }

  // データ開始行（ヘッダーの次）
  const startRow = headerRowIndex + 1;

  let processedCount = 0;
  let taiwanCount = 0;

  for (let i = startRow; i < yahooData.length; i++) {
    const code = String(yahooData[i][codeColIndex] || '').trim();
    if (!code || !grouped.has(code)) continue;

    const productName = String(yahooData[i][nameColIndex] || '');
    const isTaiwan = /台湾|台灣/.test(productName);

    // ★ デバッグログ
    Logger.log(`row=${i+1} code=${code} productName=${productName} isTaiwan=${isTaiwan}`);

    if (isTaiwan) taiwanCount++;

    const items = grouped.get(code);

    // 表示名ごとに番号
    const displayNameMap = new Map();
    const uniqueDisplayNames = [];
    items.forEach(item => {
      if (!displayNameMap.has(item.displayName)) {
        displayNameMap.set(item.displayName, uniqueDisplayNames.length + 1);
        uniqueDisplayNames.push(item.displayName);
      }
    });

    // 在庫種別ごとに分類
    const stockTypeGroups = new Map();
    items.forEach(item => {
      if (!stockTypeGroups.has(item.stockType)) stockTypeGroups.set(item.stockType, []);
      stockTypeGroups.get(item.stockType).push(item);
    });

    // 在庫種別の順序
    const orderedStockTypes = [];
    if (stockTypeGroups.has('即納（日本在庫）')) orderedStockTypes.push('即納（日本在庫）');
if (stockTypeGroups.has('お取り寄せ（韓国から）')) orderedStockTypes.push('お取り寄せ（韓国から）');
for (const st of stockTypeGroups.keys()) {
      if (!orderedStockTypes.includes(st)) orderedStockTypes.push(st);
    }

    // ★ デバッグログ
    Logger.log('orderedStockTypes=' + JSON.stringify(orderedStockTypes));

    // sub-code（台湾は表示だけ差し替え）
    const subCodeParts = [];
    orderedStockTypes.forEach(stockType => {
      const itemsInStockType = stockTypeGroups.get(stockType) || [];

      const stockTypeLabel =
        (isTaiwan && stockType === 'お取り寄せ（韓国から）')
          ? 'お取り寄せ（台湾から）'
          : stockType;

      itemsInStockType.sort((a, b) =>
        displayNameMap.get(a.displayName) - displayNameMap.get(b.displayName)
      );

      itemsInStockType.forEach(item => {
        const num = displayNameMap.get(item.displayName);
        subCodeParts.push(
          '★在庫の設定:' + stockTypeLabel +
          '#★種類の選択:' + num + '.' + item.displayName +
          '=' + item.subCode
        );
      });
    });

    let subCodeStr = subCodeParts.join('&');

    // ★ 最終置換（念のため）
    subCodeStr = replaceKoreaToTaiwanLabel_(subCodeStr, isTaiwan);

    // options（台湾は表示だけ差し替え）
    const orderedStockTypesForOptions = orderedStockTypes.map(st =>
      (isTaiwan && st === 'お取り寄せ（韓国から）') ? 'お取り寄せ（台湾から）' : st
    );

    const line1 = ['★在庫の設定', ...orderedStockTypesForOptions].join(' ');
    const numberedChoices = uniqueDisplayNames.map((n, idx) => `${idx + 1}.${n}`);
    const line2 = ['★種類の選択', ...numberedChoices].join(' ');
    let optionsStr = line1 + '\n\n' + line2;

    // ★ 最終置換（念のため）
    optionsStr = replaceKoreaToTaiwanLabel_(optionsStr, isTaiwan);

    // ★ デバッグログ（書き込み直前）
    if (isTaiwan) {
      Logger.log(`[台湾] subCodeStr=${subCodeStr.substring(0, 100)}...`);
      Logger.log(`[台湾] optionsStr=${optionsStr.replace(/\n/g, '[改行]')}`);
    }

    // 書き込み（シート行番号は 1-based）
    yahooSheet.getRange(i + 1, subCodeColIndex + 1).setValue(subCodeStr);
    yahooSheet.getRange(i + 1, optionsColIndex + 1).setValue(optionsStr);

    processedCount++;
  }

  Logger.log('=== END サブコードとオプションを生成 ===');
  Logger.log(`processedCount=${processedCount}, taiwanCount=${taiwanCount}`);

  SpreadsheetApp.getUi().alert(
    `sub-code と options を生成しました。\n\n` +
    `処理件数: ${processedCount}件\n` +
    `うち台湾商品: ${taiwanCount}件`
  );
}

/**
 * 韓国から → 台湾から に一括置換（最終セーフティ）
 */
function replaceKoreaToTaiwanLabel_(str, isTaiwan) {
  if (!isTaiwan) return str;
  return String(str).replace(/お取り寄せ（韓国から）/g, 'お取り寄せ（台湾から）');
}