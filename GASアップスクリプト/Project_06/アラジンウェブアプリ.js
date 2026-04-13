// アラジンウェブアプリ.gs
// Chrome拡張機能からのデータを受け取ってシートに書き込む
// ウェブアプリとしてデプロイ必要（全員アクセス可）

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === 'updateAladinData') {
      const result = アラジンデータを更新_(data);
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: 'unknown action' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (e) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: e.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function アラジンデータを更新_(data) {
  const { itemId, genre } = data;
  if (!itemId) return { success: false, error: 'itemId missing' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ['韓国書籍', '韓国マンガ', '韓国音楽映像', '韓国グッズ'];
  const genreSheetMap = {
    '書籍': ['韓国書籍'],
    'マンガ': ['韓国マンガ'],
    '音楽映像': ['韓国音楽映像'],
    'グッズ': ['韓国グッズ']
  };

  const preferredSheets = genreSheetMap[String(genre || '').trim()] || [];
  const targetSheets = preferredSheets.concat(allSheets.filter(name => !preferredSheets.includes(name)));

  for (const sheetName of targetSheets) {
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) continue;

    const context = シートコンテキストを取得_(sh);
    if (!context.idCol) continue;

    const row = ItemId一致行を探す_(sh, context.idCol, itemId);
    if (!row) continue;

    アラジン行へ書き込む_(sh, context, row, data);
    return { success: true, sheet: sheetName, row, mode: 'updated' };
  }

  const createSheetName = targetSheets.find(name => ss.getSheetByName(name));
  if (!createSheetName) {
    return { success: false, error: '対象シートが見つかりませんでした' };
  }

  const createSheet = ss.getSheetByName(createSheetName);
  const context = シートコンテキストを取得_(createSheet);
  const row = 追加対象行を決める_(createSheet);

  if (row > createSheet.getMaxRows()) {
    createSheet.insertRowsAfter(createSheet.getMaxRows(), row - createSheet.getMaxRows());
  }

  アラジン行へ書き込む_(createSheet, context, row, data);
  return { success: true, sheet: createSheetName, row, mode: 'created' };
}

function シートコンテキストを取得_(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const 列 = {};
  headers.forEach((header, index) => {
    if (header) 列[String(header).trim()] = index + 1;
  });

  const fallbackMap = {
    '韓国書籍': {
      url: 'アラジンURL',
      isbn: 'ISBN',
      title: '商品名(原題)',
      author: '著者',
      publisher: '出版社',
      pubDate: '発売日',
      price: '原価',
      cover: 'メイン画像URL',
      additionalImages: '追加画像URL',
      description: '商品説明',
      categoryName: 'カテゴリ',
      itemId: 'アラジン商品ID'
    },
    '韓国マンガ': {
      url: 'アラジンURL',
      isbn: 'ISBN',
      title: '商品名(原題)',
      author: '著者',
      publisher: '出版社',
      pubDate: '発売日',
      price: '原価',
      cover: 'メイン画像URL',
      additionalImages: '追加画像URL',
      description: '商品説明',
      categoryName: 'カテゴリ',
      itemId: 'アラジン商品ID'
    },
    '韓国音楽映像': {
      url: 'アラジンURL',
      isbn: 'JANコード',
      title: '商品名(原題)',
      author: 'アーティスト名',
      pubDate: '発売日',
      price: '原価',
      cover: 'メイン画像URL',
      additionalImages: '追加画像URL',
      description: '商品説明',
      categoryName: 'カテゴリ',
      itemId: 'アラジン商品ID'
    },
    '韓国グッズ': {
      url: '購入URL',
      title: '商品名（原題）',
      pubDate: '発売日',
      price: '原価',
      cover: 'メイン画像URL',
      additionalImages: '追加画像URL',
      description: '商品説明',
      itemId: 'アラジン商品ID'
    }
  };

  const columnMapSource = (typeof ALADIN_COLUMN_MAP !== 'undefined' && ALADIN_COLUMN_MAP[sh.getName()])
    ? ALADIN_COLUMN_MAP[sh.getName()]
    : fallbackMap[sh.getName()] || {};

  return {
    headers,
    列,
    colMap: columnMapSource,
    idCol: 列['アラジン商品ID']
  };
}

function 追加対象行を決める_(sh) {
  const lastColumn = Math.max(sh.getLastColumn(), 1);
  const lastRow = Math.max(sh.getLastRow(), 1);
  if (lastRow < 2) return 2;

  const values = sh.getRange(2, 1, lastRow - 1, lastColumn).getDisplayValues();
  for (let i = values.length - 1; i >= 0; i--) {
    const hasVisibleValue = values[i].some(value => String(value).trim() !== '');
    if (hasVisibleValue) {
      return i + 3;
    }
  }

  return 2;
}
function ItemId一致行を探す_(sh, idCol, itemId) {
  const values = sh.getRange(2, idCol, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(itemId).trim()) {
      return i + 2;
    }
  }
  return null;
}

function アラジン行へ書き込む_(sh, context, row, data) {
  const { 列, colMap } = context;
  const setByHeader = (headerName, value) => {
    if (!headerName || !列[headerName]) return;
    if (value === null || value === undefined || value === '') return;
    sh.getRange(row, 列[headerName]).setValue(value);
  };

  const basicInfoCol = 列['基本情報'] || 列['備考'];
  const mergedDescription = 商品説明テキストを組み立てる_(data);

  setByHeader(colMap.url, data.pageUrl || '');
  setByHeader(colMap.isbn, data.isbn13 || '');
  setByHeader(colMap.title, data.title || '');
  setByHeader(colMap.author, data.author || '');
  setByHeader(colMap.publisher, data.publisher || '');
  setByHeader(colMap.pubDate, data.pubDate || '');
  setByHeader(colMap.price, data.priceSales || '');
  setByHeader(colMap.cover, data.cover || '');
  setByHeader(colMap.additionalImages, data.additionalImages || '');
  setByHeader(colMap.description, mergedDescription);
  setByHeader(colMap.categoryName, data.categoryName || '');
  setByHeader(colMap.itemId, String(data.itemId || ''));

  if (basicInfoCol && data.basicInfo) {
    const existing = sh.getRange(row, basicInfoCol).getValue();
    if (!existing) {
      sh.getRange(row, basicInfoCol).setValue(String(data.basicInfo).slice(0, 500));
    }
  }

  行表示を整える_(sh, row);
}

function 一行テキストへ整形_(value) {
  return String(value || '')
    .replace(/\r/g, '\n')
    .split('\n')
    .map(line => line.trim())
    .filter(Boolean)
    .join(' / ');
}

function 商品説明テキストを組み立てる_(data) {
  const description = 一行テキストへ整形_(data.description || '');
  const basicInfo = 一行テキストへ整形_(data.basicInfo || '');
  const parts = [];

  if (basicInfo) {
    parts.push(`[基本情報] ${basicInfo}`);
  }
  if (description) {
    parts.push(description);
  }

  return parts.join(' / ').slice(0, 4000);
}

function 行表示を整える_(sh, row) {
  const lastColumn = Math.max(sh.getLastColumn(), 1);
  const range = sh.getRange(row, 1, 1, lastColumn);
  range.setWrap(false);
  range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  range.setVerticalAlignment('middle');
  sh.setRowHeight(row, 21);
}





