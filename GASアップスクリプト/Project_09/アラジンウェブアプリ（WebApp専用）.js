// アラジンウェブアプリ.gs
// Chrome拡張機能からのデータを受け取ってシートに書き込む
// ウェブアプリとしてデプロイ必要（全員アクセス可）
// 共通マスターファイルに配置して、カテゴリごとに各スプレッドシートへ振り分ける

const ALADIN_SPREADSHEET_DESTINATION_MAP = {
  '韓国書籍': {
    spreadsheetIdOrUrl: 'https://docs.google.com/spreadsheets/d/1XjvUhDvZF-xka2HpYE70iasXuhOzUdehXjQlVGR6WPc/edit?gid=545067170#gid=545067170',
    spreadsheetLabel: '商品登録よろしく_韓国書籍音楽映像'
  },
  '韓国音楽映像': {
    spreadsheetIdOrUrl: 'https://docs.google.com/spreadsheets/d/1XjvUhDvZF-xka2HpYE70iasXuhOzUdehXjQlVGR6WPc/edit?gid=545067170#gid=545067170',
    spreadsheetLabel: '商品登録よろしく_韓国書籍音楽映像'
  },
  '韓国雑誌': {
    spreadsheetIdOrUrl: 'https://docs.google.com/spreadsheets/d/1XjvUhDvZF-xka2HpYE70iasXuhOzUdehXjQlVGR6WPc/edit?gid=545067170#gid=545067170',
    spreadsheetLabel: '商品登録よろしく_韓国書籍音楽映像'
  },
  '韓国マンガ': {
    spreadsheetIdOrUrl: 'https://docs.google.com/spreadsheets/d/1fItni5nPV_tLshtKQBXe8_Ag965G-nBn25JPvgg6BU4/edit?gid=1100321429#gid=1100321429',
    spreadsheetLabel: '商品登録よろしく_韓国マンガ'
  },
  '韓国グッズ': {
    spreadsheetIdOrUrl: 'https://docs.google.com/spreadsheets/d/1BwNNvx0PakFzugprDJKmmZ5tjDTZijKZ8VH-anTuwDU/edit?gid=1562610398#gid=1562610398',
    spreadsheetLabel: '商品登録よろしく_韓国グッズ'
  }
};

const ALADIN_ALL_TARGET_SHEETS = ['韓国書籍', '韓国雑誌', '韓国マンガ', '韓国音楽映像', '韓国グッズ'];
const ALADIN_GENRE_SHEET_MAP = {
  '書籍': ['韓国書籍'],
  '雑誌': ['韓国雑誌'],
  'マンガ': ['韓国マンガ'],
  '音楽映像': ['韓国音楽映像'],
  'グッズ': ['韓国グッズ']
};

const ALADIN_APPEND_VALUE_HEADERS = [
  'アラジン商品コード', 'サイト商品コード', 'ItemId',
  'アラジンURL', '購入URL', 'リンク',
  'ISBN', 'JANコード',
  '商品名(日本語)', '商品名（日本語）', '日本語タイトル',
  '商品名(原題)', '商品名（原題）', '作品名（原題）', '原題タイトル', '原題商品タイトル',
  '商品名(タイトル)', '商品名（タイトル）',
  '雑誌名', '著者', '作者', '出版社', 'アーティスト名', 'アーティスト', '事務所/レーベル',
  '発売日', '原価', 'メイン画像URL', 'メイン画像', '追加画像URL', '追加画像',
  '商品説明', '備考', 'カテゴリ', 'ジャンル分類', '特典メモ', '付録情報',
  '表紙情報', '年', '月', '号数'
];

const ALADIN_LOOKUP_WRITEBACK_HEADERS = [
  '日本語タイトル',
  '作品名（日本語）',
  '日本語タイトル照会結果',
  '日本語タイトル照会元',
  '日本語タイトル照会ログ',
  '検索用正規化タイトル'
];
const ALADIN_MASTER_SPREADSHEET_ID = '1ZAy5wtQCq1ixl47MMEIGnrVWeae-F_NfN34lSOUsN5M';
const ALADIN_JAPANESE_TITLE_DICTIONARY_SHEET_NAME = '日本語タイトル辞書';
const ALADIN_BUILTIN_TITLE_ALIAS_DICTIONARY = {
  '韓国|まんが|나의히어로아카데미아': '僕のヒーローアカデミア',
  '韓国|グッズ|나의히어로아카데미아': '僕のヒーローアカデミア',
  '韓国|まんが|하이큐': 'ハイキュー!!',
  '韓国|グッズ|하이큐': 'ハイキュー!!',
  '韓国|まんが|주술회전': '呪術廻戦',
  '韓国|グッズ|주술회전': '呪術廻戦'
};
let _aladinJapaneseTitleDictionaryCache = null;

/** スクリプトプロパティに台湾側など「完全版」lookupJapaneseTitle を実装した Web アプリ URL を入れると転送する */
const SHARED_TITLE_LOOKUP_WEBAPP_URL_PROPERTY = 'SHARED_TITLE_LOOKUP_WEBAPP_URL';

function lookupJapaneseTitleViaSharedWebApp_(payload) {
  var url = PropertiesService.getScriptProperties().getProperty(SHARED_TITLE_LOOKUP_WEBAPP_URL_PROPERTY);
  url = url ? String(url).trim() : '';
  if (!url || !/^https:\/\/script\.(google\.com|googleusercontent\.com)\//i.test(url)) {
    return null;
  }
  try {
    var res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      muteHttpExceptions: true,
      followRedirects: true,
      payload: JSON.stringify(Object.assign({ action: 'lookupJapaneseTitle' }, payload || {})),
    });
    var code = res.getResponseCode();
    var text = res.getContentText() || '';
    if (code < 200 || code >= 300) {
      Logger.log('SHARED_TITLE_LOOKUP forward HTTP ' + code + ' body=' + text.substring(0, 300));
      return null;
    }
    return JSON.parse(text);
  } catch (error) {
    Logger.log('SHARED_TITLE_LOOKUP forward error: ' + error.message);
    return null;
  }
}

function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    const data = JSON.parse(raw);

    if (data.action === 'lookupJapaneseTitle') {
      const result = lookupJapaneseTitleAction_(data);
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === 'upsertProductWithLookup') {
      const result = upsertProductWithLookupAction_(data);
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === 'updateAladinData') {
      const result = アラジンデータを更新_(data);
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: 'unknown action' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function lookupText_(value) {
  return String(value || '').trim();
}

function lookupTitleKey_(value) {
  return lookupText_(value)
    .toLowerCase()
    .replace(/[\s\u3000"'“”‘’・･:：!！?？,，、.．\-ー‐―–—~〜～/／\\|()\[\]{}【】「」『』<>《》]/g, '');
}

function lookupLanguageKey_(value) {
  const text = lookupText_(value);
  if (/韓国|韩国|韓語|韩语|kr|korean/i.test(text)) return '韓国';
  if (/台湾|台灣|繁体|繁體|tw/i.test(text)) return '台湾';
  return text || '韓国';
}

function lookupCategoryKey_(value) {
  const text = lookupText_(value);
  if (/bl/i.test(text)) return 'まんが';
  if (/まんが|マンガ|漫画|漫畫|comic|manga|webtoon|만화/i.test(text)) return 'まんが';
  if (/goods|グッズ|굿즈|アクリル|아크릴|키링|포토카드|스티커/i.test(text)) return 'グッズ';
  if (/雑誌|magazine|잡지|매거진/i.test(text)) return '雑誌';
  if (/音楽|映像|cd|dvd|blu|ost|album|음반/i.test(text)) return '音楽映像';
  if (/書籍|小説|소설|novel|book|도서/i.test(text)) return '書籍';
  return text || '書籍';
}

function lookupCategoriesForAnalysis_(titleAnalysis, data) {
  const itemType = lookupText_(titleAnalysis && titleAnalysis.itemType);
  const genre = 補正ジャンルを判定_(data || {});
  const category = lookupText_(titleAnalysis && titleAnalysis.category) || lookupText_(data && data.categoryName) || genre;
  const categories = [];

  if (itemType === 'manga' || itemType === 'bl_manga') categories.push('まんが');
  if (itemType === 'goods') categories.push('グッズ', 'まんが');
  if (itemType === 'light_novel' || itemType === 'novel_book') categories.push('書籍', 'まんが');
  if (itemType === 'magazine') categories.push('雑誌');
  if (itemType === 'music_video') categories.push('音楽映像');
  categories.push(lookupCategoryKey_(category));
  categories.push(lookupCategoryKey_(genre));
  return categories.filter(function(categoryName, index, array) {
    return categoryName && array.indexOf(categoryName) === index;
  });
}

function parseDictionaryAliases_(value) {
  return String(value || '')
    .split(/[\r\n;；,，、/／|｜]+/)
    .map(function(part) { return lookupText_(part); })
    .filter(Boolean);
}

function getAladinJapaneseTitleDictionaryIndex_() {
  if (_aladinJapaneseTitleDictionaryCache) return _aladinJapaneseTitleDictionaryCache;

  const byKey = {};
  try {
    const master = SpreadsheetApp.openById(ALADIN_MASTER_SPREADSHEET_ID);
    const sheet = master.getSheetByName(ALADIN_JAPANESE_TITLE_DICTIONARY_SHEET_NAME);
    if (!sheet) throw new Error('dictionary sheet missing');

    const lastRow = sheet.getLastRow();
    const lastColumn = Math.max(sheet.getLastColumn(), 1);
    if (lastRow >= 2) {
      const values = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
      const headerMap = {};
      values[0].forEach(function(header, idx) {
        const key = String(header || '').replace(/\s+/g, '').toLowerCase();
        if (key) headerMap[key] = idx;
      });

      const languageIdx = headerMap['言語'];
      const categoryIdx = headerMap['カテゴリ'];
      const originalTitleIdx = headerMap['原題タイトル'];
      const japaneseTitleIdx = headerMap['日本語タイトル'];
      const aliasIdx = headerMap['別名'];

      values.slice(1).forEach(function(row) {
        const language = lookupLanguageKey_(row[languageIdx]);
        const category = lookupCategoryKey_(row[categoryIdx]);
        const japaneseTitle = lookupText_(row[japaneseTitleIdx]);
        const candidates = [row[originalTitleIdx] || ''].concat(parseDictionaryAliases_(aliasIdx >= 0 ? row[aliasIdx] : ''));
        if (!language || !category || !japaneseTitle) return;
        candidates.forEach(function(candidate) {
          const titleKey = lookupTitleKey_(candidate);
          if (!titleKey) return;
          const key = [language, category, titleKey].join('|');
          if (!byKey[key]) byKey[key] = [];
          byKey[key].push({
            japaneseTitle: japaneseTitle,
            provider: 'titleAliasDictionary',
            score: 1000,
            note: 'alias dictionary exact match'
          });
        });
      });
    }
  } catch (error) {
    // 辞書シートが読めない場合でも、内蔵最低限辞書と外部provider結果で処理を続ける。
    Logger.log('Aladin Japanese title dictionary unavailable: ' + (error && error.message ? error.message : error));
  }

  Object.keys(ALADIN_BUILTIN_TITLE_ALIAS_DICTIONARY).forEach(function(key) {
    if (!byKey[key]) byKey[key] = [];
    byKey[key].push({
      japaneseTitle: ALADIN_BUILTIN_TITLE_ALIAS_DICTIONARY[key],
      provider: 'titleAliasDictionary',
      score: 1000,
      note: 'builtin alias dictionary exact match'
    });
  });

  _aladinJapaneseTitleDictionaryCache = { byKey: byKey };
  return _aladinJapaneseTitleDictionaryCache;
}

function lookupTitleAliasDictionary_(language, categories, searchKeys) {
  const dictionary = getAladinJapaneseTitleDictionaryIndex_();
  for (let keyIndex = 0; keyIndex < searchKeys.length; keyIndex += 1) {
    const titleKey = lookupTitleKey_(searchKeys[keyIndex]);
    if (!titleKey) continue;
    for (let categoryIndex = 0; categoryIndex < categories.length; categoryIndex += 1) {
      const category = categories[categoryIndex];
      const matches = dictionary.byKey[[language, category, titleKey].join('|')] || [];
      if (matches.length) {
        return Object.assign({}, matches[0], {
          normalizedSearchTitle: searchKeys[keyIndex],
          extractedWorkTitle: searchKeys[keyIndex]
        });
      }
    }
  }
  return null;
}

function lookupSearchKeys_(payload) {
  const analysis = payload && payload.titleAnalysis || {};
  const rawItem = payload && (payload.rawItem || payload) || {};
  const keys = [
    analysis.normalizedSearchTitle,
    analysis.extractedWorkTitle,
    analysis.originalTitle,
    analysis.originalProductTitle,
    rawItem.worksTitle,
    rawItem.productTitle,
    rawItem.title,
    payload && payload.title
  ];
  return keys
    .map(lookupText_)
    .filter(function(key, index, array) {
      return key && array.indexOf(key) === index;
    });
}

function lookupJapaneseTitleFromPayload_(payload) {
  var forwarded = lookupJapaneseTitleViaSharedWebApp_(payload);
  if (forwarded && typeof forwarded === 'object' && forwarded.success !== false) {
    return forwarded;
  }

  const analysis = payload && payload.titleAnalysis || {};
  const rawItem = payload && (payload.rawItem || payload) || {};
  const language = lookupLanguageKey_(analysis.language || rawItem.language || '韓国');
  const categories = lookupCategoriesForAnalysis_(analysis, rawItem);
  const searchKeys = lookupSearchKeys_(payload);
  const normalizedSearchTitle = searchKeys[0] || '';
  const trace = [];
  const errors = [];

  try {
    const dictHit = lookupTitleAliasDictionary_(language, categories, searchKeys);
    if (dictHit) {
      trace.push('dictionary:hit(alias=' + lookupTitleKey_(dictHit.normalizedSearchTitle || normalizedSearchTitle) + ')');
      return {
        success: true,
        lookup: {
          status: 'resolved',
          japaneseTitle: dictHit.japaneseTitle,
          provider: dictHit.provider,
          normalizedSearchTitle: normalizedSearchTitle,
          extractedWorkTitle: analysis.extractedWorkTitle || normalizedSearchTitle,
          score: dictHit.score || 1000,
          trace: trace.join(' | '),
          candidates: [dictHit],
          errors: errors
        }
      };
    }
    trace.push('dictionary:miss');
  } catch (error) {
    errors.push('dictionary:' + (error && error.message ? error.message : error));
    trace.push('dictionary:error');
  }

  const providerOrder = analysis.providerHint && Array.isArray(analysis.providerHint.preferredOrder)
    ? analysis.providerHint.preferredOrder
    : ['dictionary', 'bookwalker', 'amazon'];
  providerOrder.forEach(function(provider) {
    if (provider === 'dictionary') return;
    trace.push(provider + ':not_configured');
  });

  return {
    success: true,
    lookup: {
      status: errors.length ? 'partial_error' : 'not_found',
      japaneseTitle: '',
      provider: '',
      normalizedSearchTitle: normalizedSearchTitle,
      extractedWorkTitle: analysis.extractedWorkTitle || normalizedSearchTitle,
      score: 0,
      trace: trace.join(' | '),
      candidates: [],
      errors: errors
    }
  };
}

function lookupJapaneseTitleAction_(payload) {
  return lookupJapaneseTitleFromPayload_(payload || {});
}

function normalizeLookupResult_(lookup) {
  const source = lookup && lookup.lookup ? lookup.lookup : lookup;
  if (!source) return null;
  return {
    status: source.status || (source.japaneseTitle || source.title ? 'resolved' : 'not_found'),
    japaneseTitle: source.japaneseTitle || source.title || '',
    provider: source.provider || source.source || '',
    normalizedSearchTitle: source.normalizedSearchTitle || '',
    extractedWorkTitle: source.extractedWorkTitle || '',
    score: source.score || 0,
    trace: source.trace || source.log || '',
    candidates: Array.isArray(source.candidates) ? source.candidates : [],
    errors: Array.isArray(source.errors) ? source.errors : []
  };
}

function shouldRefreshLookup_(lookup, titleAnalysis) {
  const normalized = normalizeLookupResult_(lookup);
  if (!normalized) return true;
  const lookupKey = lookupText_(normalized.normalizedSearchTitle);
  const analysisKey = lookupText_(titleAnalysis && titleAnalysis.normalizedSearchTitle);
  return Boolean(analysisKey && lookupKey && lookupKey !== analysisKey);
}

function upsertProductWithLookupAction_(payload) {
  const rawItem = payload && payload.rawItem || {};
  const titleAnalysis = payload && payload.titleAnalysis || {};
  const lookup = shouldRefreshLookup_(payload && payload.japaneseTitleLookup, titleAnalysis)
    ? lookupJapaneseTitleFromPayload_(payload).lookup
    : normalizeLookupResult_(payload.japaneseTitleLookup);

  const data = Object.assign({}, rawItem, {
    action: 'updateAladinData',
    itemId: rawItem.itemId || payload.itemId || '',
    genre: rawItem.genre || rawItem.ジャンル || '',
    title: rawItem.title || titleAnalysis.originalProductTitle || titleAnalysis.rawTitle || '',
    worksTitle: rawItem.worksTitle || titleAnalysis.extractedWorkTitle || titleAnalysis.normalizedSearchTitle || '',
    productTitle: rawItem.productTitle || titleAnalysis.originalProductTitle || rawItem.title || '',
    magazineName: rawItem.magazineName || titleAnalysis.magazineName || '',
    author: rawItem.author || titleAnalysis.author || '',
    categoryName: rawItem.categoryName || titleAnalysis.category || '',
    titleAnalysis: titleAnalysis,
    japaneseTitleLookup: lookup
  });

  const result = アラジンデータを更新_(data);
  result.lookup = lookup;
  result.source = payload && payload.source || 'aladin';
  result.warnings = result.warnings || [];
  return result;
}

function アラジンデータを更新_(data) {
  const itemId = data && data.itemId;
  const genre = 補正ジャンルを判定_(data);
  const normalizedData = Object.assign({}, data, { genre: genre });
  if (!itemId) return { success: false, error: 'itemId missing' };

  const targetSheets = 保存先シート候補を取得_(genre);
  if (!targetSheets.length) {
    return { success: false, error: 保存先未設定メッセージを組み立て_(genre) };
  }

  for (const target of targetSheets) {
    const sh = target.sheet;
    if (!sh || sh.getLastRow() < 2) continue;

    日本語タイトル照会列を確保_(sh);
    const context = シートコンテキストを取得_(sh);
    if (!context.idCol && !context.isbnCol) continue;

    const row = 既存商品行を探す_(sh, context, normalizedData);
    if (!row) continue;

    アラジン行へ書き込む_(sh, context, row, normalizedData);
    return {
      success: true,
      sheet: target.sheetName,
      spreadsheetId: target.spreadsheetId,
      spreadsheetName: target.spreadsheetName,
      row,
      resolvedGenre: genre,
      mode: 'updated'
    };
  }

  const preferredSheets = ALADIN_GENRE_SHEET_MAP[String(genre || '').trim()] || [];
  const createTarget = targetSheets.find(target => preferredSheets.includes(target.sheetName)) || targetSheets[0];
  if (!createTarget || !createTarget.sheet) {
    return { success: false, error: '対象シートが見つかりませんでした' };
  }

  const createSheet = createTarget.sheet;
  日本語タイトル照会列を確保_(createSheet);
  const context = シートコンテキストを取得_(createSheet);
  const row = 追加対象行を決める_(createSheet, context);

  if (row > createSheet.getMaxRows()) {
    createSheet.insertRowsAfter(createSheet.getMaxRows(), row - createSheet.getMaxRows());
  }

  アラジン行へ書き込む_(createSheet, context, row, normalizedData);
  return {
    success: true,
    sheet: createTarget.sheetName,
    spreadsheetId: createTarget.spreadsheetId,
    spreadsheetName: createTarget.spreadsheetName,
    row,
    resolvedGenre: genre,
    mode: 'created'
  };
}

function 補正ジャンルを判定_(data) {
  const requestedGenre = String(data && data.genre || '').trim();
  const categoryName = String(data && data.categoryName || '').toLowerCase();
  const mallType = String(data && data.mallType || '').toLowerCase();
  const title = String(data && data.title || '').toLowerCase();
  const combined = [requestedGenre, categoryName, mallType, title].join(' ');

  const hasMagazineKeyword = /magazine|잡지|매거진|雑誌/i.test(combined);
  const hasMagazineIssuePattern = /\b20\d{2}[./-]\d{1,2}\b/.test(title) || /\b\d{4}年\s*\d{1,2}月\b/.test(title) || /\b\d{4}년\s*\d{1,2}월\b/.test(title);
  const hasMagazineBrand = /(elle|vogue|gq|esquire|allure|bazaar|harper'?s bazaar|marie claire|cosmopolitan|dazed|arena|w korea|ceci|1st look|cine21|maxim|singles|nylon|the star|star1|men'?s health)/i.test(title);

  if (requestedGenre === '雑誌' || hasMagazineKeyword || (hasMagazineIssuePattern && hasMagazineBrand)) {
    return '雑誌';
  }

  return requestedGenre || '書籍';
}

const MAGAZINE_NAME_PATTERNS_ = [
  { name: 'DAZED & CONFUSED', pattern: /dazed\s*&\s*confused/i },
  { name: "HARPER'S BAZAAR", pattern: /harper'?s\s*bazaar/i },
  { name: 'BAZAAR', pattern: /\bbazaar\b/i },
  { name: 'MARIE CLAIRE', pattern: /marie\s*claire/i },
  { name: 'COSMOPOLITAN', pattern: /cosmopolitan/i },
  { name: 'ESQUIRE', pattern: /esquire/i },
  { name: 'ALLURE', pattern: /allure/i },
  { name: 'VOGUE', pattern: /vogue/i },
  { name: 'ELLE', pattern: /elle/i },
  { name: 'GQ', pattern: /\bgq\b/i },
  { name: 'ARENA', pattern: /arena/i },
  { name: 'W KOREA', pattern: /\bw\s*korea\b/i },
  { name: 'CECI', pattern: /\bceci\b/i },
  { name: '1ST LOOK', pattern: /1st\s*look/i },
  { name: 'CINE21', pattern: /cine\s*21/i },
  { name: 'MAXIM', pattern: /maxim/i },
  { name: 'SINGLES', pattern: /singles/i },
  { name: 'NYLON', pattern: /nylon/i },
  { name: 'THE STAR', pattern: /the\s*star/i },
  { name: 'STAR1', pattern: /star\s*1|star1/i },
  { name: "MEN'S HEALTH", pattern: /men'?s\s*health/i }
];

function 雑誌名を抽出_(data) {
  const explicitName = String(data && data.magazineName || '').trim();
  if (explicitName) return explicitName;

  const title = String(data && data.title || '').trim();
  const categoryName = String(data && data.categoryName || '').trim();
  const combined = `${title} ${categoryName}`;

  for (const entry of MAGAZINE_NAME_PATTERNS_) {
    if (entry.pattern.test(combined)) {
      return entry.name;
    }
  }

  const cleaned = title
    .replace(/\b20\d{2}[./-]\d{1,2}\b.*$/i, '')
    .replace(/\b\d{4}년\s*\d{1,2}월.*$/i, '')
    .replace(/\b\d{4}年\s*\d{1,2}月.*$/i, '')
    .replace(/\s+(KOREA|코리아)\b/ig, '')
    .replace(/[<＜【\[(].*$/, '')
    .replace(/\s+/g, ' ')
    .trim();

  return cleaned || title;
}

function 雑誌号情報を抽出_(data) {
  const title = String(data && data.title || '').trim();
  const basicInfo = String(data && data.basicInfo || '').trim();
  const description = String(data && data.description || '').trim();
  const magazineName = 雑誌名を抽出_(data);
  const combined = [title, basicInfo, description].filter(Boolean).join(' / ');

  let year = '';
  let month = '';
  let issue = '';

  const yearMonthPatterns = [
    /\b(20\d{2})[./-]\s*(\d{1,2})\b/i,
    /\b(20\d{2})年\s*(\d{1,2})月\b/i,
    /\b(20\d{2})년\s*(\d{1,2})월\b/i
  ];
  for (const pattern of yearMonthPatterns) {
    const match = combined.match(pattern);
    if (match) {
      year = match[1];
      month = String(Number(match[2]));
      break;
    }
  }

  let rest = title;
  if (magazineName) {
    const escapedMagazine = magazineName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+');
    rest = rest.replace(new RegExp(escapedMagazine, 'ig'), ' ');
  }
  rest = rest
    .replace(/\b(?:KOREA|TAIWAN|HONG\s*KONG|HONGKONG|CHINA|THAILAND|HK|TW|CN|KR|TH)\b/ig, ' ')
    .replace(/\b20\d{2}[./-]\d{1,2}\b/ig, ' ')
    .replace(/\b\d{4}年\s*\d{1,2}月\b/ig, ' ')
    .replace(/\b\d{4}년\s*\d{1,2}월\b/ig, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  const issuePatterns = [
    /\b(?:VOL\.?|NO\.?|ISSUE|#)\s*([1-9]\d{0,3})\b/i,
    /\b([1-9]\d{0,3})\s*호\b/i,
    /\b([1-9]\d{0,2})\b/
  ];
  for (const pattern of issuePatterns) {
    const match = rest.match(pattern);
    if (!match) continue;
    const candidate = String(match[1] || '').trim();
    if (!candidate) continue;
    if (year && candidate === year) continue;
    if (month && candidate === month) continue;
    if (Number(candidate) > 500) continue;
    issue = candidate;
    break;
  }

  return { year, month, issue };
}

function 韓国マンガカテゴリを補正_(source) {
  const categoryName = String(source && source.categoryName || '').trim();
  const title = String(source && source.title || '').trim();
  const basicInfo = String(source && source.basicInfo || '').trim();
  const description = String(source && source.description || '').trim();
  const mainText = [categoryName, title, basicInfo].join(' ').toLowerCase();
  const combined = [mainText, description].join(' ').toLowerCase();

  if (!combined) return '';
  if (/sticker|스티커/.test(combined)) return 'ステッカー';
  if (/seal|씰/.test(combined)) return 'シール';
  if (/dvd/.test(combined)) return 'DVD';
  if (/blu[-\s]?ray|블루레이/.test(combined)) return 'Blu-ray';
  if (/\blp\b|vinyl/.test(combined)) return 'LP';
  if (/\bcd\b|음반/.test(combined)) return 'CD';
  if (/scenario|시나리오/.test(combined)) return 'シナリオ集';
  if (/script|screenplay|대본/.test(combined)) return '台本';
  if (/picture\s*book|그림책|絵本/.test(combined)) return '絵本';
  if (/papercraft|paper\s*art|cut\s*out|切り絵|종이공예|종이접기/.test(combined)) return '切り絵';
  if (/handcraft|craft|자수|뜨개|수예|手芸/.test(combined)) return '手芸';
  if (/magazine|잡지|매거진/.test(combined)) return '雑誌';
  if (/essay|에세이/.test(combined)) return 'エッセイ';
  if (/comic|comics|만화|코믹|webtoon/.test(combined)) return 'まんが';
  if (/setting|guide\s*book|guidebook|fan\s*book|fanbook|character\s*book|official\s*guide|設定集|설정집|가이드북|팬북|캐릭터북|자료집/.test(combined)) return '設定集';
  if (/art\s*book|artbook|아트북|illustration|illust|画集|화보|원화|작화집|포토북|컨셉북/.test(combined)) return 'アートブック';
  if (/goods|gift|굿즈/.test(combined)) return 'グッズ';
  if (/novel|소설|라이트노벨|light\s*novel/.test(combined)) return '小説';
  return 'まんが';
}

function 韓国書籍カテゴリを補正_(source) {
  const categoryName = String(source && source.categoryName || '').trim();
  const title = String(source && source.title || '').trim();
  const basicInfo = String(source && source.basicInfo || '').trim();
  const description = String(source && source.description || '').trim();
  const mainText = [categoryName, title, basicInfo].join(' ').toLowerCase();
  const combined = [mainText, description].join(' ').toLowerCase();

  if (!combined) return '';
  if (/sticker|스티커/.test(combined)) return 'ステッカー';
  if (/seal|씰/.test(combined)) return 'シール';
  if (/dvd/.test(combined)) return 'DVD';
  if (/blu[-\s]?ray|블루레이/.test(combined)) return 'Blu-ray';
  if (/\blp\b|vinyl/.test(combined)) return 'LP';
  if (/\bcd\b|음반/.test(combined)) return 'CD';
  if (/o\.s\.t\.|\bost\b|original\s*sound\s*track|사운드트랙/.test(combined)) return 'OST';
  if (/scenario|시나리오/.test(combined)) return 'シナリオ集';
  if (/script|screenplay|대본/.test(combined)) return '台本';
  if (/picture\s*book|그림책|絵本/.test(combined)) return '絵本';
  if (/papercraft|paper\s*art|cut\s*out|切り絵|종이공예|종이접기/.test(combined)) return '切り絵';
  if (/handcraft|craft|자수|뜨개|수예|手芸/.test(combined)) return '手芸';
  if (/essay|에세이|산문/.test(mainText)) return 'エッセイ';
  if (/참고서|수험서|문제집|기출|모의고사|수능|내신|검정고시|자격증|공무원|고시|임용|편입|leet|meet|deet|psat|ncs|toeic|toefl|ielts|teps|jlpt|hsk|topik|参考書|問題集|過去問|受験|資格/.test(mainText)) return '参考書';
  if (/교재|학습지|학습서|워크북|work\s*book|workbook|text\s*book|textbook|student\s*book|activity\s*book|teacher'?s\s*book|course\s*book|教材|教科書/.test(mainText)) return '教材';
  if (/setting|guide\s*book|guidebook|fan\s*book|fanbook|character\s*book|official\s*guide|設定集|설정집|가이드북|팬북|캐릭터북|자료집/.test(combined)) return '設定集';
  if (/art\s*book|artbook|아트북|illustration|illust|画集|화보|원화|작화집|포토북|컨셉북|예체능|미술|디자인|사진/.test(combined)) return 'アートブック';
  if (/magazine|잡지|매거진/.test(combined)) return '雑誌';
  if (/goods|gift|굿즈/.test(combined)) return 'グッズ';
  if (/comic|comics|만화|코믹|webtoon/.test(combined)) return 'まんが';
  if (/novel|소설|라이트노벨|light\s*novel|문학|시집/.test(combined)) return '小説';
  return '';
}

function 韓国音楽映像カテゴリを補正_(source) {
  const categoryName = String(source && source.categoryName || '').trim();
  const title = String(source && source.title || '').trim();
  const basicInfo = String(source && source.basicInfo || '').trim();
  const description = String(source && source.description || '').trim();
  const titleAndCategory = [title, categoryName].join(' ').toLowerCase();
  const combined = [title, categoryName, basicInfo, description].join(' ').toLowerCase();

  if (!combined) return '';
  if (/\bdvd\b|\[dvd\]|dvd\//.test(titleAndCategory)) return 'DVD';
  if (/blu[-\s]?ray|blue\s*ray|블루레이/.test(titleAndCategory)) return 'Blu-ray';
  if (/\blp\b|vinyl|record|레코드/.test(titleAndCategory)) return 'LP';
  if (/o\.s\.t\.|\bost\b|original\s*sound\s*track|사운드트랙/.test(titleAndCategory)) return 'OST';
  if (/\bcd\b|음반|album|single|mini\s*album|ep\b/.test(titleAndCategory)) return 'CD';

  if (/\bdvd\b|\[dvd\]|dvd\//.test(combined)) return 'DVD';
  if (/blu[-\s]?ray|blue\s*ray|블루레이/.test(combined)) return 'Blu-ray';
  if (/\blp\b|vinyl|record|레코드/.test(combined)) return 'LP';
  if (/o\.s\.t\.|\bost\b|original\s*sound\s*track|사운드트랙/.test(combined)) return 'OST';
  if (/\bcd\b|음반|album|single|mini\s*album|ep\b/.test(combined)) return 'CD';
  return categoryName || '';
}

const KOREA_書籍系カテゴリ有効値 = ['まんが', '小説', 'エッセイ', 'グッズ', '設定集', 'アートブック', '雑誌', 'OST', 'CD', 'DVD', 'Blu-ray', 'LP', '絵本', 'シナリオ集', '台本', 'ステッカー', 'シール', '手芸', '切り絵', '教材', '参考書'];

const KOREA_カテゴリ有効値マップ = {
  '韓国書籍': KOREA_書籍系カテゴリ有効値,
  '韓国マンガ': KOREA_書籍系カテゴリ有効値,
  '韓国音楽映像': ['OST', 'CD', 'DVD', 'Blu-ray', 'LP'],
  '韓国雑誌': ['雑誌'],
  '韓国グッズ': ['グッズ']
};

// シートのカテゴリ列はプルダウン（入力規則）なので、有効値以外を書き込むとエラーになる。
// 分類できなかった値（アラジンの "국내도서>만화>BL만화" のような生パス）は空欄に落とす。
// 台湾側（拡張機能 popup.shared.js の getCategoryValue）と同じ方針。
function カテゴリ入力値を検証_(sheetName, value) {
  const text = String(value == null ? '' : value).trim();
  if (!text) return '';

  const validValues = KOREA_カテゴリ有効値マップ[sheetName];
  if (validValues) return validValues.indexOf(text) === -1 ? '' : text;

  // ホワイトリスト未定義のシートは、台湾側と同じ形の安全網だけ掛ける
  if (/[>＞]/.test(text) || text.length > 12) return '';
  return text;
}

function 判定済みカテゴリを取得_(sheetName, source) {
  const value = String(source && (source.sheetCategory || source.normalizedCategory || source.登録カテゴリ) || '').trim();
  if (!value) return '';
  return カテゴリ入力値を検証_(sheetName, value);
}

function カテゴリ入力値を補正_(sheetName, source) {
  const normalizedCategory = 判定済みカテゴリを取得_(sheetName, source);
  if (normalizedCategory) return normalizedCategory;

  let detected = '';
  if (sheetName === '韓国書籍') {
    detected = 韓国書籍カテゴリを補正_(source);
  } else if (sheetName === '韓国マンガ') {
    detected = 韓国マンガカテゴリを補正_(source);
  } else if (sheetName === '韓国音楽映像') {
    detected = 韓国音楽映像カテゴリを補正_(source);
  } else {
    detected = String(source && source.categoryName || '').trim();
  }

  return カテゴリ入力値を検証_(sheetName, detected);
}
function 保存先シート候補を取得_(genre) {
  const preferredSheets = ALADIN_GENRE_SHEET_MAP[String(genre || '').trim()] || [];
  const targetSheetNames = preferredSheets.length ? preferredSheets : ALADIN_ALL_TARGET_SHEETS;
  const spreadsheetCache = {};
  const targets = [];

  targetSheetNames.forEach(sheetName => {
    const target = 保存先シートを取得_(sheetName, spreadsheetCache);
    if (target) targets.push(target);
  });

  return targets;
}

function 保存先シートを取得_(sheetName, spreadsheetCache) {
  const destination = ALADIN_SPREADSHEET_DESTINATION_MAP[sheetName] || {};
  const spreadsheetId = スプレッドシートIDを抽出_(destination.spreadsheetIdOrUrl || '');
  if (!spreadsheetId) return null;

  let ss = spreadsheetCache[spreadsheetId];
  if (!ss) {
    try {
      ss = SpreadsheetApp.openById(spreadsheetId);
    } catch (error) {
      throw new Error(`保存先を開けませんでした: ${sheetName} (${destination.spreadsheetLabel || spreadsheetId}) / ${error.message}`);
    }
    spreadsheetCache[spreadsheetId] = ss;
  }

  const sh = ss.getSheetByName(sheetName);
  if (!sh) {
    throw new Error(`保存先シートが見つかりません: ${sheetName} (${destination.spreadsheetLabel || ss.getName()})`);
  }

  return {
    sheetName,
    sheet: sh,
    spreadsheetId,
    spreadsheetName: ss.getName()
  };
}

function 保存先未設定メッセージを組み立て_(genre) {
  const preferredSheets = ALADIN_GENRE_SHEET_MAP[String(genre || '').trim()] || [];
  const targetSheetNames = preferredSheets.length ? preferredSheets : ALADIN_ALL_TARGET_SHEETS;
  const missingLabels = targetSheetNames
    .filter(sheetName => {
      const destination = ALADIN_SPREADSHEET_DESTINATION_MAP[sheetName] || {};
      return !スプレッドシートIDを抽出_(destination.spreadsheetIdOrUrl || '');
    })
    .map(sheetName => {
      const destination = ALADIN_SPREADSHEET_DESTINATION_MAP[sheetName] || {};
      return `${sheetName} -> ${destination.spreadsheetLabel || '未設定'}`;
    });

  if (!missingLabels.length) {
    return '保存先シートが見つかりませんでした';
  }

  return `保存先スプレッドシート未設定: ${missingLabels.join(', ')}`;
}

function スプレッドシートIDを抽出_(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  const match = text.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : text;
}

function シートコンテキストを取得_(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const 列 = {};
  headers.forEach((header, index) => {
    if (header) 列[String(header).trim()] = index + 1;
  });

  const fallbackMap = {
    '韓国書籍': {
      url: ['リンク', 'アラジンURL', '購入URL'],
      isbn: 'ISBN',
      title: ['原題商品タイトル', '原題タイトル', '作品名（原題）', '商品名（原題）', '商品名(原題)'],
      author: ['作者', '著者'],
      publisher: '出版社',
      pubDate: '発売日',
      price: '原価',
      cover: ['メイン画像', 'メイン画像URL'],
      additionalImages: ['追加画像', '追加画像URL'],
      description: ['商品説明', '備考'],
      categoryName: ['カテゴリ', 'ジャンル分類'],
      itemId: ['サイト商品コード', 'アラジン商品コード']
    },
    '韓国マンガ': {
      url: ['リンク', 'アラジンURL', '購入URL'],
      isbn: 'ISBN',
      title: ['原題商品タイトル', '原題タイトル', '作品名（原題）', '商品名（原題）', '商品名(原題)'],
      author: ['作者', '著者'],
      publisher: '出版社',
      pubDate: '発売日',
      price: '原価',
      cover: ['メイン画像', 'メイン画像URL'],
      additionalImages: ['追加画像', '追加画像URL'],
      description: ['商品説明', '備考'],
      categoryName: ['カテゴリ', 'ジャンル分類'],
      itemId: ['サイト商品コード', 'アラジン商品コード']
    },
    '韓国音楽映像': {
      url: ['リンク', 'アラジンURL', '購入URL'],
      isbn: 'JANコード',
      title: ['原題商品タイトル', '原題タイトル', '作品名（原題）', '商品名（原題）', '商品名(原題)'],
      worksTitle: ['商品名(タイトル)', '商品名（タイトル）', 'Worksタイトル'],
      author: ['アーティスト名', 'アーティスト'],
      pubDate: '発売日',
      price: '原価',
      cover: ['メイン画像', 'メイン画像URL'],
      additionalImages: ['追加画像', '追加画像URL'],
      description: ['商品説明', '備考'],
      categoryName: ['カテゴリ', 'ジャンル分類'],
      itemId: ['サイト商品コード', 'アラジン商品コード']
    },
    '韓国グッズ': {
      url: ['リンク', '購入URL', 'アラジンURL'],
      title: ['原題商品タイトル', '原題タイトル', '作品名（原題）', '商品名（原題）', '商品名(原題)'],
      pubDate: '発売日',
      price: '原価',
      cover: ['メイン画像', 'メイン画像URL'],
      additionalImages: ['追加画像', '追加画像URL'],
      description: ['商品説明', '備考'],
      itemId: ['サイト商品コード', 'アラジン商品コード']
    },
    '韓国雑誌': {
      url: ['リンク', 'アラジンURL', '購入URL'],
      title: ['原題商品タイトル', '原題商品名', '原題タイトル', '商品名（原題）', '商品名(原題)'],
      magazineName: '雑誌名',
      pubDate: '発売日',
      price: '原価',
      cover: ['メイン画像', 'メイン画像URL'],
      additionalImages: ['追加画像', '追加画像URL'],
      description: ['商品説明', '備考'],
      itemId: ['サイト商品コード', 'アラジン商品コード']
    }
  };

  const columnMapSource = (typeof ALADIN_COLUMN_MAP !== 'undefined' && ALADIN_COLUMN_MAP[sh.getName()])
    ? ALADIN_COLUMN_MAP[sh.getName()]
    : fallbackMap[sh.getName()] || {};

  return {
    headers,
    列,
    colMap: columnMapSource,
    idCol: 最初に見つかった列番号を返す_(列, ['サイト商品コード', 'アラジン商品コード', 'ItemId']),
    isbnCol: 最初に見つかった列番号を返す_(列, ['ISBN', 'ISBN13', 'ISBNコード'])
  };
}

function 最初に見つかった列番号を返す_(列, headerNames) {
  const names = Array.isArray(headerNames) ? headerNames : [headerNames];
  for (const name of names) {
    if (name && 列[name]) return 列[name];
  }
  return null;
}

function 最初に見つかったヘッダー名を返す_(列, headerNames) {
  const names = Array.isArray(headerNames) ? headerNames : [headerNames];
  return names.find(name => name && 列[name]) || null;
}

function 追加対象行を決める_(sh, context) {
  const lastColumn = Math.max(sh.getLastColumn(), 1);
  const lastRow = Math.max(sh.getLastRow(), 1);
  if (lastRow < 2) return 2;

  const headers = context && Array.isArray(context.headers)
    ? context.headers.map(header => String(header || '').trim())
    : sh.getRange(1, 1, 1, lastColumn).getValues()[0].map(header => String(header || '').trim());
  const appendTargetIndexes = headers
    .map((header, index) => ALADIN_APPEND_VALUE_HEADERS.includes(header) ? index : -1)
    .filter(index => index >= 0);
  const indexes = appendTargetIndexes.length
    ? appendTargetIndexes
    : Array.from({ length: lastColumn }, (_, index) => index);

  const values = sh.getRange(2, 1, lastRow - 1, lastColumn).getDisplayValues();
  for (let i = values.length - 1; i >= 0; i--) {
    const hasVisibleValue = indexes.some(index => {
      const value = String(values[i][index] || '').trim();
      return value !== '' && value !== 'FALSE';
    });
    if (hasVisibleValue) {
      return i + 3;
    }
  }

  return 2;
}

function ItemId一致行を探す_(sh, idCol, itemId) {
  if (!idCol || sh.getLastRow() < 2) return null;

  const values = sh.getRange(2, idCol, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(itemId).trim()) {
      return i + 2;
    }
  }
  return null;
}

function ISBN値を正規化_(value) {
  return String(value || '').replace(/[^\dXx]/g, '').toUpperCase();
}

function ISBN一致行を探す_(sh, isbnCol, isbn) {
  const normalizedIsbn = ISBN値を正規化_(isbn);
  if (!isbnCol || !normalizedIsbn || sh.getLastRow() < 2) return null;

  const values = sh.getRange(2, isbnCol, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (ISBN値を正規化_(values[i][0]) === normalizedIsbn) {
      return i + 2;
    }
  }
  return null;
}

function 既存商品行を探す_(sh, context, data) {
  const itemId = String(data && data.itemId || '').trim();
  const itemRow = ItemId一致行を探す_(sh, context && context.idCol, itemId);
  if (itemRow) return itemRow;

  const isbn = data && (data.isbn13 || data.isbn);
  return ISBN一致行を探す_(sh, context && context.isbnCol, isbn);
}

function WEBAPP_URLリストへ分割_(value) {
  const source = Array.isArray(value) ? value.join(';') : String(value || '');
  const matches = source.match(/https?:\/\/[^\s;]+/gi) || [];
  const seen = {};
  const urls = [];

  matches.forEach(url => {
    const clean = String(url || '').trim();
    if (!clean || seen[clean]) return;
    seen[clean] = true;
    urls.push(clean);
  });

  return urls;
}

function WEBAPP_セルへリンク値を書き込む_(range, value) {
  const urls = WEBAPP_URLリストへ分割_(value);
  if (!urls.length) {
    range.setValue(value);
    return;
  }

  const text = urls.join('\n');
  const builder = SpreadsheetApp.newRichTextValue().setText(text);
  let offset = 0;

  urls.forEach(url => {
    builder.setLinkUrl(offset, offset + url.length, url);
    offset += url.length + 1;
  });

  range.setRichTextValue(builder.build());
}

function WEBAPP_アラジンWorksタイトルを推定_(data, shName) {
  const originalTitle = String(data && (data.title || data.productTitle || data.worksTitle) || '').trim();
  if (!originalTitle) return '';

  let title = originalTitle
    .replace(/\u3000/g, ' ')
    .replace(/[（）]/g, match => match === '（' ? '(' : ')')
    .replace(/[［］]/g, match => match === '［' ? '[' : ']')
    .replace(/[：]/g, ':')
    .replace(/[‐−–—―]/g, '-')
    .replace(/\bo\.?\s*s\.?\s*t\.?\b/gi, 'OST')
    .replace(/\s+/g, ' ')
    .trim();

  const categoryText = String(
    (data && (data.sheetCategory || data.normalizedCategory || data.categoryName || data.登録カテゴリ)) || ''
  ).trim();
  const isMedia = shName === '韓国音楽映像'
    || data && data.genre === '音楽映像'
    || /^(CD|LP|OST|DVD|Blu-ray)$/i.test(categoryText)
    || /音楽|映像|음반|dvd|blu[-\s]?ray|블루레이/i.test(categoryText);

  if (!isMedia) return title;

  title = title.replace(/^\s*\[[^\]]*(?:4k|blu[-\s]?ray|블루레이|dvd|uhd|ubd)[^\]]*\]\s*/i, '');

  const dashMatch = title.match(/^(.{1,45}?)\s+-\s+(.+)$/);
  if (dashMatch) {
    const beforeDash = dashMatch[1].trim();
    const afterDash = dashMatch[2].trim();
    const afterStartsWithAlbumTitle =
      /^(?:정규|미니|싱글|스페셜|리패키지)\s*\d+\s*(?:집|앨범)\b/i.test(afterDash)
      || /^\d+(?:st|nd|rd|th)?\s+(?:full|mini)\s+album\b/i.test(afterDash)
      || /(?:no tragedy|revive)/i.test(afterDash);
    const afterLooksLikeOnlyPackage =
      /(?:상자|박스|부클릿|북클릿|소책자|컵받침대|거치대|종이\s*거치대|파우치|포토카드|포토북|카드|커버|세트|랜덤|booklet|photocard|poster|pouch|nfc|\d+\s*(?:종|p|disc)|cd\s*\()/i.test(afterDash);

    if (afterStartsWithAlbumTitle) {
      title = afterDash;
    } else if (beforeDash && afterLooksLikeOnlyPackage) {
      title = beforeDash;
    }
  }
  title = title
    .replace(/^\s*(?:정규|미니|싱글|스페셜|리패키지)\s*\d+\s*(?:집|앨범)\s*/i, '')
    .replace(/^\s*\d+(?:st|nd|rd|th)?\s+(?:full|mini)\s+album\s*/i, '')
    .replace(/\[[^\]]*(?:180g|white\s*marble|vinyl|lp|ver\.?|한정|限定|picture|픽처|photocard|poster|cd|dvd|blu[-\s]?ray|ubd|uhd|disc)[^\]]*\]/gi, ' ')
    .replace(/\([^)]*(?:album\s*ver\.?|mubeat|pouch|파우치|mini\s*cd|미니\s*cd|photocard|포토카드|포토북|booklet|소책자|nfc|disc|cd|dvd|blu[-\s]?ray|ubd|uhd|\d+\s*종|\d+\s*p)[^)]*\)/gi, ' ')
    .replace(/\s+-\s+.*(?:커버|바이닐|포토|파우치|소책자|부클릿|북클릿|상자|박스|컵받침대|거치대|종이\s*거치대|랜덤|booklet|poster|photocard|album\s*ver\.?|nfc|카드|세트|컵).*$/i, ' ')
    .replace(/\s*:\s*(?:스틸북|풀슬립|full\s*slip|steelbook).*$/i, ' ')
    .replace(/\b(?:4k\s*uhd|4k|uhd|ubd|2d|blu[-\s]?ray|dvd|cd|lp|vinyl|record)\b/gi, ' ')
    .replace(/(?:블루레이|스틸북|풀슬립|소책자|포토카드|포토북|북클릿|파우치|한정반|한정판|게이트폴드\s*커버|바이닐)/gi, ' ')
    .replace(/\b(?:a|b|c|d)\s*ver\.?\b/gi, ' ')
    .replace(/\b(?:album|mubeat\s*album)\s*ver\.?\b/gi, ' ')
    .replace(/\d+\s*(?:disc|종|p)\b/gi, ' ')
    .replace(/[<＞>]+/g, ' ')
    .replace(/[『』「」"'`]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  title = title
    .replace(/^(?:넷플릭스\s*)?시리즈\s+(.+?)\s+OST$/i, '$1 OST')
    .replace(/\bost\b/gi, 'OST');
  return title || originalTitle;
}

function 日本語タイトル照会列を確保_(sh) {
  const lastColumn = Math.max(sh.getLastColumn(), 1);
  const headers = sh.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  const normalized = {};
  headers.forEach(function(header) {
    const key = String(header || '').replace(/\s+/g, '');
    if (key) normalized[key] = true;
  });

  const missing = ALADIN_LOOKUP_WRITEBACK_HEADERS.filter(function(header) {
    return !normalized[String(header || '').replace(/\s+/g, '')];
  });
  if (!missing.length) return;

  sh.getRange(1, lastColumn + 1, 1, missing.length).setValues([missing]);
}

function 日本語タイトル照会ステータス表示_(lookup) {
  const result = normalizeLookupResult_(lookup);
  if (!result) return '未照会';
  if (result.status === 'resolved') return 'resolved';
  if (result.status === 'not_found') return '登録なし(全サイト)';
  if (result.status === 'partial_error') return '一部照会失敗';
  if (result.status === 'failed') return '照会失敗(' + (result.provider || '不明') + ')';
  if (result.status === 'skipped') return '未照会';
  return result.status || '未照会';
}

function 日本語タイトル照会結果を書き戻す_(sh, context, row, lookup, titleAnalysis) {
  const result = normalizeLookupResult_(lookup);
  if (!result) return;
  const headers = context && context.列 ? context.列 : {};
  const set = function(headerNames, value) {
    const headerName = 最初に見つかったヘッダー名を返す_(headers, headerNames);
    if (!headerName) return;
    sh.getRange(row, headers[headerName]).setValue(value === null || value === undefined ? '' : value);
  };
  const setIfPresent = function(headerNames, value) {
    if (value === null || value === undefined || value === '') return;
    set(headerNames, value);
  };

  if (result.status === 'resolved' && result.japaneseTitle) {
    set(['日本語タイトル', '商品名(日本語)', '商品名（日本語）'], result.japaneseTitle);
    set(['作品名（日本語）', '作品名日本語'], result.japaneseTitle);
  }
  set(['日本語タイトル照会結果'], 日本語タイトル照会ステータス表示_(result));
  set(['日本語タイトル照会元'], result.provider || '');
  set(['日本語タイトル照会ログ'], [result.trace || '', (result.errors || []).join(' / ')].filter(Boolean).join(' / '));
  setIfPresent(['検索用正規化タイトル'], result.normalizedSearchTitle || (titleAnalysis && titleAnalysis.normalizedSearchTitle) || '');
}

function アラジン行へ書き込む_(sh, context, row, data) {
  const { 列, colMap } = context;
  const setByHeader = (headerNames, value, options = {}) => {
    const headerName = 最初に見つかったヘッダー名を返す_(列, headerNames);
    if (!headerName) return;
    if (value === null || value === undefined || value === '') return;
    const range = sh.getRange(row, 列[headerName]);
    if (options.link) {
      WEBAPP_セルへリンク値を書き込む_(range, value);
      return;
    }
    range.setValue(value);
  };
  const setIfBlank = (headerNames, value) => {
    const headerName = 最初に見つかったヘッダー名を返す_(列, headerNames);
    if (!headerName) return;
    if (value === null || value === undefined || value === '') return;
    const range = sh.getRange(row, 列[headerName]);
    if (String(range.getDisplayValue() || '').trim()) return;
    range.setValue(value);
  };

  const basicInfoCol = 列['基本情報'] || null;
  const mergedDescription = 商品説明テキストを組み立てる_(data);

  setByHeader(colMap.url, data.pageUrl || '', { link: true });
  setByHeader(colMap.isbn, data.isbn13 || '');
  setByHeader(colMap.title, data.title || '');
  const worksTitle = data.worksTitle || data.productTitle || WEBAPP_アラジンWorksタイトルを推定_(data, sh.getName());
  setByHeader(colMap.worksTitle || ['商品名(タイトル)', '商品名（タイトル）', 'Worksタイトル'], worksTitle);
  setByHeader(['商品名(タイトル)', '商品名（タイトル）', 'Worksタイトル'], worksTitle);
  if (/^韓国/.test(sh.getName())) {
    setByHeader(colMap.language || ['言語'], data.language || data.言語 || '韓国');
  }
  setByHeader(colMap.author, data.author || '');
  setByHeader(colMap.publisher, data.publisher || '');
  setByHeader(colMap.pubDate, data.pubDate || '');
  setByHeader(colMap.price, data.priceSales || '');
  setByHeader(colMap.cover, data.cover || '', { link: true });
  setByHeader(colMap.additionalImages, data.additionalImages || '', { link: true });
  setByHeader(colMap.description, mergedDescription);
  setByHeader(colMap.categoryName, カテゴリ入力値を補正_(sh.getName(), data));
  setByHeader(colMap.itemId, String(data.itemId || ''));

  if (sh.getName() === '韓国雑誌') {
    const magazineName = 雑誌名を抽出_(data);
    const issueInfo = 雑誌号情報を抽出_(data);
    setByHeader(colMap.magazineName || '雑誌名', magazineName);
    setByHeader('原題タイトル', magazineName || data.title || '');
    setByHeader('原題商品名', data.title || '');
    setIfBlank('年', issueInfo.year);
    setIfBlank('月', issueInfo.month);
    setIfBlank('号数', issueInfo.issue);
  }

  if (basicInfoCol && data.basicInfo) {
    const existing = sh.getRange(row, basicInfoCol).getValue();
    if (!existing) {
      sh.getRange(row, basicInfoCol).setValue(String(data.basicInfo).slice(0, 500));
    }
  }

  日本語タイトル照会結果を書き戻す_(sh, context, row, data.japaneseTitleLookup, data.titleAnalysis || {});
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














