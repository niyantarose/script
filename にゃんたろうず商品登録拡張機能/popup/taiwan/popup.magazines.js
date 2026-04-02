function resolveTodayString() {
  if (typeof formatToday === 'function') return formatToday();
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  return `${year}/${month}/${day}`;
}
// popup.magazines.js - 台湾 books.com.tw 拡張 雑誌処理

function extractYearMonthParts(dateText) {
  const text = trimValue(dateText);
  if (!text) return { year: '', month: '' };

  const match = text.match(/(\d{4})[\/-年](\d{1,2})/);
  if (!match) return { year: '', month: '' };

  return {
    year: match[1],
    month: String(parseInt(match[2], 10) || match[2]),
  };
}

function extractMagazineTitleText(value) {
  const fallback = trimValue(value);
  if (!fallback) return '';

  const normalized = extractOriginalTitleText(fallback).replace(/\s+/g, ' ').trim();
  const patterns = [
    /^([A-Za-z][A-Za-z0-9 '&+./-]*?)(?=\s*\d{1,2}\s*月(?:號|号)?(?:\s*\/\s*\d{4})?)/u,
    /^([A-Za-z][A-Za-z0-9 '&+./-]*?)(?=\s*(?:第\s*\d+\s*(?:期|號|号)|Vol\.?\s*\d+|No\.?\s*\d+|Issue\s*\d+))/iu,
    /^([A-Za-z][A-Za-z0-9 '&+./-]*?)(?=[㐀-鿿])/u,
    /^([A-Za-z][A-Za-z0-9 '&+./-]{0,80})/u,
  ];

  for (const pattern of patterns) {
    const match = normalized.match(pattern);
    if (match?.[1]) return normalizeMagazineBrandName(match[1]) || fallback;
  }

  return normalizeMagazineBrandName(normalized) || fallback;
}

function normalizeMagazineBrandName(value) {
  const text = trimValue(value)
    .replace(/\b(?:KOREA|TAIWAN|HONG\s*KONG|HONGKONG|CHINA|THAILAND|HK|TW|CN|KR|TH)\b/gi, ' ')
    .replace(/(?:韓国版|韓國版|台湾版|臺灣版|中国版|中國版|香港版|タイ版)/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!text) return '';

  const key = text
    .normalize('NFKC')
    .toUpperCase()
    .replace(/&/g, ' AND ')
    .replace(/[^A-Z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  const map = {
    'MARIE CLAIRE': 'Marie Claire',
    'VOGUE': 'VOGUE',
    'BAZAAR': 'BAZAAR',
    'GQ': 'GQ',
    'ELLE': 'ELLE',
    'ESQUIRE': 'Esquire',
    'ALLURE': 'allure',
    'MENS HEALTH': "Men's Health",
    'MEN S HEALTH': "Men's Health",
    'L OFFICIEL': "L'OFFICIEL",
    'L OFFICIEL HOMMES': "L'OFFICIEL HOMMES",
    'THE BIG ISSUE': 'THE BIG ISSUE',
    'W': 'W',
  };
  return map[key] || text;
}

function isMagazineProduct(product) {
  if (!isBookProduct(product)) return false;

  const source = [
    trimValue(product?.カテゴリ || ''),
    trimValue(product?.ページタイトル || ''),
    trimValue(product?.商品名 || ''),
    trimValue(product?.原題タイトル || ''),
    trimValue(product?.商品説明 || ''),
  ].filter(Boolean).join('\n');

  if (/雜誌|杂志|雑誌/u.test(source) || getCategoryValue(product) === '雑誌') return true;

  const rawTitle = trimValue(product?.商品名 || product?.ページタイトル || '');
  if (!rawTitle || !hasMagazineIssueMarker(rawTitle)) return false;

  const brandName = extractMagazineTitleText(rawTitle);
  return !!brandName && brandName !== rawTitle;
}

function buildMagazineSheetRow(product) {
  const code = extractProductCode(product);
  const description = cleanDescription(product?.商品説明 || '');
  const additional = getAdditionalImagesValue(product);
  const rawTitle = trimValue(product?.商品名 || '');
  const originalTitle = trimValue(product?.原題タイトル || extractMagazineTitleText(rawTitle));
  const issueInfo = extractMagazineIssueParts(rawTitle, product?.発売日 || '');

  return {
    '発番発行': '',
    '登録状況': '',
    '言語': getLanguageCodeValue(product),
    '雑誌名': '',
    '年': issueInfo.year,
    '月': issueInfo.month,
    '号数': issueInfo.issue,
    '表紙情報': rawTitle,
    '特典メモ': '',
    '親コード': trimValue(product?.親コード || ''),
    '商品名（出品用）': '',
    '粗利益率': trimValue(product?.粗利益率 || ''),
    '登録日': resolveTodayString(),
    '売価': trimValue(product?.売価 || ''),
    '配送パターン': trimValue(product?.配送パターン || ''),
    '登録者': trimValue(product?.登録者 || ''),
    '商品説明': description,
    '原価': trimValue(product?.価格 || ''),
    '原題タイトル': originalTitle,
    '原題商品名': rawTitle,
    '博客來商品コード': code,
    '博客來URL': trimValue(product?.URL || ''),
    'メイン画像URL': trimValue(product?.画像URL || ''),
    '追加画像URL': additional,
  };
}

function hasMagazineIssueMarker(value) {
  const text = trimValue(value);
  if (!text) return false;
  return /(?:\d{1,2}(?:\s*[./\-・~～]\s*\d{1,2})?\s*月(?:號|号)?(?:\s*\/\s*20\d{2})?|20\d{2}\s*(?:年|\/|-|\.)\s*\d{1,2}(?:\s*[./\-・~～]\s*\d{1,2})?\s*月?(?:號|号)?|第\s*\d+\s*(?:期|號|号)|VOL\.?\s*\d+|NO\.?\s*\d+|ISSUE\s*\d+)/i.test(text);
}

function extractMagazineIssueParts(rawTitle, releaseDateText) {
  const text = trimValue(rawTitle);
  let year = '';
  let month = '';
  let issue = '';

  let match = text.match(/(\d{1,2}(?:\s*[./\-・~～]\s*\d{1,2})?)\s*月(?:號|号)?\s*\/\s*(20\d{2})/i);
  if (match) {
    month = match[1].replace(/\s*/g, '');
    year = match[2];
  }

  if (!year || !month) {
    match = text.match(/(20\d{2})\s*(?:年|\/|-|\.)\s*(\d{1,2}(?:\s*[./\-・~～]\s*\d{1,2})?)\s*月?(?:號|号)?/i);
    if (match) {
      year = match[1];
      month = match[2].replace(/\s*/g, '');
    }
  }

  if (!year || !month) {
    const release = extractYearMonthParts(releaseDateText);
    year = year || release.year;
    month = month || release.month;
  }

  match = text.match(/第\s*([0-9]{1,5})\s*(?:期|號|号)/i);
  if (match) issue = match[1];
  if (!issue) {
    match = text.match(/\bVOL\.?\s*([0-9]{1,5})/i);
    if (match) issue = match[1];
  }

  return { year, month, issue };
}

function getProductSheetType(product) {
  const kind = String(product?.種別 || '').trim();
  if (/^書籍$/u.test(kind)) return isMagazineProduct(product) ? 'magazine' : 'book';
  if (/^グッズ$/u.test(kind)) return 'goods';
  if (!isBookProduct(product)) return 'goods';
  return isMagazineProduct(product) ? 'magazine' : 'book';
}

function getProductKind(product) {
  return getProductSheetType(product);
}

function getProductKindLabel(product) {
  switch (getProductSheetType(product)) {
    case 'magazine':
      return '雑誌';
    case 'goods':
      return 'グッズ';
    case 'book':
    default:
      return getBookGenreLabel(product);
  }
}

function getProductBadgeClass(product) {
  switch (getProductSheetType(product)) {
    case 'magazine':
      return 'badge-magazine';
    case 'goods':
      return 'badge-goods';
    case 'book':
    default:
      return isComicBookProduct(product) ? 'badge-comic' : 'badge-book-other';
  }
}

function buildMagazineCsvRow(product) {
  const rawTitle = trimValue(product?.商品名 || '');
  return [
    ...buildCsvBaseRow(product),
    trimValue(product?.原題タイトル || extractMagazineTitleText(rawTitle)),
    rawTitle,
  ];
}


