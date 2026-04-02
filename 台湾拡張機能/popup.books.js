// popup.books.js - 台湾 books.com.tw 拡張 書籍共通 / まんが・非まんが振り分け

function isBookProduct(product) {
  const code = extractProductCode(product);
  return !!code && !/^[NM]/i.test(code);
}

function normalizeBookGenreLabel(value) {
  const text = trimValue(value);
  if (!text) return '書籍';

  const map = {
    'まんが': 'まんが',
    '小説': '小説',
    '設定集': '設定集',
    'アートブック': 'アートブック',
    'エッセイ': 'エッセイ',
    '絵本': '絵本',
    'シナリオ集': 'シナリオ集',
    '台本': '台本',
    'CD': 'CD',
    'DVD': 'DVD',
    'Blu-ray': 'Blu-ray',
    'LP': 'LP',
  };

  return map[text] || text || '書籍';
}

function getBookGenreLabel(product) {
  return normalizeBookGenreLabel(getCategoryValue(product));
}

function isComicBookProduct(product) {
  return isBookProduct(product) && getBookGenreLabel(product) === 'まんが';
}

function isNonMangaBookProduct(product) {
  return isBookProduct(product) && !isMagazineProduct(product) && !isComicBookProduct(product);
}

function buildBookMemo(product) {
  const sections = [
    trimValue(product?.翻訳者 || '') ? '翻訳者：' + trimValue(product?.翻訳者 || '') : '',
    trimValue(product?.イラストレーター || '') ? 'イラストレーター：' + trimValue(product?.イラストレーター || '') : '',
    trimValue(product?.出版社 || '') ? '出版社：' + trimValue(product?.出版社 || '') : '',
    trimValue(product?.規格サイズ || '') ? '規格：' + trimValue(product?.規格サイズ || '') : '',
    cleanDescription(product?.商品説明 || ''),
  ];
  return sections.filter(Boolean).join('\n');
}

function buildBookDescription(product) {
  const meta = [
    trimValue(product?.著者 || '') ? '作者：' + trimValue(product?.著者 || '') : '',
    trimValue(product?.翻訳者 || '') ? '翻訳者：' + trimValue(product?.翻訳者 || '') : '',
    trimValue(product?.イラストレーター || '') ? 'イラストレーター：' + trimValue(product?.イラストレーター || '') : '',
    trimValue(product?.出版社 || '') ? '出版社：' + trimValue(product?.出版社 || '') : '',
    trimValue(product?.発売日 || '') ? '出版日期：' + trimValue(product?.発売日 || '') : '',
    trimValue(product?.規格 || '') ? '規格：' + trimValue(product?.規格 || '')
      : trimValue(product?.規格サイズ || '') ? '規格：' + trimValue(product?.規格サイズ || '') : '',
  ].filter(Boolean);
  const body = cleanDescription(product?.商品説明 || '');
  return meta.length > 0 && body ? meta.join('\n') + '\n\n' + body
    : meta.length > 0 ? meta.join('\n')
    : body;
}

function buildCommonBookSheetRow(product, overrides = {}) {
  const code = extractProductCode(product);
  const description = buildBookDescription(product);
  const memo = buildBookMemo(product);
  const additional = getAdditionalImagesValue(product);
  const rawTitle = trimValue((overrides.rawTitle ?? product?.商品名) || '');
  const originalTitle = trimValue((overrides.originalTitle ?? product?.原題タイトル) || extractOriginalTitleText(rawTitle));
  const category = trimValue(overrides.category || getBookGenreLabel(product));

  return {
    '発番発行': '',
    '登録状況': '',
    '商品コード（SKU）': trimValue(product?.SKU || ''),
    'サイト商品コード': code,
    'タイトル': trimValue(product?.タイトル || ''),
    '作者': trimValue(product?.著者 || ''),
    '日本語タイトル': trimValue(product?.日本語タイトル || ''),
    'リンク': trimValue(product?.URL || ''),
    '原題タイトル': originalTitle,
    '原題商品タイトル': rawTitle,
    '売価': trimValue(product?.売価 || ''),
    '原価': trimValue(product?.価格 || ''),
    '粗利益率': trimValue(product?.粗利益率 || ''),
    'ISBN': trimValue(product?.ISBN || ''),
    '発売日': trimValue(product?.発売日 || ''),
    '言語': getLanguageValue(product),
    '単巻数': trimValue(product?.単巻数 || ''),
    'セット巻数開始番号': trimValue(product?.セット巻数開始番号 || ''),
    'セット巻数終了番号': trimValue(product?.セット巻数終了番号 || ''),
    'カテゴリ': category,
    '形態（通常/初回限定/特装）': detectEditionType(product),
    '配送パターン': trimValue(product?.配送パターン || ''),
    '特典メモ': extractBonusMemo(product),
    '商品説明': description,
    ' メモ': memo,
    'メイン画像': trimValue(product?.画像URL || ''),
    '追加画像': additional,
    '予約開始日': trimValue(product?.予約開始日 || ''),
    '予約終了日': trimValue(product?.予約終了日 || ''),
    '入荷予定日': trimValue(product?.入荷予定日 || ''),
    '発売日メモ（延期など）': trimValue(product?.発売日メモ || ''),
    '作品ID(W)（自動）': '',
    'SKU（自動）': '',
    'ステータス（自動）': '',
    '残日数（自動）': '',
    'アラート（自動）': '',
  };
}

function buildCommonBookCsvRow(product) {
  return [
    ...buildCsvBaseRow(product),
    product.ISBN || '',
    product.著者 || '',
    product.翻訳者 || '',
    product.イラストレーター || '',
    product.出版社 || '',
  ];
}

function buildBookSheetRow(product) {
  if (isComicBookProduct(product) && typeof buildComicBookSheetRow === 'function') {
    return buildComicBookSheetRow(product);
  }
  if (isNonMangaBookProduct(product) && typeof buildNonMangaBookSheetRow === 'function') {
    return buildNonMangaBookSheetRow(product);
  }
  return buildCommonBookSheetRow(product);
}

function buildBookCsvRow(product) {
  if (isComicBookProduct(product) && typeof buildComicBookCsvRow === 'function') {
    return buildComicBookCsvRow(product);
  }
  if (isNonMangaBookProduct(product) && typeof buildNonMangaBookCsvRow === 'function') {
    return buildNonMangaBookCsvRow(product);
  }
  return buildCommonBookCsvRow(product);
}

