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

function buildBookSheetDescription(product) {
  const lines = cleanDescription(product?.商品説明 || '')
    .split('\n')
    .map(line => line.trim())
    .filter(Boolean);
  const bonusLines = cleanDescription(product?.特典情報 || product?.補足項目 || '')
    .split('\n')
    .map(line => line.trim())
    .filter(Boolean);

  const metadataLines = [
    trimValue(product?.著者 || product?.作者 || '') ? '作者：' + trimValue(product?.著者 || product?.作者 || '') : '',
    trimValue(product?.翻訳者 || product?.譯者 || '') ? '譯者：' + trimValue(product?.翻訳者 || product?.譯者 || '') : '',
    trimValue(product?.イラストレーター || product?.插畫 || product?.插画 || '') ? '插畫：' + trimValue(product?.イラストレーター || product?.插畫 || product?.插画 || '') : '',
    trimValue(product?.出版社 || '') ? '出版社：' + trimValue(product?.出版社 || '') : '',
    trimValue(product?.発売日 || product?.出版日期 || '') ? '出版日期：' + trimValue(product?.発売日 || product?.出版日期 || '') : '',
    trimValue(product?.規格サイズ || product?.規格 || '') ? '規格：' + trimValue(product?.規格サイズ || product?.規格 || '') : '',
  ].filter(Boolean);

  for (const line of metadataLines) {
    if (!lines.includes(line)) lines.push(line);
  }
  for (const line of bonusLines) {
    if (!lines.includes(line)) lines.push(line);
  }

  return lines.join('\n');
}

/**
 * タイトル文字列から巻数番号を抽出する（堅牢版）。
 * core/titleAnalysis.js の extractVolume は正規表現末尾の \b（CJK文字に効かない）が原因で
 * 「第02巻」「第04巻」のような日本語構成タイトルから巻数を取りこぼすため、ここで専用実装する。
 * 優先順位: 第N巻/集等のマーカー > vol./括弧囲み > 版種・末尾の数字。
 */
function extractTaiwanVolumeNumber_(text) {
  const s = String(text || '').replace(/\s+/g, ' ').trim();
  if (!s) return '';
  const groups = [
    // 強: 第N巻 / N巻 / N集 / N冊 / N話 / N期 / N号 等（\b を使わない）
    [/(?:第\s*)?(\d{1,3})\s*(?:巻|卷|集|冊|册|話|期|号|號)/u],
    // 中: vol.N / 末尾の (N) / 】N 等
    [/(?:vol(?:ume)?\.?\s*)(\d{1,3})/iu, /[)\]】》」』]\s*(\d{1,3})\s*$/u, /\((\d{1,3})\)\s*$/u],
    // 緩: 版種キーワード直前の数字 / 末尾の数字（タイトル途中の数字は拾わない）
    [/(?:^|[^\d])(\d{1,3})\s*(?:首刷|初版|初回|限定|特[裝装]|通常|完)/u, /(?:^|[^\d])(\d{1,3})\s*$/u],
  ];
  for (const group of groups) {
    for (const p of group) {
      const m = s.match(p);
      if (m && m[1]) {
        const n = parseInt(m[1], 10);
        if (Number.isFinite(n) && n > 0) return String(n);
      }
    }
  }
  return '';
}

/**
 * タイトル文字列から巻数列（単巻数 / セット巻数開始番号 / セット巻数終了番号）を導出する。
 * ルール:
 *   1) 上下巻（「上+下」「上下巻」等）→ セット扱い：開始=1 / 終了=2
 *   2) 巻数の数字が取れる（第N巻 / N巻 / N集 / 末尾N 等）→ 単巻数=N
 *   3) どちらも無い → 単巻数=1（デフォルト）
 * すでに明示値が入っている場合はそれを尊重する（手動指定・既存値を壊さない）。
 */
function deriveTaiwanVolumeColumns_(product) {
  const existingSingle = trimValue(product?.単巻数 || '');
  const existingStart = trimValue(product?.セット巻数開始番号 || '');
  const existingEnd = trimValue(product?.セット巻数終了番号 || '');
  if (existingSingle || existingStart || existingEnd) {
    return {
      単巻数: existingSingle,
      セット巻数開始番号: existingStart,
      セット巻数終了番号: existingEnd,
    };
  }

  // 巻数番号は「第N巻」を含む構成タイトルを優先（中華原題の途中数字の誤検出を避ける）。
  const texts = [
    product?.タイトル,
    product?.商品名,
    product?.原題商品タイトル,
    product?.原題タイトル,
  ].map((v) => trimValue(v || '')).filter(Boolean);
  const joined = texts.join('  ');

  // 1) 上下巻（セット）判定 → 開始=1 / 終了=2
  const isUpperLowerSet =
    /上\s*[+＋&＆・,，、/／と]\s*下/u.test(joined) ||
    /上下\s*(?:巻|卷|集|冊|册|合|套|兩|两|販|贩)/u.test(joined) ||
    /[【（(「\[]\s*上\s*[+＋]?\s*下/u.test(joined);
  if (isUpperLowerSet) {
    return { 単巻数: '', セット巻数開始番号: '1', セット巻数終了番号: '2' };
  }

  // 2) 巻数番号を本文から抽出 → 単巻数（マーカー優先で全テキストを走査）
  let vol = '';
  for (const t of texts) {
    vol = extractTaiwanVolumeNumber_(t);
    if (vol) break;
  }
  if (vol) {
    return { 単巻数: vol, セット巻数開始番号: '', セット巻数終了番号: '' };
  }

  // 3) 何も取れなければデフォルト1巻
  return { 単巻数: '1', セット巻数開始番号: '', セット巻数終了番号: '' };
}

function buildCommonBookSheetRow(product, overrides = {}) {
  const code = extractProductCode(product);
  const description = buildBookSheetDescription(product);
  const memo = buildBookMemo(product);
  const additional = getAdditionalImagesValue(product);
  const rawTitle = trimValue((overrides.rawTitle ?? product?.商品名) || '');
  const originalTitle = finalizeSheetOriginalWorkTitleRow((overrides.originalTitle ?? product?.原題タイトル) || rawTitle);
  const category = trimValue(overrides.category || getBookGenreLabel(product));
  const 巻数列 = deriveTaiwanVolumeColumns_(product);

  return {
    '発番発行': '',
    '登録状況': '',
    '商品コード（SKU）': trimValue(product?.SKU || ''),
    'サイト商品コード': code,
    'タイトル': trimValue(product?.タイトル || ''),
    '作者': trimValue(product?.著者 || ''),
    '日本語タイトル': normalizeSheetJapaneseWorkTitleForRow(getDisplayJapaneseTitleValue(product?.日本語タイトル)),
    'リンク': trimValue(product?.URL || ''),
    '原題タイトル': originalTitle,
    '原題商品タイトル': rawTitle,
    '売価': trimValue(product?.売価 || ''),
    '原価': trimValue(product?.価格 || ''),
    '粗利益率': trimValue(product?.粗利益率 || ''),
    'ISBN': /^\d{13}$/.test(trimValue(product?.ISBN || '')) ? trimValue(product?.ISBN || '') : '',
    '発売日': trimValue(product?.発売日 || ''),
    '言語': getLanguageValue(product),
    '単巻数': 巻数列.単巻数,
    'セット巻数開始番号': 巻数列.セット巻数開始番号,
    'セット巻数終了番号': 巻数列.セット巻数終了番号,
    'カテゴリ': category,
    '形態（通常/初回限定/特装）': detectEditionType(product),
    '配送パターン': trimValue(product?.配送パターン || ''),
    '特典メモ': '',
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


