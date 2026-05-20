(function attachTitleAnalysis(global) {
  'use strict';

  const MAGAZINE_PATTERNS = [
    { name: 'DAZED & CONFUSED', pattern: /dazed\s*&\s*confused/i },
    { name: "HARPER'S BAZAAR", pattern: /harper'?s\s*bazaar/i },
    { name: 'MARIE CLAIRE', pattern: /marie\s*claire/i },
    { name: 'COSMOPOLITAN', pattern: /cosmopolitan/i },
    { name: 'W KOREA', pattern: /\bw\s*korea\b/i },
    { name: '1ST LOOK', pattern: /1st\s*look/i },
    { name: 'CINE21', pattern: /cine\s*21/i },
    { name: "MEN'S HEALTH", pattern: /men'?s\s*health/i },
    { name: 'BAZAAR', pattern: /\bbazaar\b/i },
    { name: 'ESQUIRE', pattern: /esquire/i },
    { name: 'ALLURE', pattern: /allure/i },
    { name: 'VOGUE', pattern: /vogue/i },
    { name: 'ELLE', pattern: /elle/i },
    { name: 'GQ', pattern: /\bgq\b/i },
    { name: 'ARENA', pattern: /arena/i },
    { name: 'CECI', pattern: /\bceci\b/i },
    { name: 'MAXIM', pattern: /maxim/i },
    { name: 'SINGLES', pattern: /singles/i },
    { name: 'NYLON', pattern: /nylon/i },
    { name: 'THE STAR', pattern: /the\s*star/i },
    { name: 'STAR1', pattern: /star\s*1|star1/i }
  ];

  const PROVIDER_ORDER = {
    manga: ['titleAliasDictionary', 'mangaUpdates', 'aniList', 'bookWalker', 'amazon', 'myAnimeList', 'mangaDex'],
    bl_manga: ['titleAliasDictionary', 'chilchil', 'mangaUpdates', 'aniList', 'bookWalker', 'amazon'],
    light_novel: ['titleAliasDictionary', 'bookWalker', 'amazon', 'googleBooks', 'openBD', 'aniList', 'mangaUpdates'],
    novel_book: ['titleAliasDictionary', 'bookWalker', 'amazon', 'googleBooks', 'openBD', 'ndlSearch'],
    goods: ['titleAliasDictionary', 'aniList', 'mangaUpdates', 'bookWalker', 'amazon'],
    magazine: ['magazineMaster', 'googleBooks', 'amazon', 'bookWalker'],
    music_video: ['titleAliasDictionary', 'amazon'],
    unknown: ['titleAliasDictionary', 'bookWalker', 'amazon', 'googleBooks']
  };

  function compact(text) {
    return String(text || '')
      .replace(/\u3000/g, ' ')
      .replace(/[（）]/g, c => (c === '（' ? '(' : ')'))
      .replace(/[［］]/g, c => (c === '［' ? '[' : ']'))
      .replace(/[【】]/g, c => (c === '【' ? '[' : ']'))
      .replace(/[：]/g, ':')
      .replace(/[‐−–—―]/g, '-')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function stripWrapping(text) {
    return compact(text)
      .replace(/^[《「『【\[\(<＜]+/, '')
      .replace(/[》」』】\]\)>＞]+$/, '')
      .trim();
  }

  function sourceOf(rawItem) {
    const source = String(rawItem?.source || '').trim();
    if (source) return source;
    const url = String(rawItem?.pageUrl || rawItem?.url || '').trim();
    if (/aladin\.co\.kr/i.test(url)) return 'aladin';
    return 'aladin';
  }

  function rawTitleOf(rawItem) {
    return compact(
      rawItem?.title ||
      rawItem?.productTitle ||
      rawItem?.originalProductTitle ||
      rawItem?.name ||
      ''
    );
  }

  function extractVolume(text) {
    const source = compact(text);
    const patterns = [
      /\bvol\.?\s*(\d{1,3})\b/i,
      /\bvolume\s*(\d{1,3})\b/i,
      /(?:^|\s)(\d{1,3})\s*(?:권|卷|巻|集)\b/i,
      /(?:^|\s)(\d{1,3})(?=\s*(?:한정판|초판|특장판|限定|特装|通常|$))/i,
      /[（(]\s*(\d{1,3})\s*[)）]/
    ];
    for (const pattern of patterns) {
      const match = source.match(pattern);
      if (match) return match[1];
    }
    return '';
  }

  function extractEdition(text) {
    const source = compact(text);
    const patterns = [
      /(首刷限定版|初版限定版|首刷版|初版版|特裝版|特装版|限定版|通常版|愛蔵版|完全版)/u,
      /(한정판|초판\s*한정판|초판|특장판|일반판|완전판|소장판|애장판)/u,
      /\[([^\]]*(?:限定|한정|초판|특장판|通常|일반)[^\]]*)\]/iu,
      /\(([^)]*(?:限定|한정|초판|특장판|通常|일반)[^)]*)\)/iu
    ];
    for (const pattern of patterns) {
      const match = source.match(pattern);
      if (match) return compact(match[1]);
    }
    return '';
  }

  function extractMagazineInfo(rawItem) {
    const title = rawTitleOf(rawItem);
    const combined = compact(`${title} ${rawItem?.categoryName || ''}`);
    const result = { magazineName: '', year: '', month: '', issue: '' };

    for (const entry of MAGAZINE_PATTERNS) {
      if (entry.pattern.test(combined)) {
        result.magazineName = entry.name;
        break;
      }
    }
    const hasMagazineSignal = /magazine|잡지|매거진|雑誌/i.test(combined)
      || /\b20\d{2}[./-]\d{1,2}\b/.test(combined)
      || /\b20\d{2}\s*(?:年|년)\s*\d{1,2}\s*(?:月|월)\b/u.test(combined)
      || /\d{2,5}\s*(?:号|號|호)\b/u.test(combined);
    if (!result.magazineName && !hasMagazineSignal) {
      return result;
    }
    if (!result.magazineName) {
      result.magazineName = compact(title
        .replace(/\b20\d{2}[./-]\d{1,2}\b.*$/i, '')
        .replace(/\b\d{4}년\s*\d{1,2}월.*$/i, '')
        .replace(/\b\d{4}年\s*\d{1,2}月.*$/i, '')
        .replace(/\s+(KOREA|코리아)\b/ig, '')
        .replace(/[<＜【\[(].*$/, ''));
    }

    let match = combined.match(/\b(20\d{2})[./-](\d{1,2})\b/);
    if (!match) match = combined.match(/\b(20\d{2})\s*(?:年|년)\s*(\d{1,2})\s*(?:月|월)\b/u);
    if (match) {
      result.year = match[1];
      result.month = String(Number(match[2]));
    }
    const issueMatch = combined.match(/(\d{2,5})\s*(?:号|號|호)\b/u);
    if (issueMatch) result.issue = issueMatch[1];
    return result;
  }

  function stripCommonNoise(text) {
    return compact(text)
      .replace(/\[[^\]]*(?:限定|한정|초판|특장판|부록|포토카드|특전|예약|予約|特典)[^\]]*\]/giu, ' ')
      .replace(/\([^)]*(?:限定|한정|초판|특장판|부록|포토카드|특전|예약|予約|特典)[^)]*\)/giu, ' ')
      .replace(/(?:首刷限定版|初版限定版|首刷版|特裝版|特装版|限定版|通常版|한정판|초판\s*한정판|초판|특장판|일반판)/giu, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function stripKrGoodsNoise(text) {
    const source = stripCommonNoise(text);
    const goodsPattern = /(?:랜덤\s*)?(?:SD\s*)?(?:아크릴\s*(?:스탠드|키링|참|코롯토|블럭|액자)?|스탠드|키링|포토카드|엽서|포스터|스티커|틴케이스|인형|카드|배지|뱃지|캔뱃지|클리어파일|파일|마우스패드|쿠션|담요|텀블러|머그|파우치|북마크|책갈피|색지|태피스트리|굿즈)(?:[\s\w가-힣-]*)*/iu;
    const match = source.match(goodsPattern);
    if (!match) return { title: source, goodsType: '', optionType: '' };
    const before = source.slice(0, match.index).trim();
    const after = source.slice(match.index + match[0].length).trim();
    const optionMatch = after.match(/^(?:ver\.?\s*)?([A-Z]|\d{1,3})\b/i);
    return {
      title: before || source.replace(match[0], '').trim(),
      goodsType: compact(match[0]),
      optionType: optionMatch ? optionMatch[1] : ''
    };
  }

  function isBLLike(rawItem, parsed) {
    const text = [
      rawTitleOf(rawItem),
      rawItem?.categoryName,
      rawItem?.mallType,
      rawItem?.label,
      parsed?.edition,
      parsed?.goodsType
    ].join(' ').toLowerCase();
    return /(?:\bbl\b|ボーイズラブ|耽美|danmei|비엘|브로맨스|blコミック|blノベル|yaoi|boys'? love)/i.test(text);
  }

  function detectItemType(rawItem, parsed) {
    const title = rawTitleOf(rawItem);
    const text = [
      title,
      rawItem?.categoryName,
      rawItem?.mallType,
      rawItem?.description,
      rawItem?.basicInfo
    ].join(' ').toLowerCase();

    const magazineInfo = extractMagazineInfo(rawItem);
    const hasMagazineIssue = Boolean(magazineInfo.year || magazineInfo.issue);
    if (/magazine|잡지|매거진|雑誌/i.test(text) || (magazineInfo.magazineName && hasMagazineIssue)) return 'magazine';
    if (/굿즈|goods|gift|아크릴|스탠드|키링|포토카드|엽서|포스터|스티커|틴케이스|인형|배지|뱃지|色紙|壓克力|立牌|鑰匙圈|徽章/i.test(text) || parsed?.goodsType) return 'goods';
    if (/\b(?:cd|dvd|blu-ray|bluray|ost|album|lp)\b|음반|블루레이/i.test(text)) return 'music_video';
    if (/light\s*novel|라이트노벨|小説|소설|ノベライズ|novel/i.test(text)) return isBLLike(rawItem, parsed) ? 'bl_manga' : 'light_novel';
    if (/comic|comics|만화|코믹|webtoon|漫畫|漫画|まんが|マンガ/i.test(text)) return isBLLike(rawItem, parsed) ? 'bl_manga' : 'manga';
    if (/book|도서|書籍|本|essay|에세이/i.test(text)) return 'novel_book';
    return 'unknown';
  }

  function removeVolumeAndEdition(text, volume, edition) {
    let result = compact(text);
    if (edition) {
      result = result.replace(new RegExp(escapeRegExp(edition), 'giu'), ' ');
    }
    if (volume) {
      result = result
        .replace(new RegExp(`(?:^|\\s)${escapeRegExp(volume)}\\s*(?:권|卷|巻|集)\\b`, 'giu'), ' ')
        .replace(new RegExp(`(?:^|\\s)${escapeRegExp(volume)}(?=\\s*$)`, 'giu'), ' ');
    }
    return compact(result);
  }

  function escapeRegExp(text) {
    return String(text || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  function splitSubtitle(text) {
    const title = compact(text);
    const match = title.match(/^(.+?)[~〜～]\s*(.+)$/u);
    if (!match) return { workTitle: title, subtitle: '' };
    return { workTitle: compact(match[1]), subtitle: compact(match[2]) };
  }

  function buildNormalizedSearchTitle(parsed) {
    const parts = [parsed?.extractedWorkTitle || '', parsed?.subtitle || '']
      .map(compact)
      .filter(Boolean);
    return parts.join(' ').trim() || compact(parsed?.originalTitle || parsed?.rawTitle || '');
  }

  function providerHintFor(itemType) {
    return {
      preferredOrder: PROVIDER_ORDER[itemType] || PROVIDER_ORDER.unknown,
      reason: `${itemType || 'unknown'} detected`
    };
  }

  function analyzeAladinTitle(rawItem) {
    const rawTitle = rawTitleOf(rawItem);
    const volume = extractVolume(rawTitle);
    const edition = extractEdition(rawTitle);
    const magazineInfo = extractMagazineInfo(rawItem);
    let working = removeVolumeAndEdition(rawTitle, volume, edition);
    working = stripCommonNoise(working);

    const goodsParts = stripKrGoodsNoise(working);
    if (goodsParts.goodsType) working = goodsParts.title;
    const subtitleParts = splitSubtitle(working);

    const parsed = {
      source: 'aladin',
      rawTitle,
      originalProductTitle: rawTitle,
      originalTitle: stripWrapping(subtitleParts.workTitle || working || rawTitle),
      extractedWorkTitle: stripWrapping(subtitleParts.workTitle || working || rawTitle),
      subtitle: subtitleParts.subtitle,
      normalizedSearchTitle: '',
      itemType: 'unknown',
      language: '韓国',
      category: rawItem?.categoryName || rawItem?.mallType || '',
      author: rawItem?.author || '',
      volume,
      edition,
      bonusInfo: '',
      goodsType: goodsParts.goodsType,
      optionType: goodsParts.optionType,
      magazineName: magazineInfo.magazineName || '',
      year: magazineInfo.year || '',
      month: magazineInfo.month || '',
      issue: magazineInfo.issue || '',
      confidence: 0.7,
      warnings: []
    };

    parsed.itemType = detectItemType(rawItem, parsed);
    if (parsed.itemType === 'magazine') {
      parsed.extractedWorkTitle = parsed.magazineName || parsed.extractedWorkTitle;
      parsed.originalTitle = parsed.magazineName || parsed.originalTitle;
      parsed.normalizedSearchTitle = parsed.magazineName || parsed.extractedWorkTitle || rawTitle;
    } else {
      parsed.normalizedSearchTitle = buildNormalizedSearchTitle(parsed);
    }
    parsed.providerHint = providerHintFor(parsed.itemType);
    if (!parsed.normalizedSearchTitle) {
      parsed.normalizedSearchTitle = rawTitle;
      parsed.warnings.push('normalizedSearchTitle fallback to rawTitle');
    }
    return parsed;
  }

  function analyzeKoreanGoodsTitle(rawItem) {
    const parsed = analyzeAladinTitle({ ...rawItem, source: 'korean_goods' });
    parsed.source = 'korean_goods';
    parsed.itemType = 'goods';
    parsed.providerHint = providerHintFor(parsed.itemType);
    return parsed;
  }

  function analyzeBooksTwTitle(rawItem) {
    const rawTitle = rawTitleOf(rawItem);
    return {
      source: 'books_tw',
      rawTitle,
      originalProductTitle: rawTitle,
      originalTitle: rawTitle,
      extractedWorkTitle: rawTitle,
      subtitle: '',
      normalizedSearchTitle: rawTitle,
      itemType: 'unknown',
      language: '台湾',
      category: rawItem?.categoryName || '',
      author: rawItem?.author || '',
      volume: extractVolume(rawTitle),
      edition: extractEdition(rawTitle),
      bonusInfo: '',
      goodsType: '',
      optionType: '',
      magazineName: '',
      year: '',
      month: '',
      issue: '',
      confidence: 0.4,
      warnings: ['books_tw analysis is handled by the integrated taiwan extension'],
      providerHint: providerHintFor('unknown')
    };
  }

  function analyzeProductTitle(rawItem) {
    try {
      const source = sourceOf(rawItem);
      if (source === 'books_tw') return analyzeBooksTwTitle(rawItem);
      if (source === 'korean_goods') return analyzeKoreanGoodsTitle(rawItem);
      return analyzeAladinTitle({ ...rawItem, source: 'aladin' });
    } catch (error) {
      const rawTitle = rawTitleOf(rawItem);
      return {
        source: sourceOf(rawItem),
        rawTitle,
        originalProductTitle: rawTitle,
        originalTitle: rawTitle,
        extractedWorkTitle: rawTitle,
        subtitle: '',
        normalizedSearchTitle: rawTitle,
        itemType: 'unknown',
        language: sourceOf(rawItem) === 'books_tw' ? '台湾' : '韓国',
        category: rawItem?.categoryName || '',
        author: rawItem?.author || '',
        volume: '',
        edition: '',
        bonusInfo: '',
        goodsType: '',
        optionType: '',
        magazineName: '',
        year: '',
        month: '',
        issue: '',
        confidence: 0.1,
        warnings: [`analysis failed: ${error.message || error}`],
        providerHint: providerHintFor('unknown')
      };
    }
  }

  function isValidGasWebAppUrl(url) {
    var u = String(url || '').trim();
    return /^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec(?:[?#].*)?$/i.test(u)
      || /^https:\/\/script\.google\.com\/home\/macros\/exec\?[^#]*(?:\bid=|\/exec\b)/i.test(u)
      || /^https:\/\/script\.googleusercontent\.com\/macros\/exec\b/i.test(u);
  }

  function buildLookupPayload(rawItem, titleAnalysis) {
    const analysis = titleAnalysis || analyzeProductTitle(rawItem);
    return {
      action: 'lookupJapaneseTitle',
      source: analysis.source || sourceOf(rawItem),
      rawItem: rawItem || {},
      titleAnalysis: analysis
    };
  }

  function normalizeLookup(lookup) {
    if (!lookup) return null;
    if (lookup.lookup) return normalizeLookup(lookup.lookup);
    return {
      status: lookup.status || (lookup.japaneseTitle || lookup.title ? 'resolved' : 'not_found'),
      japaneseTitle: lookup.japaneseTitle || lookup.title || '',
      provider: lookup.provider || lookup.source || '',
      normalizedSearchTitle: lookup.normalizedSearchTitle || '',
      extractedWorkTitle: lookup.extractedWorkTitle || '',
      score: lookup.score || '',
      trace: lookup.trace || lookup.log || '',
      candidates: Array.isArray(lookup.candidates) ? lookup.candidates : [],
      errors: Array.isArray(lookup.errors) ? lookup.errors : []
    };
  }

  async function tryMangaUpdatesDirectJapanese_(titleAnalysis) {
    const g = globalThis.titleLookupMangaUpdates;
    if (!g || typeof g.lookupJapaneseTitle !== 'function') return '';
    const itemType = String(titleAnalysis && titleAnalysis.itemType || '').trim();
    const tryTypes = new Set(['manga', 'bl_manga', 'goods', 'light_novel', 'unknown']);
    if (!tryTypes.has(itemType)) return '';
    const q = String(
      titleAnalysis && (titleAnalysis.normalizedSearchTitle || titleAnalysis.extractedWorkTitle) || ''
    ).trim();
    if (!q) return '';
    try {
      return String(await g.lookupJapaneseTitle(q)).trim();
    } catch (_) {
      return '';
    }
  }

  async function enrichLookupWithDirectMangaUpdates_(titleAnalysis, json) {
    const lookup = normalizeLookup(json && json.lookup);
    if (!lookup) return json;
    if (String(lookup.japaneseTitle || '').trim()) return { ...json, lookup };
    const g = globalThis.titleLookupMangaUpdates;
    if (!g || typeof g.lookupJapaneseTitle !== 'function') return { ...json, lookup };
    const itemType = String(titleAnalysis && titleAnalysis.itemType || '').trim();
    const tryTypes = new Set(['manga', 'bl_manga', 'goods', 'light_novel', 'unknown']);
    if (!tryTypes.has(itemType)) return { ...json, lookup };
    const q = String(
      titleAnalysis && (titleAnalysis.normalizedSearchTitle || titleAnalysis.extractedWorkTitle) || ''
    ).trim();
    if (!q) return { ...json, lookup };
    try {
      const jp = String(await g.lookupJapaneseTitle(q)).trim();
      if (!jp) return { ...json, lookup };
      return {
        ...json,
        lookup: {
          ...lookup,
          status: 'resolved',
          japaneseTitle: jp,
          provider: lookup.provider ? `${lookup.provider}+mangaupdatesClient` : 'mangaUpdates(extension)',
          trace: `${lookup.trace || ''} | mangaupdates:direct_hit`.replace(/^\s*\|\s*/, '').trim(),
        },
      };
    } catch (_) {
      return { ...json, lookup };
    }
  }

  async function lookupJapaneseTitle(titleAnalysis, options) {
    const gasUrl = String(options?.gasUrl || '').trim();
    if (!isValidGasWebAppUrl(gasUrl)) {
      return {
        success: false,
        lookup: {
          status: 'skipped',
          japaneseTitle: '',
          provider: '',
          normalizedSearchTitle: titleAnalysis?.normalizedSearchTitle || '',
          trace: 'GAS URL is not a deployed web app URL',
          candidates: [],
          errors: []
        }
      };
    }
    const gasPromise = fetch(gasUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(buildLookupPayload(options?.rawItem || {}, titleAnalysis))
    }).then(async response => {
      if (!response.ok) throw new Error(`GAS lookup HTTP ${response.status}`);
      return response.json();
    });
    const [json, muEarly] = await Promise.all([gasPromise, tryMangaUpdatesDirectJapanese_(titleAnalysis)]);
    let merged = { ...json, lookup: normalizeLookup(json.lookup) };
    if (muEarly && !String(merged.lookup?.japaneseTitle || '').trim()) {
      merged = {
        ...merged,
        lookup: {
          ...merged.lookup,
          status: 'resolved',
          japaneseTitle: muEarly,
          provider: merged.lookup.provider ? `${merged.lookup.provider}+mangaupdatesClient` : 'mangaUpdates(extension)',
          trace: `${merged.lookup.trace || ''} | mangaupdates:parallel_direct`.replace(/^\s*\|\s*/, '').trim(),
        },
      };
    }
    merged = await enrichLookupWithDirectMangaUpdates_(titleAnalysis, merged);
    return merged;
  }

  function applyJapaneseTitleLookupToProduct(product, lookupResponse) {
    const lookup = normalizeLookup(lookupResponse?.lookup || lookupResponse);
    return {
      ...product,
      japaneseTitleLookup: lookup,
      日本語タイトル: lookup?.japaneseTitle || product?.日本語タイトル || '',
      titleAnalysis: product?.titleAnalysis || analyzeProductTitle(product)
    };
  }

  function buildSavePayload(rawItem, titleAnalysis, japaneseTitleLookup) {
    const analysis = titleAnalysis || rawItem?.titleAnalysis || analyzeProductTitle(rawItem);
    return {
      action: 'upsertProductWithLookup',
      source: analysis.source || sourceOf(rawItem),
      rawItem: rawItem || {},
      titleAnalysis: analysis,
      japaneseTitleLookup: normalizeLookup(japaneseTitleLookup || rawItem?.japaneseTitleLookup)
    };
  }

  function renderLookupResult(result) {
    const lookup = normalizeLookup(result?.lookup || result);
    if (!lookup) return '日本語タイトル照会: 未実行';
    if (lookup.status === 'resolved') {
      return `日本語タイトル: ${lookup.japaneseTitle || '(空)'} / ${lookup.provider || 'provider不明'}`;
    }
    if (lookup.status === 'not_found') return `日本語タイトル: 登録なし / ${lookup.trace || '照会済み'}`;
    if (lookup.status === 'partial_error') return `日本語タイトル: 一部照会失敗 / ${lookup.trace || lookup.provider || ''}`;
    if (lookup.status === 'failed') return `日本語タイトル: 照会失敗 / ${lookup.trace || lookup.provider || ''}`;
    if (lookup.status === 'skipped') return `日本語タイトル: 未照会 / ${lookup.trace || ''}`;
    return `日本語タイトル: ${lookup.status || '不明'} / ${lookup.trace || ''}`;
  }

  Object.assign(global, {
    analyzeProductTitle,
    analyzeAladinTitle,
    analyzeBooksTwTitle,
    analyzeKoreanGoodsTitle,
    detectItemType,
    isBLLike,
    extractVolume,
    extractEdition,
    extractMagazineInfo,
    stripCommonNoise,
    stripKrGoodsNoise,
    buildNormalizedSearchTitle,
    buildLookupPayload,
    buildSavePayload,
    lookupJapaneseTitle,
    applyJapaneseTitleLookupToProduct,
    renderLookupResult,
    isValidGasWebAppUrl
  });
})(globalThis);
