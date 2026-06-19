(function initTitleAnalysis(global) {
  const MAGAZINE_BRANDS = [
    'ELLE', 'VOGUE', 'GQ', 'ESQUIRE', 'ALLURE', 'BAZAAR', 'HARPER\'S BAZAAR',
    'MARIE CLAIRE', 'COSMOPOLITAN', 'DAZED', 'ARENA', 'W KOREA', '1ST LOOK',
    'CINE21', 'MAXIM', 'SINGLES', 'NYLON', 'THE STAR'
  ];

  const PROVIDER_HINTS = {
    manga: ['titleAliasDictionary', 'mangaUpdates', 'aniList', 'bookWalker', 'amazon', 'myAnimeList', 'mangaDex'],
    bl_manga: ['titleAliasDictionary', 'chilchil', 'mangaUpdates', 'aniList', 'bookWalker', 'amazon'],
    light_novel: ['titleAliasDictionary', 'bookWalker', 'amazon', 'googleBooks', 'openBD', 'aniList', 'mangaUpdates'],
    novel_book: ['titleAliasDictionary', 'bookWalker', 'amazon', 'googleBooks', 'openBD', 'ndlSearch'],
    goods: ['titleAliasDictionary', 'aniList', 'mangaUpdates', 'bookWalker', 'amazon'],
    magazine: ['magazineMaster', 'googleBooks', 'amazon', 'bookWalker'],
    music_video: ['titleAliasDictionary', 'amazon'],
    unknown: ['titleAliasDictionary', 'bookWalker', 'amazon', 'googleBooks']
  };

  function toText(value) {
    return String(value ?? '').trim();
  }

  function compactSpaces(value) {
    return toText(value)
      .replace(/\u3000/g, ' ')
      .replace(/[（]/g, '(')
      .replace(/[）]/g, ')')
      .replace(/[［]/g, '[')
      .replace(/[］]/g, ']')
      .replace(/[｛]/g, '{')
      .replace(/[｝]/g, '}')
      .replace(/[＋]/g, '+')
      .replace(/[－–—―]/g, '-')
      .replace(/[：]/g, ':')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function uniq(values) {
    const seen = new Set();
    const list = [];
    for (const value of values || []) {
      const text = compactSpaces(value);
      if (!text || seen.has(text)) continue;
      seen.add(text);
      list.push(text);
    }
    return list;
  }

  function getField(rawItem, keys) {
    for (const key of keys) {
      const value = rawItem?.[key];
      if (value != null && String(value).trim()) return String(value).trim();
    }
    return '';
  }

  function getRawTitle(rawItem) {
    return getField(rawItem, [
      'title', '商品名', 'name', 'productTitle', '原題商品タイトル',
      '原題商品名', '商品名（原題）', 'ページタイトル'
    ]);
  }

  function normalizeSourceName(value) {
    const text = toText(value).toLowerCase();
    if (!text) return '';
    if (/books\.com\.tw|books_tw|bookstw|博客來|博客来/.test(text)) return 'books_tw';
    if (/aladin|アラジン|알라딘/.test(text)) return 'aladin';
    if (/korean[_\s-]*goods|韓国グッズ|한국.*굿즈/.test(text)) return 'korean_goods';
    return text;
  }

  function detectSourceFromUrl(url) {
    const text = toText(url);
    if (/books\.com\.tw/i.test(text)) return 'books_tw';
    if (/aladin\.co\.kr/i.test(text)) return 'aladin';
    return 'unknown';
  }

  function detectSource(rawItem) {
    const explicit = toText(rawItem?.source || rawItem?.sourceSite || rawItem?.site || rawItem?.mallType).toLowerCase();
    const url = toText(rawItem?.pageUrl || rawItem?.URL || rawItem?.url);
    const sourceFromUrl = detectSourceFromUrl(url);
    if (sourceFromUrl !== 'unknown') return sourceFromUrl;
    const normalized = normalizeSourceName(explicit);
    if (normalized) return normalized;
    if (/aladin\.co\.kr/i.test(url) || /aladin|アラジン|알라딘/.test(explicit)) return 'aladin';
    if (/korean[_\s-]*goods|韓国グッズ|한국.*굿즈/.test(explicit)) return 'korean_goods';
    return explicit || 'unknown';
  }

  function stripBracketNoise(text) {
    return compactSpaces(text)
      .replace(/\[[^\]]*(?:예약|특전|초판|한정|限定|特典|首刷|初回|初版|特[裝装]|通常|贈品|赠品|予約|預購|预购)[^\]]*\]/gi, ' ')
      .replace(/【[^】]*(?:예약|특전|초판|한정|限定|特典|首刷|初回|初版|特[裝装]|通常|贈品|赠品|予約|預購|预购)[^】]*】/gi, ' ')
      .replace(/\([^)]*(?:예약|특전|초판|한정|限定|特典|首刷|初回|初版|特[裝装]|通常|贈品|赠品|予約|預購|预购)[^)]*\)/gi, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function extractEdition(text) {
    const source = compactSpaces(text);
    const patterns = [
      /首刷限定版/u,
      /初版限定版/u,
      /首刷(?:限定)?/u,
      /初回(?:限定)?/u,
      /特[裝装]版/u,
      /限定版/u,
      /通常版/u,
      /한정판/i,
      /초판(?:한정)?/i,
      /특장판/i,
      /일반판/i,
      /limited\s*edition/i,
      /special\s*edition/i
    ];
    for (const pattern of patterns) {
      const match = source.match(pattern);
      if (match) return match[0];
    }
    return '';
  }

  function extractVolume(text) {
    const source = stripBracketNoise(compactSpaces(text));
    const candidates = [
      /(?:^|\s)(?:第\s*)?(\d{1,3})\s*(?:巻|卷|集|冊|册|권|권째|화|話|号|號)\b/iu,
      /(?:vol(?:ume)?\.?\s*)(\d{1,3})\b/iu,
      /(?:^|[^\d])(\d{1,3})\s*(?:漫畫|漫画|コミック|單行本|单行本)\s*$/iu,
      /(?:^|\s)(\d{1,3})\s*(?:권|巻|卷)\b/iu,
      /\((\d{1,3})\)\s*$/u,
      /[)\]】》」』]\s*(\d{1,3})\s*$/u,
      /(?:^|\s)(\d{1,3})\s*(?:首刷|初版|初回|限定|特[裝装]|한정판|초판|특장판|通常|$)/iu
    ];
    for (const pattern of candidates) {
      const match = source.match(pattern);
      if (match?.[1]) return match[1].replace(/^0+/, '') || '0';
    }
    return '';
  }

  function extractMagazineInfo(rawItem) {
    const title = compactSpaces(getRawTitle(rawItem));
    const source = [
      title,
      rawItem?.categoryName,
      rawItem?.カテゴリ,
      rawItem?.mallType,
      rawItem?.ジャンル
    ].map(toText).filter(Boolean).join(' ');

    const upper = source.toUpperCase();
    let magazineName = '';
    for (const brand of MAGAZINE_BRANDS) {
      const pattern = new RegExp(`(^|[^A-Z0-9])${brand.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}([^A-Z0-9]|$)`, 'i');
      if (pattern.test(upper)) {
        magazineName = brand === 'HARPER\'S BAZAAR' ? 'BAZAAR' : brand;
        break;
      }
    }

    const yearMonth = source.match(/\b(20\d{2})[.\-/年\s]+(1[0-2]|0?[1-9])\s*(?:月|호|号|號)?\b/u);
    const issue = source.match(/(?:第\s*)?(\d{2,5})\s*(?:号|號|호|期)\b/u)
      || source.match(/\b(?:NO\.?|ISSUE|VOL\.?)\s*(\d{1,5})\b/i);
    const noIssueTitle = title.match(/^(.+?)\s+(?:NO\.?|ISSUE|VOL\.?)\s*\d{1,5}\b.*$/i);
    if (!magazineName && noIssueTitle?.[1]) {
      magazineName = compactSpaces(noIssueTitle[1]);
    }

    return {
      magazineName,
      year: yearMonth?.[1] || '',
      month: yearMonth?.[2] ? String(Number(yearMonth[2])) : '',
      issue: issue?.[1] || ''
    };
  }

  function stripCommonNoise(text) {
    let value = compactSpaces(text);
    value = stripBracketNoise(value);
    value = value
      .replace(/^\s*(?:小説|小說|小说|漫画|漫畫|コミック|만화|코믹|소설|라이트노벨|light\s*novel)\s+/iu, '')
      .replace(/\s*(?:首刷(?:限定)?版?|初版(?:限定)?版?|初回(?:限定)?版?|限定版|通常版|特[裝装]版|한정판|초판(?:한정)?|특장판|일반판)\s*$/iu, '')
      .replace(/\s*(?:預購|预购|予約|예약|贈品|赠品|特典|특전).*/iu, ' ')
      .replace(/\s*\((?:完|全)\)\s*$/u, '')
      .replace(/\s+/g, ' ')
      .trim();
    const volume = extractVolume(value);
    if (volume) {
      value = value
        .replace(new RegExp(`\\s*\\(?0*${volume}\\)?\\s*$`, 'u'), '')
        .replace(new RegExp(`\\s*0*${volume}\\s*(?:권|巻|卷|集|冊|册|話|号|號|漫畫|漫画|コミック|單行本|单行本)\\s*$`, 'iu'), '')
        .trim();
    }
    return value;
  }

  function stripWorkTitleNoise(text) {
    let value = compactSpaces(text);
    value = stripBracketNoise(value)
      .replace(/\s*(?:首刷(?:限定)?版?|初版(?:限定)?版?|初回(?:限定)?版?|限定版|通常版|特[裝装]版|한정판|초판(?:한정)?|특장판|일반판)\s*/giu, ' ')
      .replace(/\s+/g, ' ')
      .trim();

    const volume = extractVolume(value);
    if (volume) {
      value = value
        .replace(new RegExp(`\\s*[（(]?0*${volume}[）)]?\\s*$`, 'u'), '')
        .replace(new RegExp(`\\s*0*${volume}\\s*(?:巻|卷|集|冊|册|권|話|号|號|漫畫|漫画|コミック|單行本|单行本)\\s*$`, 'iu'), '')
        .trim();
    }

    return compactSpaces(value);
  }

  function splitSubtitle(text) {
    const value = compactSpaces(text).replace(/[《》「」『』]/g, '').trim();
    const wave = value.match(/^(.+?)[~～〜](.+?)[~～〜]?$/u);
    if (wave?.[1] && wave?.[2]) {
      return { workTitle: wave[1].trim(), subtitle: wave[2].trim() };
    }
    const dash = value.match(/^(.{2,40}?)\s+[-:：]\s+(.{2,80})$/u);
    if (dash?.[1] && dash?.[2]) {
      return { workTitle: dash[1].trim(), subtitle: dash[2].trim() };
    }
    return { workTitle: value, subtitle: '' };
  }

  function stripGoodsWorkTitleDecoration(text) {
    return compactSpaces(text)
      .replace(/\s*[!！?？]{1,4}\s*$/u, '')
      .trim();
  }

  function stripTwGoodsNoise(text) {
    const source = compactSpaces(text)
      .replace(/\((?:漫畫|漫画|小說|小说|小説|動畫|动画|アニメ|電影|映画)[^)]*\)/gu, ' ')
      .replace(/（(?:漫畫|漫画|小說|小说|小説|動畫|动画|アニメ|電影|映画)[^）]*）/gu, ' ')
      .replace(/\s+/g, ' ')
      .trim();

    // 末尾は A〜Z の品番、(A1)、または「2入」「3入」など入数のみ許容（資料夾2入 等）
    let match = source.match(/^(.+?)(?:\s+|[-:：])((?:[\p{Script=Han}A-Za-z0-9]+\s*){0,5}(?:壓克力色紙|压克力色纸|壓克力便利夾|压克力便利夹|壓克力筋牌|压克力筋牌|壓克力立牌|压克力立牌|壓克力|压克力|多功能手機包|手機包|手机包|透明立方|立牌|色紙|色纸|貼紙|贴纸|明信片|鑰匙圈|钥匙圈|徽章|胸章|海報|海报|掛軸|挂轴|門簾|门帘|卡片|杯墊|杯垫|資料夾|资料夹|文件夾|文件夹|吊飾|吊饰|抱枕|托特袋|帆布袋|手提袋))(?:\s+([A-Z0-9]+)|\(([A-Z0-9]+)\)|\s*[\d０-９]{1,4}\s*入)?\s*$/u);
    if (!match) {
      match = source.match(/^(.+?)((?:壓克力色紙|压克力色纸|壓克力便利夾|压克力便利夹|壓克力筋牌|压克力筋牌|壓克力立牌|压克力立牌|壓克力|压克力|多功能手機包|手機包|手机包|透明立方|立牌|色紙|色纸|貼紙|贴纸|明信片|鑰匙圈|钥匙圈|徽章|胸章|海報|海报|掛軸|挂轴|門簾|门帘|卡片|杯墊|杯垫|資料夾|资料夹|文件夾|文件夹|吊飾|吊饰|抱枕|托特袋|帆布袋|手提袋))(?:\s+([A-Z0-9]+)|\(([A-Z0-9]+)\)|\s*[\d０-９]{1,4}\s*入)?\s*$/u);
    }
    if (!match) {
      return { title: source, goodsType: '', optionType: '', bonusInfo: '' };
    }

    const rawTitle = stripGoodsWorkTitleDecoration(
      compactSpaces(match[1]).replace(/([\p{Script=Han}ぁ-ゖァ-ヺー々〆ヵヶ]{2,})\s*[A-Z]{1,4}$/u, '$1')
    );
    return {
      title: rawTitle,
      goodsType: compactSpaces(match[2]),
      optionType: compactSpaces(match[3] || match[4] || ''),
      bonusInfo: ''
    };
  }

  function stripKrGoodsNoise(text) {
    const source = compactSpaces(text);
    const match = source.match(/^(.+?)\s+((?:랜덤\s+)?(?:SD\s+)?(?:아크릴\s*(?:스탠드|키링|카드|블럭|보드)|포토카드|엽서|포스터|스티커|틴케이스|인형|카드|키링|클리어\s*파일|파일|배지|뱃지|캔뱃지|마우스패드|담요|쿠션)(?:\s+[A-Z0-9]+)?)\s*$/iu);
    if (!match) {
      return { title: source, goodsType: '', optionType: '', bonusInfo: '' };
    }

    return {
      title: compactSpaces(match[1]),
      goodsType: compactSpaces(match[2]),
      optionType: '',
      bonusInfo: ''
    };
  }

  function isBLLike(rawItem, parsed = {}) {
    try {
      const source = [
        getRawTitle(rawItem),
        rawItem?.categoryName,
        rawItem?.カテゴリ,
        rawItem?.label,
        rawItem?.ジャンル,
        rawItem?.mallType,
        parsed.extractedWorkTitle
      ].map(toText).filter(Boolean).join(' ');
      return /(?:\bBL\b|ボーイズラブ|耽美|danmei|브로맨스|비엘|blコミック|blノベル|BL漫畫|BL漫画|BL소설|BL만화)/i.test(source);
    } catch {
      return false;
    }
  }

  function detectItemType(rawItem, parsed = {}) {
    const title = getRawTitle(rawItem);
    const magazine = parsed.magazineName || extractMagazineInfo(rawItem).magazineName;
    const source = [
      title,
      rawItem?.categoryName,
      rawItem?.カテゴリ,
      rawItem?.mallType,
      rawItem?.ジャンル,
      parsed.goodsType
    ].map(toText).filter(Boolean).join(' ');

    if (magazine && (parsed.year || parsed.month || /magazine|잡지|매거진|雑誌|雜誌|杂志/i.test(source))) return 'magazine';
    if (parsed.goodsType) return 'goods';
    if (/(?:굿즈|goods|グッズ|아크릴|스탠드|키링|포토카드|엽서|포스터|스티커|壓克力|压克力|立牌|鑰匙圈|钥匙圈|色紙|色纸|貼紙|贴纸|明信片|海報|海报|徽章|手機包|手机包|資料夾|资料夹|文件夾|文件夹|周邊|周边|動漫周邊|动漫周边)/i.test(source)) return 'goods';
    if (/(?:Blu-?ray|DVD|CD|OST|album|LP|음반|블루레이|앨범|アルバム|音樂|音乐)/i.test(source)) return 'music_video';
    if (isBLLike(rawItem, parsed)) return 'bl_manga';
    if (/(?:light\s*novel|라이트노벨|輕小說|轻小说|ライトノベル|ノベライズ)/i.test(source)) return 'light_novel';
    if (/^\s*(?:小説|小說|小说|소설)\s+/u.test(title)) return 'light_novel';
    if (/(?:comic|comics|만화|코믹|webtoon|漫畫|漫画|まんが|コミック)/i.test(source)) return 'manga';
    if (/(?:小説|小說|小说|소설|novel|文學|文学|エッセイ|一般書籍|実用書)/i.test(source)) return 'novel_book';
    if (/[가-힣]/.test(title) && parsed.volume && !parsed.goodsType) return 'manga';
    return 'unknown';
  }

  /**
   * 検索キー用ノイズ語彙（タイトル辞書ではなく汎用ルール）
   * グッズ種別・数量・タイプ記号・版種・特典・末尾装飾を末尾から反復的に剥がす
   */
  const SEARCH_NOISE_PATTERNS = [
    // 数量（中国語/日本語/韓国語/英語）
    /[\s　]*[\d０-９]{1,4}\s*(?:入|件|個|組|입|セット|set|pcs?\.?|pack|枚|장)\s*$/iu,
    // 末尾のタイプ記号（A, B1, (A), （A）等）
    /[\s　]+[A-Z][0-9]?\s*$/u,
    /[\s　]*[（(\[【][A-Z][0-9]?[）)\]】]\s*$/u,
    // 数字単独の末尾コード（例: -01, _02）
    /[\s　]*[-_][\d０-９]{1,3}\s*$/u,
    // グッズ種別（中・台・韓・日）
    /[\s　]*(?:壓克力(?:立牌|色紙|筋牌|便利夾|キーホルダー)?|压克力(?:立牌|色纸|筋牌|便利夹)?|多功能手機包|多功能手机包|手機包|手机包|透明立方|立牌|色紙|色纸|貼紙|贴纸|明信片|鑰匙圈|钥匙圈|徽章|胸章|海報|海报|掛軸|挂轴|門簾|门帘|卡片|杯墊|杯垫|資料夾|资料夹|文件夾|文件夹|吊飾|吊饰|抱枕|托特袋|帆布袋|手提袋|周邊|周边|アクリル(?:スタンド|キーホルダー|ボード|ブロック)?|アクスタ|キーホルダー|缶バッジ|ポストカード|ステッカー|クリアファイル|ブロマイド|タペストリー|아크릴(?:스탠드|키링|카드|블럭|보드)?|포토카드|엽서|포스터|스티커|틴케이스|인형|키링|클리어\s*파일|뱃지|배지|캔뱃지|마우스패드|담요|쿠션)\s*$/iu,
    // 版種
    /[\s　]*(?:首刷(?:限定)?版?|初版(?:限定)?版?|初回(?:限定)?版?|限定版|通常版|特[裝装]版|한정판|초판(?:한정)?|특장판|일반판|限定|通常|首刷|初版|初回|초판|한정)\s*$/iu,
    // 特典・予約
    /[\s　]*(?:預購特典|预购特典|預購|预购|予約特典|予約|특전\s*포함|특전|예약|特典|贈品|赠品)\s*$/iu,
    // 末尾装飾（!!、★、❤ 等）
    /[\s　]*[!！]{1,3}\s*$/u,
    /[\s　]*[★☆♪♡❤❀＊♥◆◇♢♤♧]+\s*$/u,
    // 末尾の括弧付きノイズ（短いもの）
    /[\s　]*[（(][^）)]{1,15}[）)]\s*$/u,
    /[\s　]*【[^】]{1,15}】\s*$/u,
    /[\s　]*\[[^\]]{1,15}\]\s*$/u,
  ];

  function stripSearchNoise(text) {
    let value = compactSpaces(text);
    if (!value) return '';
    let prev;
    let iterations = 0;
    do {
      prev = value;
      for (const pattern of SEARCH_NOISE_PATTERNS) {
        value = value.replace(pattern, '').trim();
      }
      iterations += 1;
    } while (prev !== value && iterations < 12);
    return compactSpaces(value);
  }

  function buildNormalizedSearchTitle(parsed) {
    const parts = [parsed?.extractedWorkTitle, parsed?.subtitle]
      .map(compactSpaces)
      .filter(Boolean);
    const normalized = uniq(parts).join(' ').trim();
    const candidate = normalized || compactSpaces(parsed?.rawTitle || '');
    const cleaned = stripSearchNoise(candidate);
    // ノイズ除去で空になった場合は元の候補を返す
    return cleaned && cleaned.length >= 2 ? cleaned : candidate;
  }

  function buildProviderHint(itemType) {
    const preferredOrder = PROVIDER_HINTS[itemType] || PROVIDER_HINTS.unknown;
    return {
      preferredOrder,
      reason: `${itemType || 'unknown'} detected`
    };
  }

  function baseParsed(rawItem, source) {
    const rawTitle = compactSpaces(getRawTitle(rawItem));
    return {
      source,
      rawTitle,
      originalTitle: '',
      originalProductTitle: rawTitle,
      extractedWorkTitle: rawTitle,
      subtitle: '',
      normalizedSearchTitle: rawTitle,
      itemType: 'unknown',
      language: '',
      category: getField(rawItem, ['categoryName', 'カテゴリ', 'ジャンル', 'mallType']),
      author: getField(rawItem, ['author', '著者', '作者']),
      volume: extractVolume(rawTitle),
      edition: extractEdition(rawTitle),
      bonusInfo: '',
      goodsType: '',
      optionType: '',
      magazineName: '',
      year: '',
      month: '',
      issue: '',
      confidence: 0.55,
      warnings: []
    };
  }

  function finalizeParsed(rawItem, parsed) {
    if (parsed.magazineName && (parsed.year || parsed.month || parsed.issue)) {
      parsed.extractedWorkTitle = parsed.magazineName;
      parsed.subtitle = '';
    }
    parsed.extractedWorkTitle = stripWorkTitleNoise(parsed.extractedWorkTitle || parsed.originalTitle || parsed.rawTitle);
    parsed.originalTitle = stripWorkTitleNoise(parsed.originalTitle || parsed.extractedWorkTitle || parsed.rawTitle);
    parsed.normalizedSearchTitle = buildNormalizedSearchTitle(parsed);
    parsed.itemType = detectItemType(rawItem, parsed);
    parsed.providerHint = buildProviderHint(parsed.itemType);
    parsed.confidence = parsed.confidence || (parsed.normalizedSearchTitle === parsed.rawTitle ? 0.45 : 0.8);
    if (!parsed.normalizedSearchTitle) {
      parsed.normalizedSearchTitle = parsed.rawTitle;
      parsed.warnings.push('normalizedSearchTitle fallback to rawTitle');
    }
    return parsed;
  }

  function analyzeBooksTwTitle(rawItem) {
    const parsed = baseParsed(rawItem, 'books_tw');
    parsed.language = typeof global.getLanguageValue === 'function' ? global.getLanguageValue(rawItem) : getField(rawItem, ['言語', 'language']);

    const magazine = extractMagazineInfo(rawItem);
    Object.assign(parsed, magazine);

    let cleaned = stripCommonNoise(parsed.rawTitle);
    const twGoods = stripTwGoodsNoise(cleaned);
    if (twGoods.goodsType) {
      parsed.goodsType = twGoods.goodsType;
      parsed.optionType = twGoods.optionType;
      cleaned = twGoods.title;
      parsed.confidence = 0.86;
    }

    const split = splitSubtitle(cleaned);
    parsed.extractedWorkTitle = split.workTitle;
    parsed.subtitle = split.subtitle;
    parsed.bonusInfo = /(?:首刷|初版|限定|特典|贈品|赠品)/u.test(parsed.rawTitle) ? parsed.edition : '';
    return finalizeParsed(rawItem, parsed);
  }

  function analyzeKoreanGoodsTitle(rawItem) {
    const parsed = baseParsed(rawItem, 'korean_goods');
    parsed.language = getField(rawItem, ['言語', 'language']) || '韓国';
    const krGoods = stripKrGoodsNoise(stripCommonNoise(parsed.rawTitle));
    parsed.extractedWorkTitle = krGoods.title;
    parsed.goodsType = krGoods.goodsType;
    parsed.optionType = krGoods.optionType;
    parsed.confidence = krGoods.goodsType ? 0.86 : 0.55;
    return finalizeParsed(rawItem, parsed);
  }

  function analyzeAladinTitle(rawItem) {
    const parsed = baseParsed(rawItem, 'aladin');
    parsed.language = getField(rawItem, ['言語', 'language']) || '韓国';

    const magazine = extractMagazineInfo(rawItem);
    Object.assign(parsed, magazine);

    let cleaned = stripCommonNoise(parsed.rawTitle);
    const krGoods = stripKrGoodsNoise(cleaned);
    if (krGoods.goodsType) {
      parsed.extractedWorkTitle = krGoods.title;
      parsed.goodsType = krGoods.goodsType;
      parsed.optionType = krGoods.optionType;
      parsed.confidence = 0.86;
      return finalizeParsed(rawItem, parsed);
    }

    const split = splitSubtitle(cleaned);
    parsed.extractedWorkTitle = split.workTitle;
    parsed.subtitle = split.subtitle;
    parsed.confidence = split.subtitle || parsed.volume || parsed.edition || parsed.magazineName ? 0.82 : 0.65;
    return finalizeParsed(rawItem, parsed);
  }

  function analyzeProductTitle(rawItem) {
    try {
      const source = detectSource(rawItem);
      if (source === 'books_tw') return analyzeBooksTwTitle(rawItem);
      if (source === 'aladin') return analyzeAladinTitle(rawItem);
      if (source === 'korean_goods') return analyzeKoreanGoodsTitle(rawItem);

      const parsed = baseParsed(rawItem, source);
      const krGoods = /[가-힣]/.test(parsed.rawTitle) ? stripKrGoodsNoise(stripCommonNoise(parsed.rawTitle)) : null;
      const twGoods = stripTwGoodsNoise(stripCommonNoise(parsed.rawTitle));
      if (krGoods?.goodsType) {
        parsed.source = 'korean_goods';
        parsed.extractedWorkTitle = krGoods.title;
        parsed.goodsType = krGoods.goodsType;
        parsed.optionType = krGoods.optionType;
      } else if (twGoods.goodsType) {
        parsed.extractedWorkTitle = twGoods.title;
        parsed.goodsType = twGoods.goodsType;
        parsed.optionType = twGoods.optionType;
      } else {
        const split = splitSubtitle(stripCommonNoise(parsed.rawTitle));
        parsed.extractedWorkTitle = split.workTitle;
        parsed.subtitle = split.subtitle;
      }
      return finalizeParsed(rawItem, parsed);
    } catch (error) {
      const rawTitle = compactSpaces(getRawTitle(rawItem));
      return {
        source: detectSource(rawItem),
        rawTitle,
        originalTitle: rawTitle,
        originalProductTitle: rawTitle,
        extractedWorkTitle: rawTitle,
        subtitle: '',
        normalizedSearchTitle: rawTitle,
        itemType: 'unknown',
        language: getField(rawItem, ['言語', 'language']),
        category: getField(rawItem, ['categoryName', 'カテゴリ', 'ジャンル', 'mallType']),
        author: getField(rawItem, ['author', '著者', '作者']),
        volume: '',
        edition: '',
        bonusInfo: '',
        goodsType: '',
        optionType: '',
        magazineName: '',
        year: '',
        month: '',
        issue: '',
        confidence: 0.2,
        warnings: [`title analysis failed: ${error?.message || error}`],
        providerHint: buildProviderHint('unknown')
      };
    }
  }

  async function scrapeCurrentPageBySource(source) {
    const normalizedSource = normalizeSourceName(source) || source || 'unknown';
    if (normalizedSource === 'books_tw' && typeof global.requestActiveTabProductInfo === 'function') {
      return global.requestActiveTabProductInfo();
    }
    if (normalizedSource === 'aladin' && typeof global.collectCurrentPageProduct === 'function') {
      const result = await global.collectCurrentPageProduct();
      return result?.product || result;
    }
    throw new Error(`scrapeCurrentPageBySource unsupported source: ${normalizedSource}`);
  }

  function buildLookupPayload(rawItem, titleAnalysis) {
    const analysis = titleAnalysis || analyzeProductTitle(rawItem || {});
    return {
      action: 'lookupJapaneseTitle',
      source: analysis.source || detectSource(rawItem || {}),
      titleAnalysis: analysis
    };
  }

  // ===== 漢字重なり検証（誤マッチ防止）=====
  function _collectHanChars(text) {
    const m = String(text || '').match(/[㐀-䶿一-鿿]/g);
    return m || [];
  }

  function _hasJapaneseKana(text) {
    return /[ぁ-ゖァ-ヺ]/.test(String(text || ''));
  }

  const HAN_OVERLAP_EQUIV = {
    與: '与', 与: '與',
    戀: '恋', 恋: '戀',
    為: '为', 为: '為',
    臺: '台', 台: '臺',
    國: '国', 国: '國',
    學: '学', 学: '學',
    體: '体', 体: '體',
    廣: '广', 广: '廣',
    樂: '乐', 乐: '樂',
    電: '电', 电: '電',
    視: '视', 视: '視',
    劇: '剧', 剧: '劇',
    場: '场', 场: '場',
    畫: '画', 画: '畫',
    書: '书', 书: '書',
    來: '来', 来: '來',
    發: '发', 发: '發',
    網: '网', 网: '網',
    龍: '龙', 龙: '龍',
    馬: '马', 马: '馬',
    車: '车', 车: '車',
    門: '门', 门: '門',
    開: '开', 开: '開',
    關: '关', 关: '關',
    風: '风', 风: '風',
    氣: '气', 气: '氣',
    愛: '爱', 爱: '愛',
    見: '见', 见: '見',
    說: '说', 说: '說',
    話: '话', 话: '話',
    語: '语', 语: '語',
    讀: '读', 读: '讀',
    聲: '声', 声: '聲',
    頭: '头', 头: '頭',
    臉: '脸', 脸: '臉',
    髮: '发', 髪: '发',
    總: '总', 总: '總',
    無: '无', 无: '無',
    時: '时', 时: '時',
    間: '间', 间: '間',
    東: '东', 东: '東',
    絲: '丝', 丝: '絲',
    縣: '县', 县: '縣',
    區: '区', 区: '區',
    醫: '医', 医: '醫',
    藥: '药', 药: '藥',
    買: '买', 买: '買',
    賣: '卖', 卖: '賣',
    價: '价', 价: '價',
    錢: '钱', 钱: '錢',
    銀: '银', 银: '銀',
    鐵: '铁', 铁: '鐵',
    銅: '铜', 铜: '銅',
    點: '点', 点: '點',
    線: '线', 线: '線',
    機: '机', 机: '機',
    飛: '飞', 飞: '飛',
    鳥: '鸟', 鸟: '鳥',
    魚: '鱼', 鱼: '魚',
    島: '岛', 岛: '島',
    灣: '湾', 湾: '灣',
    鄉: '乡', 乡: '鄉',
    鎮: '镇', 镇: '鎮',
    縣: '县',
    護: '护', 护: '護',
    擊: '击', 击: '擊',
    戰: '战', 战: '戰',
    勝: '胜', 胜: '勝',
    負: '负', 负: '負',
    號: '号', 号: '號',
    碼: '码', 码: '碼',
    錄: '录', 录: '錄',
    製: '制', 制: '製',
    復: '复', 复: '復',
    雜: '杂', 杂: '雜',
    誌: '志', 志: '誌',
    專: '专', 专: '專',
    業: '业', 业: '業',
    產: '产', 产: '產',
    質: '质', 质: '質',
    類: '类', 类: '類',
    極: '极', 极: '極',
    樂: '乐',
    歲: '岁', 岁: '歲',
    華: '华', 华: '華',
    聖: '圣', 圣: '聖',
    靈: '灵', 灵: '靈',
    龍: '龙',
    術: '术', 术: '術',
    師: '师', 师: '師',
    殺: '杀', 杀: '殺',
    惡: '恶', 恶: '惡',
    魔: '魔',
    獸: '兽', 兽: '獸',
    靈: '灵',
  };

  function _normalizeHanForOverlap_(text) {
    return String(text || '').split('').map(ch => HAN_OVERLAP_EQUIV[ch] || ch).join('');
  }

  function _hanOverlapCount(a, b) {
    const aNorm = _normalizeHanForOverlap_(a);
    const bNorm = _normalizeHanForOverlap_(b);
    const aChars = _collectHanChars(aNorm);
    const bChars = _collectHanChars(bNorm);
    if (!aChars.length || !bChars.length) return 0;
    const setA = new Set(aChars);
    const seenB = new Set();
    let hit = 0;
    for (const ch of bChars) {
      if (seenB.has(ch)) continue;
      seenB.add(ch);
      if (setA.has(ch)) hit += 1;
    }
    return hit;
  }

  /**
   * GAS等から取得した日本語タイトル候補が、元のクエリと整合しているか検証する。
   * 全球高考 → 週刊ダイヤモンド のような漢字ゼロ重なりの誤マッチを弾く。
   * @returns {boolean} 採用してよいなら true
   */
  function validateJapaneseTitleAgainstQuery_(jpTitle, originalQueries, provider) {
    const jp = compactSpaces(jpTitle);
    if (!jp) return false;
    // 信頼度の高いプロバイダーは検証スキップ（辞書一致など）
    const trustedProviders = new Set(['titleAliasDictionary', '自社タイトル別名辞書', 'chilchil']);
    if (trustedProviders.has(String(provider || ''))) return true;
    // クエリから漢字を集める
    const queries = Array.isArray(originalQueries) ? originalQueries : [originalQueries];
    const queryHan = [];
    for (const q of queries) {
      queryHan.push(..._collectHanChars(q));
    }
    // 漢字クエリでない場合は検証スキップ（ローマ字/ハングル等）
    if (!queryHan.length) return true;
    if (!_hasJapaneseKana(jp) && !_collectHanChars(jp).length) return false;
    const maxOverlap = Math.max(
      0,
      ...queries.map(q => _hanOverlapCount(q, jp))
    );
    const minOverlap = Math.min(2, Math.max(1, Math.floor(queryHan.length / 3)));
    if (maxOverlap >= minOverlap) return true;
    // 漢字クエリに対し、日本語候補がカナのみで漢字重なりゼロ → 別作品誤マッチ（例: バトルファッカーB子）
    if (!_collectHanChars(jp).length && _hasJapaneseKana(jp)) return false;
    return false;
  }

  function stripMuVolumeNoise_(value) {
    let s = compactSpaces(value);
    const rules = [
      /[\s　]*(?:第)?\s*[0-9０-９]{1,4}\s*(?:漫畫|漫画|コミック|單行本|单行本)\s*$/i,
      /[\s　]*(?:漫畫|漫画|コミック|單行本|单行本)\s*(?:第)?\s*[0-9０-９]{1,4}\s*$/i,
      /[\s　]*[#＃]?\s*[0-9０-９]{1,4}\s*$/i,
      /[\s　]*(?:vol\.?|v\.?|第)\s*[0-9０-９]{1,4}\s*(?:巻|卷|集|冊|話|回)?\s*$/i,
      /[\s　]*(?:巻|卷|集|冊|話|回)\s*[0-9０-９]{1,4}\s*$/i,
      /[\s　]*(?:漫畫|漫画|コミック|單行本|单行本)\s*$/i,
    ];
    for (let i = 0; i < 4; i += 1) {
      const prev = s;
      for (const re of rules) {
        s = s.replace(re, '').trim();
      }
      if (s === prev) break;
    }
    return s;
  }

  function expandChineseScriptVariants_(text) {
    const base = compactSpaces(text);
    if (!base) return [];
    const out = [base];
    let variant = '';
    for (const ch of base) {
      variant += HAN_OVERLAP_EQUIV[ch] || ch;
    }
    variant = compactSpaces(variant);
    if (variant && variant !== base) out.push(variant);
    return out;
  }

  function buildMuQueryVariants_(titleAnalysis, rawItem) {
    const item = rawItem || {};
    const seeds = uniq([
      titleAnalysis?.normalizedSearchTitle,
      titleAnalysis?.extractedWorkTitle,
      titleAnalysis?.originalTitle,
      titleAnalysis?.originalProductTitle,
      item['原題タイトル'],
      item['原題商品タイトル'],
      item['原題商品名'],
      item['商品名'],
      item.title,
    ].map(compactSpaces).filter(Boolean));
    const expanded = [];
    for (const seed of seeds) {
      const stripped = stripMuVolumeNoise_(seed);
      const noNumber = compactSpaces(stripped).replace(/[0-9０-９]+/g, '').trim();
      expanded.push(seed, stripped, noNumber);
      expanded.push(...expandChineseScriptVariants_(stripped || seed));
    }
    return uniq(expanded.filter(title => title.length >= 2));
  }

  async function invokeMangaUpdatesLookupDetailed_(queries) {
    const client = globalThis.titleLookupMangaUpdates;
    if (client && typeof client.lookupJapaneseTitleDetailed === 'function') {
      return client.lookupJapaneseTitleDetailed(queries);
    }
    if (typeof chrome === 'undefined' || !chrome.runtime?.sendMessage) return null;
    const response = await new Promise((resolve, reject) => {
      chrome.runtime.sendMessage(
        { action: 'lookupMangaUpdatesJapaneseTitleDirect', queries },
        (res) => {
          if (chrome.runtime.lastError) {
            reject(new Error(chrome.runtime.lastError.message));
            return;
          }
          resolve(res);
        }
      );
    });
    if (response?.ok && response.result) return response.result;
    if (response?.error) throw new Error(response.error);
    return null;
  }

  async function invokeMangaUpdatesLookupJapanese_(queries) {
    const client = globalThis.titleLookupMangaUpdates;
    if (client && typeof client.lookupJapaneseTitle === 'function') {
      return client.lookupJapaneseTitle(queries);
    }
    const detailed = await invokeMangaUpdatesLookupDetailed_(queries);
    return detailed?.japaneseTitle ? String(detailed.japaneseTitle).trim() : '';
  }

  async function tryMangaUpdatesDirectJapanese_(titleAnalysis, rawItem) {
    const itemType = toText(titleAnalysis?.itemType);
    const tryTypes = new Set(['manga', 'bl_manga', 'goods', 'light_novel', 'unknown']);
    if (!tryTypes.has(itemType)) {
      return { japaneseTitle: '', muTrace: `mangaupdatesClient:skip(itemType=${itemType || 'empty'})`, muStatus: 'skipped', matchedTitles: [] };
    }
    const queries = buildMuQueryVariants_(titleAnalysis, rawItem || {});
    if (!queries.length) {
      return { japaneseTitle: '', muTrace: 'mangaupdatesClient:skip(empty_query)', muStatus: 'skipped', matchedTitles: [] };
    }
    const traceQ = queries.length > 1 ? `${queries[0]}=>${queries[1]}` : queries[0];
    try {
      const detailed = await invokeMangaUpdatesLookupDetailed_(queries);
      if (detailed) {
        const jpRaw = compactSpaces(detailed?.japaneseTitle || '');
        const jp = jpRaw && validateJapaneseTitleAgainstQuery_(jpRaw, queries, 'mangaUpdates(extension)')
          ? jpRaw
          : '';
        const status = String(detailed?.status || '');
        const matchedTitles = Array.isArray(detailed?.matchedTitles) ? detailed.matchedTitles : [];
        if (jp) {
          return { japaneseTitle: jp, muTrace: `mangaupdatesClient:hit(q=${traceQ})`, muStatus: 'resolved', matchedTitles };
        }
        if (status === 'series_found_no_japanese' && matchedTitles.length) {
          const head = matchedTitles.slice(0, 4).join('/');
          return {
            japaneseTitle: '',
            muTrace: `mangaupdatesClient:series_found_no_jp(q=${traceQ} titles=${head})`,
            muStatus: 'series_found_no_japanese',
            matchedTitles,
          };
        }
        return { japaneseTitle: '', muTrace: `mangaupdatesClient:miss(q=${traceQ})`, muStatus: 'not_found', matchedTitles };
      }
      const jpRaw = compactSpaces(await invokeMangaUpdatesLookupJapanese_(queries));
      const jp = jpRaw && validateJapaneseTitleAgainstQuery_(jpRaw, queries, 'mangaUpdates(extension)')
        ? jpRaw
        : '';
      if (jp) return { japaneseTitle: jp, muTrace: `mangaupdatesClient:hit(q=${traceQ})`, muStatus: 'resolved', matchedTitles: [] };
      return { japaneseTitle: '', muTrace: 'mangaupdatesClient:skip(no_client)', muStatus: 'skipped', matchedTitles: [] };
    } catch (error) {
      const msg = error?.message || String(error || 'error');
      return { japaneseTitle: '', muTrace: `mangaupdatesClient:error(${msg})`, muStatus: 'error', matchedTitles: [] };
    }
  }

  async function enrichLookupWithDirectMangaUpdates_(titleAnalysis, data) {
    const lookup = data?.lookup;
    if (!lookup || typeof lookup !== 'object') return data;
    if (compactSpaces(lookup.japaneseTitle || lookup.title || '')) return data;
    const itemType = toText(titleAnalysis?.itemType);
    const tryTypes = new Set(['manga', 'bl_manga', 'goods', 'light_novel', 'unknown']);
    if (!tryTypes.has(itemType)) return data;
    const baseQueries = buildMuQueryVariants_(titleAnalysis, data?.rawItem || {});
    if (!baseQueries.length) return data;
    const q = baseQueries[0];
    const aliasCandidates = Array.isArray(lookup.candidates) ? lookup.candidates : [];
    const queries = uniq(baseQueries.concat(aliasCandidates.map((c) => compactSpaces(c)).filter(Boolean)));
    const hasJapaneseSignal = (s) => /[ぁ-ゖァ-ヺーゝゞ々〆〇]/.test(String(s || ''));
    try {
      const jpRaw = compactSpaces(await invokeMangaUpdatesLookupJapanese_(queries));
      const jp = jpRaw && validateJapaneseTitleAgainstQuery_(jpRaw, baseQueries, 'mangaUpdates(extension)')
        ? jpRaw
        : '';
      if (!jp) {
        const traceQ = queries.length > 1 ? `${q}+${queries.length - 1}cand` : q;
        const jpFromCandidates = aliasCandidates.map((c) => compactSpaces(c)).find((c) => {
          if (!hasJapaneseSignal(c)) return false;
          return validateJapaneseTitleAgainstQuery_(c, baseQueries, 'mangaUpdates(candidate)');
        }) || '';
        if (jpFromCandidates) {
          return {
            ...data,
            lookup: {
              ...lookup,
              status: 'resolved',
              japaneseTitle: jpFromCandidates,
              provider: lookup.provider ? `${lookup.provider}+candidate` : 'cascade(candidate)',
              trace: `${lookup.trace || lookup.log || ''} | mangaupdatesClient:retry_miss(q=${traceQ}) | candidate:fallback_hit`.replace(/^\s*\|\s*/, '').trim(),
            },
          };
        }
        return {
          ...data,
          lookup: {
            ...lookup,
            trace: `${lookup.trace || lookup.log || ''} | mangaupdatesClient:retry_miss(q=${traceQ})`.replace(/^\s*\|\s*/, '').trim(),
          },
        };
      }
      return {
        ...data,
        lookup: {
          ...lookup,
          status: 'resolved',
          japaneseTitle: jp,
          provider: lookup.provider ? `${lookup.provider}+mangaupdatesClient` : 'mangaUpdates(extension)',
          trace: `${lookup.trace || lookup.log || ''} | mangaupdates:direct_hit`.replace(/^\s*\|\s*/, '').trim(),
        },
      };
    } catch (error) {
      const msg = error?.message || String(error || 'error');
      const baseLookup = lookup && typeof lookup === 'object' ? lookup : {};
      return {
        ...data,
        lookup: {
          ...baseLookup,
          trace: `${baseLookup.trace || ''} | mangaupdatesClient:retry_error(${msg})`.replace(/^\s*\|\s*/, '').trim(),
        },
      };
    }
  }

  async function lookupJapaneseTitle(titleAnalysis, options = {}) {
    const gasUrl = toText(options.url || options.gasUrl || global.gasWebAppUrl || '');
    if (!/^https:\/\/(?:script\.google\.com|script\.googleusercontent\.com)\//i.test(gasUrl)) {
      throw new Error('GAS Webアプリ URL が未設定です');
    }

    const payload = buildLookupPayload(options.rawItem || {}, titleAnalysis);
    const gasPromise = fetch(gasUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload),
      credentials: 'omit',
      redirect: 'follow'
    }).then(async response => {
      if (!response.ok) throw new Error(`GAS lookup HTTP ${response.status}`);
      const data = await response.json().catch(() => null);
      if (!data || typeof data !== 'object') throw new Error('GAS lookup JSON応答を確認できませんでした');
      if (data.success === false || data.ok === false) {
        throw new Error(data.error || '日本語タイトル照会に失敗しました');
      }
      return data;
    });
    const muPromise = tryMangaUpdatesDirectJapanese_(titleAnalysis, options.rawItem || {});
    const [data, muPack] = await Promise.all([gasPromise, muPromise]);
    const muEarly = muPack && typeof muPack === 'object' ? muPack.japaneseTitle : '';
    const muTraceLine = muPack && typeof muPack === 'object' ? muPack.muTrace : '';
    const muStatus = muPack && typeof muPack === 'object' ? String(muPack.muStatus || '') : '';
    const muMatchedTitles = muPack && typeof muPack === 'object' && Array.isArray(muPack.matchedTitles)
      ? muPack.matchedTitles
      : [];
    let merged = data;
    function mergeLookupTrace(lookup, extra) {
      if (!extra) return lookup;
      const base = lookup && typeof lookup === 'object' ? { ...lookup } : {};
      const nextTrace = [base.trace, base.log, extra].filter(Boolean).join(' | ').replace(/^\s*\|\s*/, '').trim();
      return { ...base, trace: nextTrace };
    }
    if (merged && merged.lookup) {
      merged = { ...merged, lookup: mergeLookupTrace(merged.lookup, muTraceLine) };
    } else if (muTraceLine) {
      merged = { ...merged, lookup: mergeLookupTrace(null, muTraceLine) };
    }

    // === GAS結果の検証: 漢字重なりゼロ等の誤マッチを検知 ===
    const originalQueries = buildMuQueryVariants_(titleAnalysis, options.rawItem || {});
    const gasLookup = merged?.lookup || {};
    const gasJp = compactSpaces(gasLookup.japaneseTitle || gasLookup.title || '');
    const gasProvider = String(gasLookup.provider || '');
    if (gasJp && !validateJapaneseTitleAgainstQuery_(gasJp, originalQueries, gasProvider)) {
      // 誤マッチと判定: タイトルをクリアして再検索を許可する
      const baseLookup = merged.lookup && typeof merged.lookup === 'object' ? merged.lookup : {};
      merged = {
        ...merged,
        lookup: {
          ...baseLookup,
          status: 'not_found',
          japaneseTitle: '',
          provider: gasProvider ? `${gasProvider}(rejected)` : '',
          // candidates は残す（後段で参考に使う）
          trace: `${baseLookup.trace || ''} | validate:rejected(${gasProvider || 'unknown'}:${gasJp})`.replace(/^\s*\|\s*/, '').trim(),
        },
      };
    }

    if (muEarly && !compactSpaces(merged?.lookup?.japaneseTitle || merged?.lookup?.title || '')) {
      const baseLookup = merged.lookup && typeof merged.lookup === 'object' ? merged.lookup : {};
      if (validateJapaneseTitleAgainstQuery_(muEarly, originalQueries, 'mangaUpdates(extension)')) {
        merged = {
          ...merged,
          lookup: {
            ...baseLookup,
            status: 'resolved',
            japaneseTitle: muEarly,
            provider: baseLookup.provider ? `${baseLookup.provider}+mangaupdatesClient` : 'mangaUpdates(extension)',
            trace: `${baseLookup.trace || ''} | mangaupdates:parallel_direct`.replace(/^\s*\|\s*/, '').trim(),
          },
        };
      }
    }
    merged = await enrichLookupWithDirectMangaUpdates_(titleAnalysis, merged);

    // MUに作品自体は登録されているが日本語タイトルが無い場合の状態を反映
    const finalLookup = merged?.lookup || {};
    const finalJp = compactSpaces(finalLookup.japaneseTitle || finalLookup.title || '');
    if (!finalJp && muStatus === 'series_found_no_japanese' && muMatchedTitles.length) {
      const baseLookup = merged.lookup && typeof merged.lookup === 'object' ? merged.lookup : {};
      const existingCandidates = Array.isArray(baseLookup.candidates) ? baseLookup.candidates : [];
      const mergedCandidates = uniq([...existingCandidates, ...muMatchedTitles]);
      merged = {
        ...merged,
        lookup: {
          ...baseLookup,
          status: 'series_found_no_japanese',
          provider: baseLookup.provider
            ? `${baseLookup.provider}+mangaupdates(series_no_jp)`
            : 'mangaUpdates(series_no_jp)',
          candidates: mergedCandidates,
          trace: `${baseLookup.trace || ''} | mangaupdates:series_registered_no_jp`.replace(/^\s*\|\s*/, '').trim(),
        },
      };
    }
    return merged;
  }

  function buildPayloadForGas(rawItem) {
    const payload = { ...(rawItem || {}) };
    payload.titleAnalysis = analyzeProductTitle(rawItem || {});
    return payload;
  }

  function buildSavePayload(rawItem, titleAnalysis, japaneseTitleLookup) {
    const analysis = titleAnalysis || analyzeProductTitle(rawItem || {});
    return {
      action: 'upsertProductWithLookup',
      source: analysis.source || detectSource(rawItem || {}),
      rawItem: rawItem || {},
      titleAnalysis: analysis,
      japaneseTitleLookup: japaneseTitleLookup || null
    };
  }

  async function saveProductWithLookup(payload, options = {}) {
    const gasUrl = toText(options.url || options.gasUrl || global.gasWebAppUrl || '');
    if (!/^https:\/\/(?:script\.google\.com|script\.googleusercontent\.com)\//i.test(gasUrl)) {
      throw new Error('GAS Webアプリ URL が未設定です');
    }
    const response = await fetch(gasUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload),
      credentials: 'omit',
      redirect: 'follow'
    });
    if (!response.ok) throw new Error(`GAS save HTTP ${response.status}`);
    const data = await response.json().catch(() => null);
    if (!data || typeof data !== 'object') throw new Error('GAS save JSON応答を確認できませんでした');
    if (data.success === false || data.ok === false) throw new Error(data.error || '保存に失敗しました');
    return data;
  }

  function isJapaneseTitleStatusPlaceholder_(value) {
    const t = compactSpaces(value);
    if (!t) return true;
    return /^(?:登録なし|照会失敗|未照会)(?:\s*[（(]|$)/u.test(t);
  }

  /** シート「原題タイトル」向け: 巻・コミック語尾など（GAS stripLookupVolumeNoise_ 相当） */
  function stripVolumeNoiseForSheetOriginal(value) {
    let s = compactSpaces(value);
    if (!s) return '';
    const rules = [
      /[\s　]*(?:vol\.?|v\.?|第)\s*[0-9０-９]{1,4}\s*(?:巻|卷|集|冊|話|回|期|号|號)?\s*$/iu,
      /[\s　]*[#＃]?\s*[0-9０-９]{1,4}\s*$/u,
      /[\s　]*(?:第)?\s*[0-9０-９]{1,4}\s*(?:漫畫|漫画|コミック|單行本|单行本)\s*$/iu,
      /[\s　]*(?:漫畫|漫画|コミック|單行本|单行本)\s*(?:第)?\s*[0-9０-９]{1,4}\s*$/iu,
      /[\s　]*(?:漫畫|漫画|コミック|單行本|单行本)\s*$/iu,
      /[\s　]+(?:資料夾|资料夹|文件夾|文件夹)[\s　]*[0-9０-９]{0,4}[\s　]*入\s*$/iu
    ];
    for (let i = 0; i < 4; i++) {
      const prev = s;
      for (const re of rules) {
        s = String(s || '').replace(re, '').trim();
      }
      s = compactSpaces(s);
      if (s === prev) break;
    }
    return s;
  }

  /** シート「日本語タイトル」向け: (TV) 等メディア表記 + 末尾の巻数を除去（親シリーズ・作品名は維持） */
  function normalizeSheetJapaneseWorkTitle(value) {
    if (isJapaneseTitleStatusPlaceholder_(value)) return compactSpaces(value);
    let s = compactSpaces(value);
    if (!s) return '';
    for (let i = 0; i < 8; i++) {
      const prev = s;
      s = s
        .replace(/\s*[\[(（〈【]\s*(?:TV|T\.V\.|OVA|OAD|ONA|SP|映画|電影|劇場版|剧场版|アニメ(?:版|ーション)?|ドラマ(?:版)?|真人版|Web\s*アニメ|Anime|Movie|Film)\s*[\])）〉】]/giu, '')
        .replace(/\s*[\[\(]\s*[Tt][Vv]\s*[\]\)]\s*$/u, '')
        // 末尾の巻数表記: (4) / 第4巻 / 4巻（巻数は単巻数列が保持するため作品名から落とす）
        .replace(/\s*[（(]\s*[0-9０-９]{1,4}\s*[）)]\s*$/u, '')
        .replace(/\s*第?\s*[0-9０-９]{1,4}\s*[巻卷]\s*$/u, '')
        .trim();
      s = compactSpaces(s);
      if (s === prev) break;
    }
    return s;
  }

  function applyJapaneseTitleLookupToProduct(rawItem, lookupResult) {
    const product = { ...(rawItem || {}) };
    const lookup = lookupResult?.lookup || lookupResult || null;
    if (!lookup || typeof lookup !== 'object') return product;

    product.japaneseTitleLookup = lookup;
    const japaneseTitle = normalizeSheetJapaneseWorkTitle(compactSpaces(lookup.japaneseTitle || lookup.title || ''));
    if (lookup.status === 'resolved' && japaneseTitle) {
      product.日本語タイトル = japaneseTitle;
      product['作品名（日本語）'] = product['作品名（日本語）'] || japaneseTitle;
      product.作品名日本語 = product.作品名日本語 || japaneseTitle;
      product['商品名（日本語）'] = product['商品名（日本語）'] || japaneseTitle;
      product['商品名(日本語)'] = product['商品名(日本語)'] || japaneseTitle;
    }
    return product;
  }

  function renderLookupResult(result) {
    const lookup = result?.lookup || null;
    if (!lookup) {
      return result?.success || result?.ok
        ? `保存成功${result?.sheet ? ` / ${result.sheet}` : ''}${result?.row ? ` 行${result.row}` : ''}`
        : `保存失敗${result?.error ? `: ${result.error}` : ''}`;
    }
    const titleValue = lookup.japaneseTitle || lookup.title || '';
    const status = lookup.status || (titleValue ? 'resolved' : 'not_found');
    const statusLabel = {
      resolved: '解決',
      not_found: '登録なし(全サイト)',
      partial_error: '一部照会失敗',
      failed: '照会失敗',
      skipped: '未照会'
    }[status] || status;
    const title = titleValue ? ` / ${titleValue}` : '';
    const provider = lookup.provider ? ` / ${lookup.provider}` : '';
    const log = lookup.trace || lookup.log ? ` / ${lookup.trace || lookup.log}` : '';
    return `${statusLabel}${title}${provider}${log}`;
  }

  function renderSaveResult(result) {
    if (!result || typeof result !== 'object') return '保存失敗: 応答なし';
    const ok = result.success !== false && result.ok !== false;
    const location = [
      result.sheet ? `保存先:${result.sheet}` : '',
      result.row ? `行:${result.row}` : '',
      result.mode ? `mode:${result.mode}` : ''
    ].filter(Boolean).join(' / ');
    const lookupText = result.lookup ? ` / 照会:${renderLookupResult(result)}` : '';
    return `${ok ? '保存成功' : '保存失敗'}${location ? ` / ${location}` : ''}${lookupText}${result.error ? ` / ${result.error}` : ''}`;
  }

  const api = {
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
    stripTwGoodsNoise,
    stripKrGoodsNoise,
    stripSearchNoise,
    buildNormalizedSearchTitle,
    detectSourceFromUrl,
    scrapeCurrentPageBySource,
    buildLookupPayload,
    lookupJapaneseTitle,
    buildPayloadForGas,
    buildSavePayload,
    saveProductWithLookup,
    applyJapaneseTitleLookupToProduct,
    validateJapaneseTitleAgainstQuery: validateJapaneseTitleAgainstQuery_,
    buildMuQueryVariants: buildMuQueryVariants_,
    stripVolumeNoiseForSheetOriginal,
    normalizeSheetJapaneseWorkTitle,
    renderLookupResult,
    renderSaveResult
  };

  global.TitleAnalysis = api;
  for (const [name, fn] of Object.entries(api)) {
    if (typeof global[name] !== 'function') global[name] = fn;
  }
})(globalThis);
