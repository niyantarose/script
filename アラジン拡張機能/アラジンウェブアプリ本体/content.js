(() => {
  const VERSION = '2026-03-11-a';
  if (globalThis.__ALADIN_SCRAPER__?.version === VERSION) {
    return;
  }

  const SHARED_DESCRIPTION_SELECTORS = [
    '.pContent[id$="_Introduce"]',
    '#bookDescriptionToggle',
    '#productDescription',
    '.Ere_prod_mconts_box',
    '.Ere_prod_mconts',
    'div[class*="mconts"]'
  ];

  const SHARED_IMAGE_SELECTORS = [
    '#swiper-container-cover .swiper-slide img',
    '.pContent[id$="_Introduce"] img',
    '.Ere_prod_mconts_box img',
    '#bookDescriptionToggle img',
    '#productImageContainer img',
    '#productDescription img',
    '.detail_img img',
    '.Ere_prod_mconts img',
    'div[class*="mconts"] img',
    '.conts_info_list1 img',
    '.option_image img'
  ];

  const SHARED_BASIC_INFO_SELECTORS = [
    '.Ere_prod_mconts_box .conts_info_list1',
    '.Ere_prod_info_list',
    '#productBasicInfo',
    '.basic_info',
    '.info_list',
    '.conts_info_list1'
  ];

  const SHARED_IMAGE_EXCLUDE_ANCESTORS = [
    '#swiper_itemEvent',
    '.conts_info_list6',
    '.Ere_music_chart2',
    '.np_this_event_td',
    '.event_nav',
    '#np_series',
    '.series_wrap',
    '.seriesStartAnchor',
    '.seriesEndAnchor'
  ];

  const DEFAULT_EXCLUDE_SECTION_LABELS = [
    /이벤트/i,
    /기본정보/i,
    /시리즈/i,
    /주제 분류/i,
    /관련상품/i
  ];

  const DESCRIPTION_EXCLUDE_SECTION_LABELS = [
    ...DEFAULT_EXCLUDE_SECTION_LABELS,
    /목차/i,
    /트랙/i,
    /수록곡/i,
    /disc/i,
    /contents/i
  ];

  const SCRAPER_RULES = {
    書籍: {
      descriptionSelectors: [
        '.pContent[id$="_Introduce"]',
        '#bookDescriptionToggle',
        '#productDescription',
        '.Ere_prod_mconts_box'
      ],
      descriptionPlans: [
        {
          selectors: ['.pContent[id$="_Introduce"]'],
          minLength: 20,
          maxLength: 2000
        },
        {
          selectors: ['#bookDescriptionToggle', '#productDescription', '.Ere_prod_mconts_box'],
          includeSectionLabels: [/소개/i, /책소개/i, /출판사/i],
          excludeSectionLabels: DESCRIPTION_EXCLUDE_SECTION_LABELS,
          minLength: 20,
          maxLength: 2000
        }
      ],
      basicInfoSelectors: [
        '.Ere_prod_info_list',
        '#productBasicInfo',
        '.basic_info',
        '.conts_info_list1'
      ],
      waitSelectors: ['.pContent[id$="_Introduce"] img'],
      imagePlans: [
        {
          kind: 'coverSlides',
          selectors: ['#swiper-container-cover .swiper-slide img'],
          allowPathPatterns: [/\/letslook\//i],
          preferLast: true,
          maxCount: 1
        },
        {
          kind: 'introduce',
          selectors: ['.pContent[id$="_Introduce"] img'],
          maxCount: 2
        },
        {
          kind: 'section',
          selectors: ['#bookDescriptionToggle img', '#productDescription img', '.Ere_prod_mconts_box img'],
          includeSectionLabels: [/소개/i, /책소개/i, /출판사/i],
          excludeSectionLabels: DEFAULT_EXCLUDE_SECTION_LABELS,
          maxCount: 2
        }
      ],
      excludeAncestorSelectors: SHARED_IMAGE_EXCLUDE_ANCESTORS,
      excludeSectionLabels: DEFAULT_EXCLUDE_SECTION_LABELS,
      allowGenericFallback: true,
      fallbackMaxCount: 3
    },
    マンガ: {
      descriptionSelectors: [
        '.pContent[id$="_Introduce"]',
        '#bookDescriptionToggle',
        '#productDescription',
        '.Ere_prod_mconts_box'
      ],
      descriptionPlans: [
        {
          selectors: ['.pContent[id$="_Introduce"]'],
          minLength: 20,
          maxLength: 2000
        },
        {
          selectors: ['#bookDescriptionToggle', '#productDescription', '.Ere_prod_mconts_box'],
          includeSectionLabels: [/소개/i, /책소개/i, /출판사/i],
          excludeSectionLabels: DESCRIPTION_EXCLUDE_SECTION_LABELS,
          minLength: 20,
          maxLength: 2000
        }
      ],
      basicInfoSelectors: [
        '.Ere_prod_mconts_box .conts_info_list1',
        '.conts_info_list1',
        '.Ere_prod_info_list',
        '#productBasicInfo',
        '.basic_info'
      ],
      waitSelectors: ['.Ere_prod_mconts_box .conts_info_list1', '.pContent[id$="_Introduce"]', '.pContent[id$="_Introduce"] img'],
      imagePlans: [
        {
          kind: 'coverViewer',
          selectors: ['#CoverMainImage', '#swiper-container-cover .swiper-slide img'],
          interactCarousel: true,
          clickSelectors: ['.prev_box [onclick]', '.prev_box button', '.prev_box [role="button"]', '.prev_box a[href^="javascript"]', '.swiper-button-next', '.cover [class*="next"]'],
          maxClicks: 8,
          clickDelayMs: 250,
          maxCount: 6
        },
        {
          kind: 'coverSlides',
          selectors: ['#swiper-container-cover .swiper-slide img'],
          allowPathPatterns: [/\/letslook\//i],
          preferLast: true,
          maxCount: 1
        },
        {
          kind: 'introduce',
          selectors: ['.pContent[id$="_Introduce"] img'],
          maxCount: 3
        },
        {
          kind: 'section',
          selectors: ['#bookDescriptionToggle img', '#productDescription img', '.Ere_prod_mconts_box img'],
          includeSectionLabels: [/소개/i, /책소개/i, /출판사/i],
          excludeSectionLabels: DEFAULT_EXCLUDE_SECTION_LABELS,
          maxCount: 2
        }
      ],
      excludeAncestorSelectors: SHARED_IMAGE_EXCLUDE_ANCESTORS,
      excludeSectionLabels: DEFAULT_EXCLUDE_SECTION_LABELS,
      allowGenericFallback: true,
      fallbackMaxCount: 2
    },
    音楽映像: {
      descriptionSelectors: [
        '.pContent[id$="_Introduce"]',
        '#productDescription',
        '.Ere_prod_mconts_box',
        '.Ere_prod_mconts'
      ],
      descriptionPlans: [
        {
          selectors: ['.pContent[id$="_Introduce"]'],
          minLength: 20,
          maxLength: 2000
        },
        {
          selectors: ['.Ere_prod_mconts_box', '#productDescription', '.Ere_prod_mconts'],
          includeSectionLabels: [/소개/i],
          excludeSectionLabels: DESCRIPTION_EXCLUDE_SECTION_LABELS,
          minLength: 20,
          maxLength: 2000
        }
      ],
      basicInfoSelectors: [
        '.Ere_prod_mconts_box .conts_info_list1',
        '.conts_info_list1',
        '.Ere_prod_info_list',
        '.basic_info',
        '.info_list'
      ],
      waitSelectors: ['.Ere_prod_mconts_box .conts_info_list1', '.pContent[id$="_Introduce"]', '.Ere_prod_mconts_box img', '.pContent[id$="_Introduce"] img'],
      imagePlans: [
        {
          kind: 'coverViewer',
          selectors: ['#CoverMainImage', '#swiper-container-cover .swiper-slide img'],
          interactCarousel: true,
          clickSelectors: ['.prev_box [onclick]', '.prev_box button', '.prev_box [role="button"]', '.prev_box a[href^="javascript"]', '.swiper-button-next', '.cover [class*="next"]'],
          maxClicks: 8,
          clickDelayMs: 250,
          maxCount: 8
        },
        {
          kind: 'section',
          selectors: ['.Ere_prod_mconts_box img', '#productDescription img', '.detail_img img', '.Ere_prod_mconts img'],
          includeSectionLabels: [/소개/i],
          excludeSectionLabels: DEFAULT_EXCLUDE_SECTION_LABELS,
          maxCount: 10
        },
        {
          kind: 'introduce',
          selectors: ['.pContent[id$="_Introduce"] img'],
          maxCount: 6
        }
      ],
      excludeAncestorSelectors: SHARED_IMAGE_EXCLUDE_ANCESTORS,
      excludeSectionLabels: DEFAULT_EXCLUDE_SECTION_LABELS,
      allowGenericFallback: true,
      fallbackMaxCount: 6
    },
    グッズ: {
      descriptionSelectors: [
        '.pContent[id$="_Introduce"]',
        '#productDescription',
        '.Ere_prod_mconts_box',
        '.Ere_prod_mconts'
      ],
      descriptionPlans: [
        {
          selectors: ['.pContent[id$="_Introduce"]'],
          minLength: 20,
          maxLength: 2000
        },
        {
          selectors: ['#productDescription', '.Ere_prod_mconts_box', '.Ere_prod_mconts'],
          includeSectionLabels: [/상품/i, /상세/i, /설명/i, /정보/i, /구성/i],
          excludeSectionLabels: DESCRIPTION_EXCLUDE_SECTION_LABELS,
          minLength: 20,
          maxLength: 2000
        }
      ],
      basicInfoSelectors: [
        '.info_list',
        '.Ere_prod_info_list',
        '.basic_info',
        '.conts_info_list1'
      ],
      waitSelectors: ['.option_image img', '.Ere_prod_mconts_box img', '.pContent[id$="_Introduce"] img'],
      imagePlans: [
        {
          kind: 'section',
          selectors: ['.option_image img', '.Ere_prod_mconts_box img', '#productDescription img', '#productImageContainer img', '.detail_img img'],
          includeSectionLabels: [/상품/i, /상세/i, /구성/i, /이미지/i, /설명/i, /사이즈/i, /정보/i],
          excludeSectionLabels: DEFAULT_EXCLUDE_SECTION_LABELS,
          maxCount: 12
        },
        {
          kind: 'introduce',
          selectors: ['.pContent[id$="_Introduce"] img'],
          maxCount: 4
        }
      ],
      excludeAncestorSelectors: SHARED_IMAGE_EXCLUDE_ANCESTORS,
      excludeSectionLabels: DEFAULT_EXCLUDE_SECTION_LABELS,
      allowGenericFallback: true,
      fallbackMaxCount: 8
    },
    default: {
      descriptionSelectors: SHARED_DESCRIPTION_SELECTORS,
      descriptionPlans: [
        {
          selectors: ['.pContent[id$="_Introduce"]'],
          minLength: 20,
          maxLength: 2000
        },
        {
          selectors: SHARED_DESCRIPTION_SELECTORS,
          excludeSectionLabels: DESCRIPTION_EXCLUDE_SECTION_LABELS,
          minLength: 20,
          maxLength: 2000
        }
      ],
      basicInfoSelectors: SHARED_BASIC_INFO_SELECTORS,
      waitSelectors: ['.pContent[id$="_Introduce"] img', '.Ere_prod_mconts_box img'],
      imagePlans: [
        {
          kind: 'coverSlides',
          selectors: ['#swiper-container-cover .swiper-slide img'],
          allowPathPatterns: [/\/letslook\//i],
          preferLast: true,
          maxCount: 1
        },
        {
          kind: 'introduce',
          selectors: ['.pContent[id$="_Introduce"] img'],
          maxCount: 2
        },
        {
          kind: 'section',
          selectors: ['.Ere_prod_mconts_box img', '#productDescription img', '.detail_img img'],
          excludeSectionLabels: DEFAULT_EXCLUDE_SECTION_LABELS,
          maxCount: 4
        }
      ],
      excludeAncestorSelectors: SHARED_IMAGE_EXCLUDE_ANCESTORS,
      excludeSectionLabels: DEFAULT_EXCLUDE_SECTION_LABELS,
      allowGenericFallback: true,
      fallbackMaxCount: 4
    }
  };

  function uniqueItems(...groups) {
    return [...new Set(groups.flat().filter(Boolean))];
  }

  function normalizeGenre(genre) {
    if (genre === '雑誌') {
      return '書籍';
    }
    return SCRAPER_RULES[genre] ? genre : 'default';
  }

  function normalizeWhitespace(text) {
    return String(text || '')
      .replace(/\u00a0/g, ' ')
      .replace(/\r/g, '\n')
      .replace(/\n{3,}/g, '\n\n')
      .replace(/[ \t]+\n/g, '\n')
      .replace(/\n[ \t]+/g, '\n')
      .replace(/[ \t]{2,}/g, ' ')
      .trim();
  }

  function toAbsoluteUrl(url) {
    try {
      return new URL(String(url || ''), location.href).href;
    } catch (_error) {
      return '';
    }
  }

  function normalizeImageUrl(url) {
    let normalized = toAbsoluteUrl(url)
      .replace(/\?.*$/, '');

    // cover500/ → cover/  (ディレクトリ名のサイズ指定を統一)
    normalized = normalized.replace(/cover\d+(?=\/)/gi, 'cover');
    // cover500.jpg → cover.jpg  (ファイル名のサイズ指定を統一)
    normalized = normalized.replace(/cover\d+(?=\.)/gi, 'cover');

    // アラジン画像のバリエーションを統一:
    // K692137983_01.jpg / K692137983_f.jpg / k692137983_1.jpg → k692137983.jpg
    // cover/ ディレクトリ内の画像のみ対象（letslook等のページ連番は保持）
    if (/(?:image|cdnimage)\.aladin\.co\.kr/i.test(normalized)) {
      normalized = normalized.toLowerCase()
        .replace(/(\/cover\/[a-z]\d+)_[a-z0-9]+(\.\w+)$/, '$1$2');
    }

    return normalized;
  }

  function isAladinDetailImage(url) {
    return /(?:image|cdnimage)\.aladin\.co\.kr/i.test(url);
  }

  function getElementText(element) {
    if (!element) {
      return '';
    }
    return normalizeWhitespace(element.innerText || element.textContent || '');
  }

  function pickFirstText(selectors, options = {}) {
    const minLength = options.minLength ?? 1;
    const maxLength = options.maxLength ?? Infinity;
    const sectionOptions = {
      includeSectionLabels: options.includeSectionLabels || [],
      excludeSectionLabels: options.excludeSectionLabels || []
    };
    const hasSectionFilter = sectionOptions.includeSectionLabels.length > 0 || sectionOptions.excludeSectionLabels.length > 0;
    const candidates = [];

    for (const selector of selectors) {
      const elements = document.querySelectorAll(selector);
      for (const element of elements) {
        if (hasSectionFilter && !shouldIncludeBySection(element, sectionOptions)) {
          continue;
        }

        const text = getElementText(element);
        if (text.length < minLength) {
          continue;
        }
        candidates.push(text.slice(0, maxLength));
      }
      if (candidates.length > 0) {
        break;
      }
    }

    if (candidates.length === 0) {
      return '';
    }

    return candidates.sort((left, right) => right.length - left.length)[0];
  }

  function pickTextByPlans(plans, fallbackSelectors, fallbackOptions = {}) {
    for (const plan of plans || []) {
      const selectors = uniqueItems(plan.selectors || []);
      if (!selectors.length) {
        continue;
      }

      const text = pickFirstText(selectors, plan);
      if (text) {
        return text;
      }
    }

    return pickFirstText(fallbackSelectors, fallbackOptions);
  }

  function cleanDescriptionText(text, options = {}) {
    const lines = normalizeWhitespace(text)
      .split('\n')
      .map(line => line.trim())
      .filter(Boolean);

    while (lines.length > 1 && /(소개|책소개|출판사 제공|상품설명|상세정보)$/i.test(lines[0])) {
      lines.shift();
    }

    const cleanedLines = [];
    const stopPatterns = options.keepStructuredSections
      ? []
      : [
          /^\[CONTENTS\]/i,
          /^\[제품구성\]/i,
          /^\[track/i,
          /^DISC\s*\d+/i,
          /^\d+\.\s*(PHOTOBOOK|DIGIPACK|DISC|FOLDED POSTER|HOLOGRAPHIC POSTCARD|MINI)/i
        ];

    for (const line of lines) {
      if (line === ',') {
        continue;
      }
      if (stopPatterns.some(pattern => pattern.test(line))) {
        break;
      }
      cleanedLines.push(line);
    }

    return normalizeWhitespace(cleanedLines.join('\n'));
  }

  function cleanBasicInfoText(text) {
    const lines = normalizeWhitespace(text)
      .split('\n')
      .map(line => line.trim())
      .filter(Boolean);

    const kept = [];
    const stopPatterns = [
      /^주제 분류/i,
      /^신간알림 신청/i
    ];
    const skipPatterns = [
      /.+\s>\s.+/
    ];

    for (const line of lines) {
      if (stopPatterns.some(pattern => pattern.test(line))) {
        break;
      }
      if (skipPatterns.some(pattern => pattern.test(line))) {
        continue;
      }
      kept.push(line);
    }

    return normalizeWhitespace(kept.join('\n').replace(/^기본정보\s*/i, ''));
  }

  function pickBasicInfoText(selectors) {
    const sectionSelector = '.Ere_prod_mconts_box';
    const labelSelector = '.Ere_prod_mconts_LL, .Ere_prod_mconts_LS, .conts_Tleft strong';
    const detailSelector = '.conts_info_list1, .Ere_prod_info_list, #productBasicInfo, .basic_info, .info_list';

    for (const section of document.querySelectorAll(sectionSelector)) {
      const label = getElementText(section.querySelector(labelSelector)).trim();
      if (!/기본정보/i.test(label)) {
        continue;
      }

      const detail = section.querySelector(detailSelector);
      if (!detail) {
        continue;
      }

      const listItems = [...detail.querySelectorAll('li')]
        .map(item => getElementText(item).trim())
        .filter(Boolean);
      if (listItems.length) {
        return listItems.join('\n');
      }

      const text = getElementText(detail).trim();
      if (text) {
        return text;
      }
    }

    return pickFirstText(selectors, { minLength: 5, maxLength: 500 });
  }

  function getImageSize(image) {
    const attrWidth = Number.parseInt(image.getAttribute('width') || '', 10);
    const attrHeight = Number.parseInt(image.getAttribute('height') || '', 10);

    return {
      width: image.naturalWidth || image.width || attrWidth || 0,
      height: image.naturalHeight || image.height || attrHeight || 0
    };
  }

  function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  function elementHasContent(element) {
    if (!element) {
      return false;
    }
    if (element instanceof HTMLImageElement) {
      return Boolean(element.currentSrc || element.src || element.getAttribute('src'));
    }
    if (element.querySelector('img')) {
      return true;
    }
    return getElementText(element).length > 20;
  }

  function isElementVisible(element) {
    if (!(element instanceof Element)) {
      return false;
    }

    const style = globalThis.getComputedStyle(element);
    if (style.display === 'none' || style.visibility === 'hidden' || style.pointerEvents === 'none') {
      return false;
    }

    const rect = element.getBoundingClientRect();
    return rect.width > 0 && rect.height > 0;
  }

  function findCarouselButtons(selectors) {
    const nodes = [];
    for (const selector of selectors || []) {
      nodes.push(...document.querySelectorAll(selector));
    }
    return [...new Set(nodes)].filter(isElementVisible);
  }

  function hasSelectorContent(selector) {
    return [...document.querySelectorAll(selector)].some(elementHasContent);
  }

  async function waitForSelectors(selectors, timeoutMs = 4000, intervalMs = 100) {
    const targets = selectors.filter(Boolean);
    if (targets.length === 0) {
      return false;
    }
    if (targets.some(hasSelectorContent)) {
      return true;
    }

    const deadline = Date.now() + timeoutMs;
    while (Date.now() < deadline) {
      await delay(intervalMs);
      if (targets.some(hasSelectorContent)) {
        return true;
      }
    }
    return false;
  }

  function getSectionLabelText(node) {
    const section = node.closest('.Ere_prod_mconts_box');
    if (!section) {
      return '';
    }

    return getElementText(
      section.querySelector('.Ere_prod_mconts_LL, .Ere_prod_mconts_LS, .conts_Tleft strong')
    );
  }

  function isExcludedByAncestor(node, selectors) {
    return selectors.some(selector => node.closest(selector));
  }

  function shouldIncludeBySection(node, options) {
    const sectionLabel = getSectionLabelText(node);
    const excludeLabels = options.excludeSectionLabels || [];
    const includeLabels = options.includeSectionLabels || [];

    if (excludeLabels.some(pattern => pattern.test(sectionLabel))) {
      return false;
    }

    if (!includeLabels.length) {
      return true;
    }

    if (!sectionLabel) {
      return false;
    }

    return includeLabels.some(pattern => pattern.test(sectionLabel));
  }

  function buildCollectOptions(plan, rule, coverUrl) {
    return {
      coverUrl,
      includeSectionLabels: plan.includeSectionLabels || rule.includeSectionLabels || [],
      excludeSectionLabels: plan.excludeSectionLabels || rule.excludeSectionLabels || [],
      excludeAncestorSelectors: uniqueItems(rule.excludeAncestorSelectors || [], plan.excludeAncestorSelectors || []),
      allowPathPatterns: plan.allowPathPatterns || [],
      maxCount: Number.isFinite(plan.maxCount) ? plan.maxCount : Infinity,
      preferLast: Boolean(plan.preferLast)
    };
  }

  function filterImageElements(elements, options, seenUrls) {
    const accepted = [];

    for (const image of elements) {
      if (accepted.length >= options.maxCount) {
        break;
      }

      if (isExcludedByAncestor(image, options.excludeAncestorSelectors)) {
        continue;
      }

      const closestLink = image.closest('a[href]');
      const href = closestLink ? toAbsoluteUrl(closestLink.getAttribute('href') || '') : '';
      if (/\/events\//i.test(href)) {
        continue;
      }

      if (!shouldIncludeBySection(image, options)) {
        continue;
      }

      const rawUrl = image.currentSrc || image.src || image.dataset.src || image.getAttribute('data-src') || image.getAttribute('src') || '';
      const absoluteUrl = toAbsoluteUrl(rawUrl);
      if (!absoluteUrl || !isAladinDetailImage(absoluteUrl)) {
        continue;
      }
      if (/icon|btn|blank|loading|arrow|logo|spacer|check_off|check_on/i.test(absoluteUrl)) {
        continue;
      }
      if (options.allowPathPatterns.length && !options.allowPathPatterns.some(pattern => pattern.test(absoluteUrl))) {
        continue;
      }

      const normalizedUrl = normalizeImageUrl(absoluteUrl);
      if (!normalizedUrl || seenUrls.has(normalizedUrl)) {
        continue;
      }

      const { width, height } = getImageSize(image);
      if ((width && width < 50) || (height && height < 50)) {
        continue;
      }

      seenUrls.add(normalizedUrl);
      accepted.push(absoluteUrl);
    }

    return accepted;
  }

  function collectBySelectors(selectors, options, seenUrls) {
    const nodes = [];
    for (const selector of selectors) {
      nodes.push(...document.querySelectorAll(selector));
    }

    const elements = options.preferLast ? nodes.reverse() : nodes;
    const accepted = filterImageElements(elements, options, seenUrls);
    return options.preferLast ? accepted.reverse() : accepted;
  }

  async function collectInteractiveImages(selectors, options, seenUrls, plan) {
    const collected = [];
    const appendFound = () => {
      const found = collectBySelectors(selectors, options, seenUrls);
      if (found.length) {
        collected.push(...found);
      }
      return found.length;
    };

    appendFound();

    if (!plan.interactCarousel) {
      return collected;
    }

    let stagnantTurns = 0;
    const maxClicks = Number.isFinite(plan.maxClicks) ? plan.maxClicks : 8;
    const clickDelayMs = Number.isFinite(plan.clickDelayMs) ? plan.clickDelayMs : 250;

    for (let step = 0; step < maxClicks; step += 1) {
      const button = findCarouselButtons(plan.clickSelectors || [])[0];
      if (!button) {
        break;
      }

      if (typeof button.click === 'function') {
        button.click();
      } else {
        button.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
      }

      await delay(clickDelayMs);
      const added = appendFound();
      if (added === 0) {
        stagnantTurns += 1;
      } else {
        stagnantTurns = 0;
      }
      if (stagnantTurns >= 2) {
        break;
      }
    }

    return collected;
  }

  async function collectImagesByRule(rule, coverUrl) {
    const normalizedCoverUrl = normalizeImageUrl(coverUrl || '');
    const seenUrls = new Set(normalizedCoverUrl ? [normalizedCoverUrl] : []);
    const collected = [];

    for (const plan of rule.imagePlans || []) {
      const selectors = uniqueItems(plan.selectors || []);
      if (!selectors.length) {
        continue;
      }

      const options = buildCollectOptions(plan, rule, coverUrl);
      collected.push(...await collectInteractiveImages(selectors, options, seenUrls, plan));
    }

    if (!collected.length && rule.allowGenericFallback) {
      const fallbackOptions = buildCollectOptions({
        selectors: SHARED_IMAGE_SELECTORS,
        maxCount: rule.fallbackMaxCount || 4
      }, rule, coverUrl);
      collected.push(...collectBySelectors(SHARED_IMAGE_SELECTORS, fallbackOptions, seenUrls));
    }

    return collected;
  }

  function extractItemIdFromLocation() {
    try {
      const itemId = new URL(location.href).searchParams.get('ItemId') || '';
      return /^\d+$/.test(itemId) ? itemId : null;
    } catch (_error) {
      return null;
    }
  }

  async function scrapePage(options = {}) {
    const genre = normalizeGenre(options.genre);
    const rule = SCRAPER_RULES[genre] || SCRAPER_RULES.default;

    await waitForSelectors(rule.waitSelectors || []);

    const descriptionSelectors = uniqueItems(rule.descriptionSelectors || [], SHARED_DESCRIPTION_SELECTORS);
    const basicInfoSelectors = uniqueItems(rule.basicInfoSelectors || [], SHARED_BASIC_INFO_SELECTORS);

    const additionalImages = await collectImagesByRule(rule, options.coverUrl || '');
    const description = cleanDescriptionText(
      pickTextByPlans(rule.descriptionPlans || [], descriptionSelectors, { minLength: 20, maxLength: 2000 }),
      { keepStructuredSections: genre === '音楽映像' || genre === 'マンガ' }
    );
    const basicInfo = cleanBasicInfoText(pickBasicInfoText(basicInfoSelectors));

    return {
      itemId: extractItemIdFromLocation(),
      pageUrl: location.href,
      genre,
      additionalImages,
      additionalImagesJoined: additionalImages.join(';'),
      description,
      basicInfo
    };
  }

  globalThis.__ALADIN_SCRAPER__ = {
    version: VERSION,
    scrapePage
  };
})();


