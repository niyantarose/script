(() => {
  const DESCRIPTION_SELECTORS = [
    '.pContent[id$="_Introduce"]',
    '#bookDescriptionToggle',
    '#productDescription',
    '.Ere_prod_mconts_box'
  ];

  const BASIC_INFO_SELECTORS = [
    '.Ere_prod_mconts_box .conts_info_list1',
    '.Ere_prod_info_list',
    '#productBasicInfo',
    '.basic_info',
    '.conts_info_list1'
  ];

  const IMAGE_SELECTORS = [
    '#CoverMainImage',
    '#swiper-container-cover .swiper-slide img',
    '.pContent[id$="_Introduce"] img',
    '#bookDescriptionToggle img',
    '#productDescription img',
    '.Ere_prod_mconts_box img'
  ];

  function normalizeWhitespace(value) {
    return String(value || '')
      .replace(/\u00a0/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function toAbsoluteUrl(value) {
    try {
      return new URL(String(value || ''), location.href).href;
    } catch (_error) {
      return '';
    }
  }

  function normalizeImageUrl(value) {
    return toAbsoluteUrl(value).replace(/\?.*$/, '');
  }

  function pickLongestText(selectors) {
    const candidates = [];
    for (const selector of selectors) {
      for (const element of document.querySelectorAll(selector)) {
        const text = normalizeWhitespace(element.innerText || element.textContent || '');
        if (text.length >= 8) {
          candidates.push(text);
        }
      }
      if (candidates.length) {
        break;
      }
    }
    return candidates.sort((left, right) => right.length - left.length)[0] || '';
  }

  function pickMetaContent(selector) {
    const element = document.querySelector(selector);
    return normalizeWhitespace(element?.getAttribute('content') || '');
  }

  function pickText(selectors) {
    for (const selector of selectors) {
      const element = document.querySelector(selector);
      const text = normalizeWhitespace(element?.innerText || element?.textContent || '');
      if (text) {
        return text;
      }
    }
    return '';
  }

  function extractFieldFromBasicInfo(basicInfoText, labelPattern) {
    const lines = String(basicInfoText || '')
      .split(/\r?\n/)
      .map(line => normalizeWhitespace(line))
      .filter(Boolean);

    const matchedLine = lines.find(line => labelPattern.test(line));
    if (!matchedLine) {
      return '';
    }

    return normalizeWhitespace(matchedLine.replace(labelPattern, ''));
  }

  function extractItemId() {
    try {
      const itemId = new URL(location.href).searchParams.get('ItemId') || '';
      return /^\d+$/.test(itemId) ? itemId : '';
    } catch (_error) {
      return '';
    }
  }

  function collectCategoryText() {
    return normalizeWhitespace(
      [...document.querySelectorAll('.Ere_location a, .Ere_location span, #ulCategory a, #ulCategory li')]
        .map(element => element.textContent || '')
        .join(' > ')
    );
  }

  function collectImageCandidates() {
    const seen = new Set();
    const result = [];

    const pushCandidate = (url, kindHint, sourceSection) => {
      const normalized = normalizeImageUrl(url);
      if (!normalized || seen.has(normalized)) {
        return;
      }
      if (!/(?:image|cdnimage)\.aladin\.co\.kr/i.test(normalized)) {
        return;
      }
      if (/icon|logo|arrow|btn|loading/i.test(normalized)) {
        return;
      }
      seen.add(normalized);
      const orderHint = result.filter(item => item.kindHint === kindHint).length + 1;
      result.push({ url: normalized, kindHint, orderHint, sourceSection });
    };

    const coverUrl = pickMetaContent('meta[property="og:image"]') || toAbsoluteUrl(document.querySelector('#CoverMainImage')?.src || '');
    if (coverUrl) {
      pushCandidate(coverUrl, 'main', 'cover');
    }

    for (const selector of IMAGE_SELECTORS) {
      for (const image of document.querySelectorAll(selector)) {
        const width = image.naturalWidth || image.width || 0;
        const height = image.naturalHeight || image.height || 0;
        if ((width && width < 50) || (height && height < 50)) {
          continue;
        }

        const section = image.closest('.pContent[id$="_Introduce"]')
          ? 'description'
          : image.closest('#swiper-container-cover')
            ? 'cover_slide'
            : 'detail';
        const kindHint = section === 'cover_slide' ? 'main' : 'detail';
        pushCandidate(image.currentSrc || image.src || image.getAttribute('src') || '', kindHint, section);
      }
    }

    return result.slice(0, 20);
  }

  function probeAladinBook() {
    const basicInfo = pickLongestText(BASIC_INFO_SELECTORS);
    const description = pickLongestText(DESCRIPTION_SELECTORS);
    const title = pickMetaContent('meta[property="og:title"]')
      || pickText(['.Ere_bo_title', 'h1', '.prod_title'])
      || document.title.replace(/\s*-\s*알라딘.*$/, '').trim();
    const author = pickText(['.Ere_sub2_title a', '.info_list a', '.author a']);
    const publisher = extractFieldFromBasicInfo(basicInfo, /^(출판사|出版社)\s*/i);
    const publishDate = extractFieldFromBasicInfo(basicInfo, /^(출간일|刊行日|発売日)\s*/i);
    const isbn = extractFieldFromBasicInfo(basicInfo, /^(ISBN|EAN)\s*/i);
    const price = pickMetaContent('meta[property="product:price:amount"]')
      || pickText(['.Ritem', '.Ere_fs24', '.price'])
      || extractFieldFromBasicInfo(basicInfo, /^(정가|価格)\s*/i);
    const coverUrl = pickMetaContent('meta[property="og:image"]') || '';

    return {
      pageUrl: location.href,
      titleCandidate: title,
      siteProductCode: extractItemId(),
      rawFields: {
        title,
        author,
        publisher,
        publishDate,
        isbn,
        price,
        coverUrl,
        categoryText: collectCategoryText()
      },
      rawSections: {
        basicInfo,
        description
      },
      imageCandidates: collectImageCandidates(),
      rawHtml: document.documentElement?.outerHTML || ''
    };
  }

  globalThis.__NYANTA_PAGE_PROBE__ = {
    probe(adapterId) {
      if (adapterId !== 'aladin_book') {
        throw new Error(`unsupported adapter: ${adapterId}`);
      }
      return probeAladinBook();
    }
  };
})();
