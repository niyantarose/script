// content.js - 韓国グッズ マルチショップ商品情報取得

const SHOP_DEFINITIONS = [
  { code: 'JMEE', name: 'JMEE', platform: 'cafe24', host: /(^|\.)j-meeshop\.com$/i },
  { code: 'MOFUN', name: 'MOFUN', platform: 'cafe24', host: /(^|\.)mostore\.co\.kr$/i },
  { code: 'TOON', name: 'TOONIQUE', platform: 'cafe24', host: /(^|\.)toonique\.co\.kr$/i },
  { code: 'WTSHOP', name: 'WEBTOONSHOP', platform: 'cafe24', host: /(^|\.)webtoonshop\.com$/i },
  { code: 'WTF', name: 'WebtoonFriends', platform: 'marpple', host: /(^|\.)webtoonfriends\.marpple\.shop$/i },
  {
    code: 'OFW',
    name: 'Official_W',
    platform: 'naver',
    host: /(^|\.)smartstore\.naver\.com$/i,
    pathPrefix: /^\/official_w(\/|$)/i,
  },
];

function parseUrl(value) {
  try {
    return new URL(String(value || ''), location.href);
  } catch (_e) {
    return null;
  }
}

function detectShop(urlValue = location.href) {
  const url = parseUrl(urlValue);
  if (!url) return null;
  const host = url.hostname.toLowerCase();
  const path = url.pathname || '/';

  for (const def of SHOP_DEFINITIONS) {
    if (!def.host.test(host)) continue;
    if (def.pathPrefix && !def.pathPrefix.test(path)) continue;
    return def;
  }
  return null;
}

function toAbsoluteUrl(rawUrl, baseUrl = location.href) {
  const text = String(rawUrl || '').trim();
  if (!text) return '';
  if (text.startsWith('data:') || text.startsWith('blob:') || text.startsWith('javascript:')) return '';
  try {
    return new URL(text, baseUrl).href;
  } catch (_e) {
    return '';
  }
}

function isImageUrl(url) {
  const text = String(url || '').trim();
  if (!/^https?:\/\//i.test(text)) return false;
  if (/image_zoom\d?\.html|\/product\/detail\.html|\.html?(\?|$)/i.test(text)) return false;
  if (/\.(jpg|jpeg|png|webp|gif|bmp|avif)(\?|$)/i.test(text)) return true;
  if (/\/web\/upload\//i.test(text)) return true;
  return false;
}

function normalizeImageUrl(rawUrl, baseUrl = location.href) {
  let url = toAbsoluteUrl(rawUrl, baseUrl);
  if (!url) return '';

  // Cafe24商品画像は解像度を上げる
  url = url.replace(/\/(web\/)?product\/(tiny|small|medium)\//i, '/$1product/big/');

  if (!isImageUrl(url)) return '';
  return url;
}

function getDedupeKey(url) {
  return String(url || '').replace(/[?#].*$/, '');
}

function pushImageUrl(list, seenKeys, rawUrl, baseUrl = location.href) {
  const normalized = normalizeImageUrl(rawUrl, baseUrl);
  if (!normalized) return;
  const key = getDedupeKey(normalized);
  if (!key || seenKeys.has(key)) return;
  seenKeys.add(key);
  list.push(normalized);
}

function readImgCandidates(img, includeParentHref = true) {
  const candidates = [
    img.getAttribute('ec-data-src'),
    img.getAttribute('data-ec-data-src'),
    img.getAttribute('data-src'),
    img.getAttribute('data-lazy-src'),
    img.getAttribute('data-original'),
    img.currentSrc,
    img.src,
  ];
  if (includeParentHref) {
    candidates.push(img.closest('a')?.href || '');
  }
  return candidates;
}

function hasTinyMinHeight(img) {
  const styleText = String(img.getAttribute('style') || '').toLowerCase();
  const match = styleText.match(/min-height\s*:\s*(\d+(?:\.\d+)?)px/);
  if (!match) return false;
  const minHeight = parseFloat(match[1]);
  return !Number.isNaN(minHeight) && minHeight > 0 && minHeight < 180;
}

function hasTinyDimensions(img) {
  const widthAttr = parseInt(img.getAttribute('width') || '0', 10) || 0;
  const heightAttr = parseInt(img.getAttribute('height') || '0', 10) || 0;
  const naturalWidth = img.naturalWidth || 0;
  const naturalHeight = img.naturalHeight || 0;
  const w = Math.max(widthAttr, naturalWidth);
  const h = Math.max(heightAttr, naturalHeight);
  return w > 0 && h > 0 && (w < 180 || h < 180);
}

function isBadgeLikeImage(img, url) {
  const lowerUrl = String(url || '').toLowerCase();
  const markerText = [
    img.alt || '',
    img.className || '',
    img.id || '',
    img.closest('[class]')?.className || '',
    img.closest('[id]')?.id || '',
  ].join(' ').toLowerCase();

  if (/(badge|icon|ico|new|event|naver|mileage|coupon|benefit|delivery|point|logo|banner)/.test(markerText)) {
    return true;
  }

  if (/\/(category|icon|common|layout|banner)\//.test(lowerUrl)) {
    return true;
  }

  if (hasTinyDimensions(img) || hasTinyMinHeight(img)) {
    return true;
  }

  return false;
}

function collectImageUrlsBySelectors(baseRoot, selectors, options = {}) {
  const root = baseRoot || document;
  const includeParentHref = options.includeParentHref !== false;
  const skipInfoArea = !!options.skipInfoArea;
  const urls = [];
  const seen = new Set();

  selectors.forEach(selector => {
    root.querySelectorAll(selector).forEach(img => {
      if (!img) return;
      if (skipInfoArea && img.closest('.infoArea, .headingArea, .xans-product-detaildesign, .xans-product-action, .xans-product-option, .xans-product-listitem')) {
        return;
      }

      const candidates = readImgCandidates(img, includeParentHref);
      for (const candidate of candidates) {
        const url = normalizeImageUrl(candidate, location.href);
        if (!url) continue;
        if (isBadgeLikeImage(img, url)) continue;
        const key = getDedupeKey(url);
        if (!key || seen.has(key)) continue;
        seen.add(key);
        urls.push(url);
        break;
      }
    });
  });

  return urls;
}

function getCafe24DetailRoot() {
  return (
    document.querySelector('.xans-product-additional.left_area') ||
    document.querySelector('#prdDetail') ||
    document.querySelector('.xans-product-additional') ||
    document.querySelector('.xans-product-detail #prdDetail') ||
    document.querySelector('.xans-product-detail .xans-product-additional') ||
    null
  );
}

function isVisibleElement(el) {
  if (!el) return false;
  const style = window.getComputedStyle(el);
  if (!style || style.display === 'none' || style.visibility === 'hidden' || style.opacity === '0') return false;
  const rect = el.getBoundingClientRect();
  return rect.width > 0 && rect.height > 0;
}

async function expandDetailMoreButtons(detailRoot) {
  if (!detailRoot) return;
  for (let i = 0; i < 5; i += 1) {
    const moreButtons = Array.from(detailRoot.querySelectorAll('button, a, [role="button"], div, span'))
      .filter(el => {
        const txt = (el.textContent || '').trim();
        if (!txt) return false;
        if (!/(더보기|더 보기|more|show more)/i.test(txt)) return false;
        if (!isVisibleElement(el)) return false;
        if (el.dataset.krGoodsExpandedClicked === '1') return false;
        return true;
      });

    if (moreButtons.length === 0) break;
    moreButtons.forEach(btn => {
      btn.dataset.krGoodsExpandedClicked = '1';
      btn.click();
    });
    await new Promise(resolve => setTimeout(resolve, 250));
  }
}

function collectImageUrlsFromHtml(root) {
  const urls = [];
  const seen = new Set();
  const html = root?.innerHTML || '';
  const absoluteMatches = html.match(/https?:\/\/[^"'\s>]+?\.(jpg|jpeg|png|webp|gif|bmp|avif)(\?[^"'\s>]*)?/gi) || [];
  const relativeMatches = html.match(/\/web\/upload\/[^"'\s>]+?\.(jpg|jpeg|png|webp|gif|bmp|avif)(\?[^"'\s>]*)?/gi) || [];
  [...absoluteMatches, ...relativeMatches].forEach(raw => pushImageUrl(urls, seen, raw, location.href));
  return urls;
}

function textFromElement(el) {
  return String(el?.innerText || el?.textContent || '').replace(/\r/g, '').trim();
}

function normalizeWhitespace(text) {
  return String(text || '').replace(/\s+/g, ' ').trim();
}

function extractProductName() {
  const selectorCandidates = [
    '.headingArea h2',
    '.headingArea',
    '.xans-product-detail .prd-name',
    '.xans-product-detail .product_name',
    '.xans-product-detail .infoArea h2',
    '.xans-product-detail .infoArea h1',
    '.infoArea h2',
    '.infoArea h1',
    '[class*="product_name"]',
    '[class*="productName"]',
    'h1',
  ];

  for (const selector of selectorCandidates) {
    const value = normalizeWhitespace(textFromElement(document.querySelector(selector)));
    if (value) return value;
  }

  const ogTitle = normalizeWhitespace(document.querySelector('meta[property="og:title"]')?.content || '');
  if (ogTitle) return ogTitle;

  if (document.title) {
    const parts = document.title.split('-').map(s => normalizeWhitespace(s)).filter(Boolean);
    if (parts.length > 0) return parts[parts.length - 1];
  }

  return '';
}

function parsePriceNumber(text) {
  const digits = String(text || '').replace(/[^\d]/g, '');
  return digits || '';
}

function findPriceInJson(value) {
  if (!value) return '';
  if (Array.isArray(value)) {
    for (const item of value) {
      const found = findPriceInJson(item);
      if (found) return found;
    }
    return '';
  }
  if (typeof value === 'object') {
    if (value.offers) {
      const fromOffers = findPriceInJson(value.offers);
      if (fromOffers) return fromOffers;
    }
    if (value.price !== undefined && value.price !== null) {
      const parsed = parsePriceNumber(value.price);
      if (parsed) return parsed;
    }
    for (const key of Object.keys(value)) {
      const found = findPriceInJson(value[key]);
      if (found) return found;
    }
  }
  return '';
}

function extractPrice() {
  const metaPrice = parsePriceNumber(document.querySelector('meta[property="product:price:amount"]')?.content || '');
  if (metaPrice) return metaPrice;

  const jsonLdScripts = Array.from(document.querySelectorAll('script[type="application/ld+json"]'));
  for (const script of jsonLdScripts) {
    try {
      const data = JSON.parse(script.textContent || '{}');
      const price = findPriceInJson(data);
      if (price) return price;
    } catch (_e) {
      // no-op
    }
  }

  const scopedRoot = document.querySelector('.xans-product-detail .infoArea, .infoArea, main, #container') || document;
  const selectorCandidates = [
    '.price',
    '.prd-price',
    '.sale_price',
    '.price_sale',
    '.salePriceText',
    '.price_num',
    '[class*="salePrice"]',
    '[class*="price"]',
  ];

  for (const selector of selectorCandidates) {
    const scopedEl = scopedRoot.querySelector(selector);
    const scopedPrice = parsePriceNumber(textFromElement(scopedEl));
    if (scopedPrice) return scopedPrice;
  }

  for (const selector of selectorCandidates) {
    const globalEl = document.querySelector(selector);
    const globalPrice = parsePriceNumber(textFromElement(globalEl));
    if (globalPrice) return globalPrice;
  }

  return '';
}

function getFallbackProductIdFromPath(pathname) {
  const segments = String(pathname || '/')
    .split('/')
    .map(x => x.trim())
    .filter(Boolean);
  if (segments.length === 0) return '';
  const last = segments[segments.length - 1];
  return last.replace(/[^\w-]/g, '').slice(0, 60);
}

function extractProductId(shop, pageUrl) {
  const search = pageUrl.searchParams;
  const pathname = pageUrl.pathname || '/';
  const queryKeys = ['product_no', 'productNo', 'goodsNo', 'id', 'no'];

  for (const key of queryKeys) {
    const value = String(search.get(key) || '').trim();
    if (/^\d+$/.test(value)) return value;
  }

  if (shop.platform === 'cafe24') {
    const match = pathname.match(/\/product\/[^/]*\/(\d+)(?:\/|$)/i);
    if (match?.[1]) return match[1];
  }

  const naverMatch = pathname.match(/\/products\/(\d+)(?:\/|$)/i);
  if (naverMatch?.[1]) return naverMatch[1];

  const genericProductMatch = pathname.match(/\/product\/(\d+)(?:\/|$)/i);
  if (genericProductMatch?.[1]) return genericProductMatch[1];

  const numericSegments = pathname.match(/\/(\d{3,})(?=\/|$)/g) || [];
  if (numericSegments.length > 0) {
    const last = numericSegments[numericSegments.length - 1];
    const found = (last.match(/\d{3,}/) || [])[0];
    if (found) return found;
  }

  const ogUrl = document.querySelector('meta[property="og:url"]')?.content || '';
  const ogParsed = parseUrl(ogUrl);
  if (ogParsed && ogParsed.href !== pageUrl.href) {
    const fromOg = extractProductId(shop, ogParsed);
    if (fromOg) return fromOg;
  }

  return getFallbackProductIdFromPath(pathname);
}

function getPreferredDetailTextRoot(shop) {
  if (shop.platform === 'cafe24') {
    return getCafe24DetailRoot() || document.querySelector('.xans-product-additional');
  }

  const candidates = [
    '#prdDetail',
    '[id*="detail"]',
    '[class*="detail"]',
    '[class*="description"]',
    '[class*="product-detail"]',
    '[data-testid*="detail"]',
  ];

  let best = null;
  let bestLen = 0;
  for (const selector of candidates) {
    const elements = Array.from(document.querySelectorAll(selector));
    elements.forEach(el => {
      const txt = textFromElement(el);
      if (txt.length > bestLen) {
        best = el;
        bestLen = txt.length;
      }
    });
  }
  return best;
}

function extractProductInfoOnly(rawText) {
  const text = String(rawText || '')
    .replace(/\r/g, '')
    .replace(/\u00A0/g, ' ')
    .trim();
  if (!text) return '';

  const startMarkers = ['[상품정보]', '상품정보'];
  let start = -1;
  startMarkers.forEach(marker => {
    const idx = text.indexOf(marker);
    if (idx >= 0 && (start === -1 || idx < start)) start = idx;
  });

  let selected = start >= 0 ? text.slice(start) : text;
  const stopMarkers = [
    '[유의사항]', '유의사항',
    '[주의사항]', '주의사항',
    '[배송안내]', '[배송]',
    '[교환/반품]', '[교환및반품]',
    '상품후기', '[상품후기]', '[review]', 'review',
  ];
  let end = selected.length;
  stopMarkers.forEach(marker => {
    const idx = selected.indexOf(marker);
    if (idx > 0 && idx < end) end = idx;
  });
  selected = selected.slice(0, end);

  return selected
    .split('\n')
    .map(line => line.trim())
    .reduce((acc, line) => {
      if (line === '' && acc[acc.length - 1] === '') return acc;
      acc.push(line);
      return acc;
    }, [])
    .join('\n')
    .trim()
    .slice(0, 1000);
}

async function collectImageSet(shop) {
  const productRoot = document.querySelector('.xans-product-detail, #contents, #container, main') || document;
  const mainAndGallery = [];
  const seen = new Set();

  const ogImage = normalizeImageUrl(document.querySelector('meta[property="og:image"]')?.content || '', location.href);
  if (ogImage) pushImageUrl(mainAndGallery, seen, ogImage, location.href);

  if (shop.platform === 'cafe24') {
    const detailRoot = getCafe24DetailRoot();
    await expandDetailMoreButtons(detailRoot);

    const cafeMain = collectImageUrlsBySelectors(productRoot, [
      '.BigImage',
      '.thumbnail img',
      '.listImg img',
      '.imgArea .BigImage',
      '.xans-product-image .BigImage',
      '.xans-product-image img',
      '.xans-product-addimage img',
      '.xans-product-detail .xans-product-addimage img',
      '.xans-product-detail .listImg img',
      'a[href*="image_zoom"] img',
      'a[href*="display_group"] img',
    ]);
    cafeMain.forEach(url => pushImageUrl(mainAndGallery, seen, url, location.href));

    const detailUrls = collectImageUrlsBySelectors(detailRoot || productRoot, [
      'img[ec-data-src]',
      'img[data-ec-data-src]',
      '.edb-img-tag-w img',
      'img',
    ], {
      includeParentHref: false,
      skipInfoArea: true,
    });

    const additionalSeen = new Set();
    const main = mainAndGallery[0] || '';
    const additional = [];
    const append = rawUrl => {
      const url = normalizeImageUrl(rawUrl, location.href);
      if (!url) return;
      const key = getDedupeKey(url);
      if (!key || additionalSeen.has(key)) return;
      if (main && key === getDedupeKey(main)) return;
      additionalSeen.add(key);
      additional.push(url);
    };

    mainAndGallery.slice(1).forEach(append);
    detailUrls.forEach(append);
    if (additional.length === 0) {
      const htmlUrls = collectImageUrlsFromHtml(detailRoot || productRoot);
      htmlUrls.forEach(append);
    }
    return { main: mainAndGallery[0] || '', additional };
  }

  const genericMain = collectImageUrlsBySelectors(productRoot, [
    '[class*="thumbnail"] img',
    '[class*="thumb"] img',
    '[class*="gallery"] img',
    '[class*="carousel"] img',
    '[class*="swiper"] img',
    '[data-testid*="image"] img',
    'img',
  ], { includeParentHref: true, skipInfoArea: false });
  genericMain.forEach(url => pushImageUrl(mainAndGallery, seen, url, location.href));

  const detailRoot = getPreferredDetailTextRoot(shop);
  const detailUrls = collectImageUrlsBySelectors(detailRoot || productRoot, [
    'img[data-src]',
    'img[data-original]',
    'img',
  ], { includeParentHref: false, skipInfoArea: false });

  const additionalSeen = new Set();
  const main = mainAndGallery[0] || '';
  const additional = [];
  const append = rawUrl => {
    const url = normalizeImageUrl(rawUrl, location.href);
    if (!url) return;
    const key = getDedupeKey(url);
    if (!key || additionalSeen.has(key)) return;
    if (main && key === getDedupeKey(main)) return;
    additionalSeen.add(key);
    additional.push(url);
  };

  mainAndGallery.slice(1).forEach(append);
  detailUrls.forEach(append);
  if (additional.length === 0) {
    const htmlUrls = collectImageUrlsFromHtml(detailRoot || productRoot);
    htmlUrls.forEach(append);
  }

  return { main: mainAndGallery[0] || '', additional };
}

async function getProductInfo() {
  try {
    const shop = detectShop(location.href);
    if (!shop) {
      return { success: false, error: '対応ショップのページではありません' };
    }

    const pageUrl = new URL(location.href);
    const 商品ID = extractProductId(shop, pageUrl);
    const 商品名 = extractProductName();
    const 価格 = extractPrice();

    const images = await collectImageSet(shop);
    const メイン画像URL = images.main || '';
    const 追加画像URL = (images.additional || []).join(';');

    const detailRoot = getPreferredDetailTextRoot(shop);
    const 商品説明Raw = textFromElement(detailRoot);
    const 商品説明 = extractProductInfoOnly(商品説明Raw);

    const 購入URL = location.origin + location.pathname;
    const safeId = 商品ID || getFallbackProductIdFromPath(location.pathname) || String(Date.now());
    const 重複チェックキー = `${shop.code}-${safeId}`;

    return {
      success: true,
      data: {
        ショップ: shop.code,
        ショップコード: shop.code,
        商品ID: safeId,
        商品名,
        価格,
        メイン画像URL,
        追加画像URL,
        商品説明,
        購入URL,
        重複チェックキー,
      }
    };
  } catch (e) {
    return { success: false, error: e?.message || String(e) };
  }
}

if (!globalThis.__KR_GOODS_LISTENER_BOUND__) {
  globalThis.__KR_GOODS_LISTENER_BOUND__ = true;
  chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === 'getProductInfo') {
      getProductInfo()
        .then(sendResponse)
        .catch(e => sendResponse({ success: false, error: e?.message || String(e) }));
      return true;
    }
    sendResponse({ success: false, error: 'unsupported action' });
    return false;
  });
}

