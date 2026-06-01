// books.com.tw 商品情報スクレイパー

async function getProductInfo() {
  const url = window.location.href;
  const productCode = url.match(/products\/([^?]+)/)?.[1] || '';

  // 商品種別判定（書籍かグッズか）
  // 書籍: B, E, F など / グッズ: N
  const isBook = !/^[NM]/i.test(productCode);

  const getText = el => (el?.innerText || el?.textContent || '').trim();
  const toSingleLine = text => (text || '').replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
  const toMultiline = text => (text || '')
    .replace(/\u00a0/g, ' ')
    .replace(/\r/g, '')
    .split('\n')
    .map(line => line.trim())
    .filter(Boolean)
    .join('\n');
  const mergeTextBlocks = blocks => {
    const cleaned = (blocks || [])
      .map(toMultiline)
      .filter(Boolean);
    const merged = [];
    for (const block of cleaned) {
      const duplicated = merged.some(existing => existing.includes(block) || block.includes(existing));
      if (!duplicated) merged.push(block);
    }
    return merged.join('\n\n');
  };
  const uniq = arr => [...new Set((arr || []).filter(Boolean))];
  const escapeRegExp = str => str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const productCodeRegexSafe = escapeRegExp(productCode);
const keepGoodsThumbVariant = !isBook;
const TARGET_IMAGE_EDGE = 1000;
// books.com.tw の画像URL末尾バリアント:
//   _t_N = サムネイル系（N=1～3）
//   _b_N = 拡大系（N=1～3、N=2 が約700px、N=3 が約1000px）
// グッズはサムネがそのまま縮小品として使われるので _t_ を温存。
// 書籍は _t_ → _b_ への置換のみ行い、N（サイズ番号）は元URLに従う。
// 過剰なアップグレードは必要以上に大きい画像を取得する原因になるため、
// _b_N の N は変更しない。
const upgradeBooksImageVariant = (imageUrl, options = {}) => {
  const keepThumbVariant = options.keepThumbVariant === true;
  return String(imageUrl || '')
    .replace(/_t_(\d+)(?=\.(?:jpe?g|png|webp)(?:$|[?#]))/gi, keepThumbVariant ? '_t_$1' : '_b_$1')
    .replace(/\.webp(?=($|[?#]))/i, '.jpeg');
};
const cleanupLabeledValue = raw => {
  const text = toSingleLine(raw)
    .replace(/新功能介紹/g, '')
    .replace(/訂閱出版社新書快訊/g, '')
    .replace(/\s{2,}/g, ' ')
    .trim();
  if (!text) return '';

  // 1つのliに複数ラベルが混在するケースを切り分ける
  const trailingLabel = /(?:作者|譯者|译者|翻譯者|翻訳者|繪者|绘者|插畫|插画|出版社|出版商|出版日期|出版时间|發行日期|語言|ISBN|定價|優惠價|优惠价)\s*[：:]/;
  const idx = text.search(trailingLabel);
  return idx > 0 ? text.slice(0, idx).trim() : text;
};

const buildBooksBannerWrapperUrl = imageUrl => {
  let urlText = String(imageUrl || '').trim();
  if (!urlText) return '';
  if (urlText.startsWith('//')) urlText = `${window.location.protocol}${urlText}`;
  if (!/^https?:\/\/addons\.books\.com\.tw\/G\/ADbanner\//i.test(urlText)) return '';
  return `https://im1.book.com.tw/image/getImage?i=${encodeURIComponent(urlText)}`;
};

// getImageラッパーURLから direct 画像URLを優先して取得
const promoteLargeImageUrl = imageUrl => {
  if (!imageUrl) return '';

  let urlText = String(imageUrl || '').trim();
  if (!urlText) return '';
  if (urlText.startsWith('//')) urlText = `${window.location.protocol}${urlText}`;

  try {
    const parsed = new URL(urlText, window.location.href);
    const original = parsed.searchParams.get('i');
    if (original) {
      let originalUrl = decodeURIComponent(original);
      if (originalUrl.startsWith('//')) {
        originalUrl = `${window.location.protocol}${originalUrl}`;
      }
      const wrappedBannerUrl = buildBooksBannerWrapperUrl(originalUrl);
      if (wrappedBannerUrl) return wrappedBannerUrl;
      return upgradeBooksImageVariant(originalUrl, { keepThumbVariant: keepGoodsThumbVariant });
    }
    const upgradedUrl = upgradeBooksImageVariant(parsed.href, { keepThumbVariant: keepGoodsThumbVariant });
    return buildBooksBannerWrapperUrl(upgradedUrl) || upgradedUrl;
  } catch {
    const upgradedUrl = upgradeBooksImageVariant(urlText, { keepThumbVariant: keepGoodsThumbVariant });
    return buildBooksBannerWrapperUrl(upgradedUrl) || upgradedUrl;
  }
};

/** 一覧・シート表示用（サムネイル優先。大きすぎないサイズ） */
const normalizeDisplayImageUrl = rawUrl => {
  if (!rawUrl) return '';

  let urlText = String(rawUrl || '').trim();
  if (!urlText) return '';
  if (urlText.startsWith('//')) urlText = `${window.location.protocol}${urlText}`;

  try {
    const parsed = new URL(urlText, window.location.href);
    const original = parsed.searchParams.get('i');
    if (original) {
      let originalUrl = decodeURIComponent(original);
      if (originalUrl.startsWith('//')) {
        originalUrl = `${window.location.protocol}${originalUrl}`;
      }
      const wrappedBannerUrl = buildBooksBannerWrapperUrl(originalUrl);
      if (wrappedBannerUrl) return wrappedBannerUrl;
      return upgradeBooksImageVariant(originalUrl, { keepThumbVariant: true });
    }

    const href = parsed.href;
    const thumbVariant = href.replace(/_b_(\d+)(?=\.(?:jpe?g|png|webp)(?:$|[?#]))/gi, '_t_$1');
    if (thumbVariant !== href) return thumbVariant;

    const upgradedUrl = upgradeBooksImageVariant(href, { keepThumbVariant: false });
    return buildBooksBannerWrapperUrl(upgradedUrl) || upgradedUrl;
  } catch {
    const upgradedUrl = upgradeBooksImageVariant(urlText, { keepThumbVariant: true });
    return buildBooksBannerWrapperUrl(upgradedUrl) || upgradedUrl;
  }
};

/** ダウンロード・追加画像用（高解像度 _b_ 優先） */
const normalizeDownloadImageUrl = rawUrl => {
  if (!rawUrl) return '';
  return promoteLargeImageUrl(rawUrl);
};

const normalizeImageUrl = rawUrl => normalizeDownloadImageUrl(rawUrl);

const IMAGE_VIEWER_OVERLAY_SELECTORS = [
  '.fancybox-overlay',
  '.fancybox-wrap',
  '.fancybox-inner',
  '.fancybox-skin',
  '.fancybox-stage',
  '.fancybox-container',
  '.fancybox__container',
  '.fancybox__backdrop',
  '.fancybox__carousel',
  '.fancybox__slide',
  '[role="dialog"]',
  '[aria-modal="true"]',
];
const IMAGE_VIEWER_OVERLAY_SELECTOR = IMAGE_VIEWER_OVERLAY_SELECTORS.join(', ');
const isInsideImageViewerOverlay = node => (
  !!node &&
  typeof node.closest === 'function' &&
  !!node.closest(IMAGE_VIEWER_OVERLAY_SELECTOR)
);
const getSanitizedHtml = root => {
  if (!root) return '';
  try {
    const clone = root.cloneNode(true);
    if (clone && typeof clone.querySelectorAll === 'function') {
      clone.querySelectorAll(IMAGE_VIEWER_OVERLAY_SELECTOR).forEach(node => node.remove());
    }
    return clone?.innerHTML || '';
  } catch {
    return root?.innerHTML || '';
  }
};
const queryAllOutsideImageViewer = selector => Array.from(document.querySelectorAll(selector))
  .filter(node => !isInsideImageViewerOverlay(node));
const queryFirstOutsideImageViewer = (...selectors) => {
  for (const selector of selectors) {
    const found = queryAllOutsideImageViewer(selector)[0];
    if (found) return found;
  }
  return null;
};

  const EXCLUDED_IMAGE_REGION_HEADING = /(?:買了此商品的人也買了|買了此商品的人還買|也買了|猜你喜歡|你可能會喜歡|其他人也買|相關推薦|熱門推薦|最近瀏覽|最近浏览|browse_history|weekly_hot)/;
  const GIFT_PROMO_HEADING = /^(?:特惠贈品|特惠赠品|首刷贈品|首刷赠品|滿額送|满额送)$/;
  const PRODUCT_IMAGE_ROOT_SELECTORS = [
    '.cnt_prod_img001',
    '#item-img-content',
    '.prod_img_box',
    '.item_img',
  ];
  const BOOK_GALLERY_MAX_IMAGES = 9;
  const BOOK_GIFT_MAX_IMAGES = 2;
  const BOOK_DETAIL_MAX_IMAGES = 12;
  const BOOK_MAX_TOTAL_IMAGES = 10;

  const isInsideExcludedImageRegion = node => {
    if (!node || typeof node.closest !== 'function') return false;
    if (node.closest('.mod_recommend, .recommend_box, .reco_box, .also_buy, .prod_recommend, [class*="recommend"], [id*="recommend"], [id*="Recommend"]')) {
      return true;
    }
    let el = node;
    for (let depth = 0; depth < 10 && el; depth += 1) {
      const heading = el.matches?.('h1, h2, h3, h4, .title, .maintitle')
        ? el
        : el.querySelector?.('h1, h2, h3, h4, .title, .maintitle');
      const headingText = toSingleLine(getText(heading));
      if (EXCLUDED_IMAGE_REGION_HEADING.test(headingText)) return true;
      const snippet = toSingleLine(getText(el)).slice(0, 120);
      if (EXCLUDED_IMAGE_REGION_HEADING.test(snippet)) return true;
      el = el.parentElement;
    }
    return false;
  };

  const isInsideGiftPromoImageRegion = node => {
    if (!node || typeof node.closest !== 'function') return false;
    if (node.closest('[id*="getGiftInfo"]')) return true;
    let el = node;
    for (let depth = 0; depth < 10 && el; depth += 1) {
      const heading = toSingleLine(getText(
        el.querySelector?.('h1, h2, h3, h4, .title, .maintitle, strong, dt') || ''
      ));
      if (GIFT_PROMO_HEADING.test(heading)) return true;
      el = el.parentElement;
    }
    return false;
  };

  const isAllowedProductImageNode = node => (
    !isInsideImageViewerOverlay(node)
    && !isInsideExcludedImageRegion(node)
    && !isInsideGiftPromoImageRegion(node)
  );

  const getProductImageRoot = () => {
    for (const selector of PRODUCT_IMAGE_ROOT_SELECTORS) {
      const found = queryAllOutsideImageViewer(selector).find(isAllowedProductImageNode);
      if (found) return found;
    }
    return null;
  };

  const queryWithinProductImageRoot = selector => {
    const root = getProductImageRoot();
    if (!root) return [];
    return Array.from(root.querySelectorAll(selector)).filter(isAllowedProductImageNode);
  };

  const isGalleryImageNode = node => (
    !!node
    && !isInsideImageViewerOverlay(node)
    && !isInsideExcludedImageRegion(node)
  );

  const getGalleryImageRoot = () => {
    for (const selector of PRODUCT_IMAGE_ROOT_SELECTORS) {
      const found = queryAllOutsideImageViewer(selector).find(isGalleryImageNode);
      if (found) return found;
    }
    return null;
  };

  const queryWithinGalleryRoot = selector => {
    const root = getGalleryImageRoot();
    if (!root) return [];
    return Array.from(root.querySelectorAll(selector)).filter(isGalleryImageNode);
  };

const scoreEdgeDistance = edge => {
    if (!Number.isFinite(edge) || edge <= 0) return 0;
    return Math.max(0, 3000 - Math.abs(edge - TARGET_IMAGE_EDGE) * 4);
  };

  const extractRequestedEdge = imageUrl => {
    try {
      const parsed = new URL(String(imageUrl || ''), window.location.href);
      const width = parseInt(parsed.searchParams.get('w') || parsed.searchParams.get('width') || '0', 10);
      const height = parseInt(parsed.searchParams.get('h') || parsed.searchParams.get('height') || '0', 10);
      return Math.max(Number.isFinite(width) ? width : 0, Number.isFinite(height) ? height : 0);
    } catch {
      return 0;
    }
  };
  const collectImageUrlsFromNode = node => {
    if (!node) return [];
    if (!isAllowedProductImageNode(node)) return [];
    const urls = [];
    const pushUrl = raw => {
      const normalized = normalizeImageUrl(raw);
      if (normalized) urls.push(normalized);
    };

    Array.from(node.querySelectorAll('img'))
      .filter(img => !isInsideImageViewerOverlay(img))
      .forEach(img => {
      const linkedHref = img.closest('a[href]')?.getAttribute('href') || '';
      if (/\/products\//i.test(linkedHref) && !new RegExp(productCodeRegexSafe, 'i').test(linkedHref)) {
        return;
      }
      pushUrl(img?.getAttribute('data-original'));
      pushUrl(img?.getAttribute('data-src'));
      pushUrl(img?.getAttribute('data-large'));
      pushUrl(img?.currentSrc);
      pushUrl(img?.src);

      const srcset = img?.getAttribute('srcset') || img?.getAttribute('data-srcset') || '';
      if (srcset) {
        srcset.split(',').forEach(part => {
          const candidate = part.trim().split(/\s+/)[0];
          pushUrl(candidate);
        });
      }
    });

    Array.from(node.querySelectorAll('a[href]'))
      .filter(a => isAllowedProductImageNode(a))
      .forEach(a => {
      const href = a.getAttribute('href') || '';
      if (/\/products\//i.test(href) && !new RegExp(productCodeRegexSafe, 'i').test(href)) {
        return;
      }
      if (/\.(?:jpe?g|webp|png)(?:\?|$)/i.test(href) || /getImage\.php/i.test(href)) {
        pushUrl(href);
      }
    });

    Array.from(node.querySelectorAll('[style*="background-image"]'))
      .filter(el => !isInsideImageViewerOverlay(el))
      .forEach(el => {
      const style = el.getAttribute('style') || '';
      Array.from(style.matchAll(/url\((['"]?)(.*?)\1\)/gi)).forEach(match => {
        pushUrl(match[2]);
      });
    });

    const html = getSanitizedHtml(node);
    Array.from(
      html.matchAll(/https?:\/\/[^"'\\s<)]+?\.(?:jpg|jpeg|webp|png)(?:\?[^"'\\s<)]*)?/gi)
    ).forEach(match => {
      const normalized = normalizeImageUrl(match[0]);
      if (!normalized) return;
      if (!new RegExp(productCodeRegexSafe, 'i').test(normalized)
        && !/(?:addons\.books\.com\.tw\/G\/ADbanner\/|\/G\/ADbanner\/)/i.test(normalized)) {
        return;
      }
      pushUrl(normalized);
    });
    Array.from(
      html.matchAll(/https?:\/\/www\.books\.com\.tw\/fancybox\/getImage\.php\?[^"'\\s<)]*/gi)
    ).forEach(match => {
      const normalized = normalizeImageUrl(match[0]);
      if (!normalized) return;
      if (!new RegExp(productCodeRegexSafe, 'i').test(normalized)
        && !/(?:addons\.books\.com\.tw\/G\/ADbanner\/|\/G\/ADbanner\/)/i.test(normalized)) {
        return;
      }
      pushUrl(normalized);
    });

    return uniq(urls);
  };

  const getInfoListItems = () => {
    const containers = Array.from(document.querySelectorAll(
      'ul.bd-info, .bd-info ul, .prod_cont_b ul, .prod_cont ul, .type02_p002 ul, .type02_p003 ul'
    ));
    const scopedLis = containers.flatMap(ul => Array.from(ul.querySelectorAll('li')));
    return scopedLis.length ? scopedLis : Array.from(document.querySelectorAll('li'));
  };
  const infoListItems = getInfoListItems();

  const isGiftSectionPending = scope => {
    if (!scope) return false;
    const text = toSingleLine(getText(scope));
    if (!text) return true;
    if (/載入中|加载中/.test(text)) return true;
    return !scope.querySelector('img, .item, .cont, .content, li, a[href], table');
  };

  const waitForGiftSectionContent = async () => {
    const sectionSelector = '.mod_a, .mod_b, .mod, .prod_cont, .type01, .type02, .type02_p001, .type02_p002, .type02_p003';
    const collectScopes = () => uniq(
      Array.from(document.querySelectorAll(sectionSelector)).filter(scope => {
        const heading = toSingleLine(getText(scope.querySelector('h1, h2, h3, h4, .title, .maintitle, strong, dt')));
        const text = toSingleLine(getText(scope));
        return /特惠贈品|特惠赠品|首刷贈品|首刷赠品|滿額送|满额送|買就送|买就送|限定特典卡|書籍延伸內容|书籍延伸内容|贈送條件|赠送条件/.test(`${heading}\n${text}`) ||
          !!scope.querySelector('[id*="getGiftInfo"]');
      })
    );

    const initialScopes = collectScopes();
    if (!initialScopes.length) return;

    for (let i = 0; i < 12; i += 1) {
      const scopes = collectScopes();
      if (scopes.length && !scopes.some(isGiftSectionPending)) return;
      await new Promise(resolve => setTimeout(resolve, 250));
    }
  };

  const getLabeledValue = labels => {
    for (const li of infoListItems) {
      const text = toSingleLine(getText(li));
      if (!text) continue;

      for (const label of labels) {
        if (!text.includes(label)) continue;

        const labelPattern = escapeRegExp(label);
        const byLabel = text.match(new RegExp(`${labelPattern}\\s*[：:]\\s*([^\\n]+)`));
        if (byLabel?.[1]) {
          return cleanupLabeledValue(byLabel[1]);
        }

        const fromLinks = cleanupLabeledValue(
          Array.from(li.querySelectorAll('a, span'))
            .map(a => getText(a))
            .join(' ')
        );
        if (fromLinks) return fromLinks;

        const afterColon = cleanupLabeledValue(text.replace(/^[^：:]+[：:]/, ''));
        if (afterColon) return afterColon;
      }
    }
    return '';
  };

  const normalizeIsbnText = value => String(value || '')
    .replace(/[０-９]/g, char => String.fromCharCode(char.charCodeAt(0) - 0xFEE0))
    .replace(/[Ｘｘ]/g, 'X');

  const isValidIsbn13 = value => {
    if (!/^\d{13}$/.test(value)) return false;
    const sum = value.split('').reduce((total, char, index) => (
      total + Number(char) * (index % 2 === 0 ? 1 : 3)
    ), 0);
    return sum % 10 === 0;
  };

  const isValidIsbn = value => value.length === 13 ? isValidIsbn13(value) : false;

  const collectIsbnCandidates = raw => {
    const text = normalizeIsbnText(raw);
    if (!text) return [];

    const candidates = [];
    const seen = new Set();
    const addCandidate = value => {
      const normalized = String(value || '').replace(/[^\d]/g, '');
      if (!/^\d{13}$/.test(normalized)) return;
      if (seen.has(normalized)) return;
      seen.add(normalized);
      candidates.push(normalized);
    };

    Array.from(text.matchAll(/\d{13}/g)).forEach(match => addCandidate(match[0]));
    Array.from(text.matchAll(/[\dXx][\dXx\-\s]{8,24}[\dXx]/g)).forEach(match => addCandidate(match[0]));

    return candidates;
  };

  const pickPreferredIsbn = candidates => {
    const list = (candidates || []).filter(Boolean);
    return list.find(value => value.length === 13 && isValidIsbn13(value))
      || list.find(value => value.length === 13)
      || '';
  };

  const extractIsbn = raw => pickPreferredIsbn(collectIsbnCandidates(raw));

  const collectIsbnCandidatesInObject = node => {
    if (node == null) return [];
    if (typeof node !== 'object') return collectIsbnCandidates(node);
    if (Array.isArray(node)) {
      return node.flatMap(item => collectIsbnCandidatesInObject(item));
    }

    const candidates = [];
    if (node.isbn != null) {
      candidates.push(...collectIsbnCandidates(node.isbn));
    }

    for (const key of Object.keys(node)) {
      candidates.push(...collectIsbnCandidatesInObject(node[key]));
    }
    return candidates;
  };

  const findIsbnInObject = node => pickPreferredIsbn(collectIsbnCandidatesInObject(node));

  const getIsbn = () => {
    const candidates = [];
    const addCandidates = value => {
      candidates.push(...collectIsbnCandidates(value));
    };

    addCandidates(getLabeledValue(['ISBN', 'isbn']));
    addCandidates(document.querySelector('meta[property="books:isbn"]')?.getAttribute('content'));
    addCandidates(document.querySelector('meta[name="description"]')?.getAttribute('content') || '');

    for (const scriptEl of Array.from(document.querySelectorAll('script[type="application/ld+json"]'))) {
      try {
        const parsed = JSON.parse(scriptEl.textContent || '');
        candidates.push(...collectIsbnCandidatesInObject(parsed));
      } catch {
        // ignore invalid JSON-LD
      }
    }

    addCandidates(infoListItems.map(li => getText(li)).join(' '));
    addCandidates(document.body?.innerText || document.documentElement?.textContent || '');
    return pickPreferredIsbn(candidates);
  };

  const getPublisher = () => {
    for (const li of infoListItems) {
      const text = toSingleLine(getText(li));
      if (!/(出版社|出版商)\s*[：:]/.test(text)) continue;

      const bySpan = cleanupLabeledValue(getText(li.querySelector('a span')));
      if (bySpan) return bySpan;

      const byAnchor = cleanupLabeledValue(getText(li.querySelector('a')));
      if (byAnchor) return byAnchor;

      const byLabel = cleanupLabeledValue(text.replace(/^[^：:]+[：:]/, ''));
      if (byLabel) return byLabel;
    }

    return cleanupLabeledValue(
      getText(document.querySelector('a[href*="sys_puballb"][href*="pubid"] span, a[href*="search=publisher"] span, a[href*="search=publisher"]'))
    );
  };


  const getCategoryPath = () => {
    const linkTexts = Array.from(document.querySelectorAll('.sort li a, .type02_p003 a, .breadcrumb a, .crumb a'))
      .map(el => toSingleLine(getText(el)))
      .filter(Boolean);
    return uniq(linkTexts).join(' > ');
  };
  const cleanupOriginalTitleOnce = value => value
    .replace(/^[台臺][湾灣]版\s*/u, '')
    // 1. Bracket noise with edition/bonus keywords (added 獨家/独家/博客來獨家/博客来独家)
    .replace(/\s*【[^】]*(?:首刷|初回|初版|限定|特裝|特装|特別|特别|通常|普通|獨家|独家|博客來獨家|博客来独家)[^】]*】\s*/gu, ' ')
    .replace(/\s*\([^)]*(?:首刷|初回|初版|限定|特裝|特装|特別|特别|通常|普通|獨家|独家|博客來獨家|博客来独家)[^)]*\)\s*/gu, ' ')
    .replace(/\s*\[[^\]]*(?:首刷|初回|初版|限定|特裝|特装|特別|特别|通常|普通|獨家|独家|博客來獨家|博客来独家)[^\]]*\]\s*/gu, ' ')
    .replace(/\s*（[^）]*(?:首刷|初回|初版|限定|特裝|特装|特別|特别|通常|普通|獨家|独家|博客來獨家|博客来独家)[^）]*）\s*/gu, ' ')
    // 2. Bracket with complete/exclusive markers at the very end
    .replace(/\s*[（(\[【]\s*(?:完|全|獨家|独家|博客來獨家|博客来独家|套書|套书|合集)\s*[）)\]】]\s*$/giu, ' ')
    // 3. Volume number followed by complete/exclusive bracket at the very end, e.g. " 3 (完)"
    .replace(/\s*(?:第?\s*)?\d+\s*(?:巻|卷|冊|册|集|話|话|部)?\s*[（(\[【]\s*(?:完|全|獨家|独家|博客來獨家|博客来独家|套書|套书|合集)\s*[）)\]】]\s*$/giu, ' ')
    // 4. Standard bracketed volume numbers at the end
    .replace(/\s*[\(（【\[]\s*(?:第?\s*)?\d+\s*(?:巻|卷|冊|册|集|話|话|部)?\s*[\)）】\]]\s*$/u, ' ')
    .replace(/\s*[-/／]\s*(?:首刷|初回|初版|限定|特裝|特装|特別|特别|通常|普通|獨家|独家)[^ ]*\s*/gu, ' ')
    .replace(/\s*(?:首刷(?:限定)?|初回(?:限定)?|初版(?:限定)?|限定|特裝|特装|特別|特别|通常|普通|獨家|独家)(?:版)?\s*$/u, '')
    .replace(/\s*第?\s*\d+\s*(?:巻|卷|冊|册|集|話|话|部)\s*$/u, '')
    .replace(/\s*\d+\s*$/u, '')
    .replace(/^《\s*([^》]+?)\s*》$/u, '$1')
    .replace(/^「\s*([^」]+?)\s*」$/u, '$1')
    .replace(/^『\s*([^』]+?)\s*』$/u, '$1')
    .replace(/\s+/g, ' ')
    .trim();

  const extractOriginalTitle = rawTitle => {
    const fallback = toSingleLine(rawTitle);
    if (!fallback) return '';

    let normalized = fallback;
    for (let i = 0; i < 4; i += 1) {
      const next = cleanupOriginalTitleOnce(normalized);
      if (next === normalized) break;
      normalized = next;
    }

    return normalized || fallback;
  };

  const stripGoodsMediaTag = raw => toSingleLine(raw)
    .replace(/\s*[（(][^）)]*(?:漫畫|漫画|コミック|小說|小说|小説|動畫|动画|アニメ|影集|ドラマ|電影|映画|劇場版)[^）)]*[)）]\s*/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  const extractGoodsWorkTitle = rawTitle => {
    const fallback = toSingleLine(rawTitle).replace(/^[台臺][湾灣]版\s*/u, '').trim();
    if (!fallback) return '';

    const normalizeCandidate = candidate => stripGoodsMediaTag(extractOriginalTitle(candidate) || candidate) || fallback;

    const mediaTagged = fallback.match(/^(.+?)\s*[（(][^）)]*(?:漫畫|漫画|コミック|小說|小说|小説|動畫|动画|アニメ|影集|ドラマ|電影|映画|劇場版)[^）)]*[)）](?:\s+.*)?$/u);
    if (mediaTagged?.[1]) {
      return normalizeCandidate(mediaTagged[1]);
    }

    const merchTagged = fallback.match(/^(.+?)\s+(?:全棉托特袋|托特袋|帆布袋|手提袋|透明立方|壓克力(?:牌|立牌|磚|砖)?|压克力(?:牌|立牌|磚|砖)?|立牌|海報|海报|掛軸|挂轴|徽章|胸章|卡片|貼紙|贴纸|明信片|抱枕|滑鼠墊|鼠標墊|鼠标垫|桌墊|桌垫|吊飾|吊饰|鑰匙圈|钥匙圈|杯墊|杯垫|色紙|色纸|套組|套装|福袋|拼圖|拼图|公仔|玩偶|資料夾|资料夹|文件夾|文件夹|T恤|毛巾|桌曆|桌历|年曆|年历)(?:\b|\s|$)/u);
    if (merchTagged?.[1]) {
      return normalizeCandidate(merchTagged[1]);
    }

    return normalizeCandidate(fallback);
  };
  const hasJapaneseSubtitleSignal = value => /[ぁ-ゖァ-ヺ々〆ヵヶー]/.test(String(value || ''));
  const normalizeJapaneseSubtitle = raw => toSingleLine(raw)
    .replace(/^(?:日文(?:版)?|日本語(?:タイトル|標題|版)?|日文標題|日文标题|日文書名|日文书名)\\s*[：:]\\s*/u, '')
    .trim();
  const isJapaneseSubtitleCandidate = (candidate, titleText) => {
    const text = normalizeJapaneseSubtitle(candidate);
    if (!text || text.length < 2 || text.length > 120) return false;
    if (!hasJapaneseSubtitleSignal(text)) return false;
    if (text === toSingleLine(titleText)) return false;
    if (/(作者|譯者|译者|翻譯者|翻訳者|出版社|出版商|出版日期|出版时间|發行日期|語言|语言|ISBN|定價|優惠價|优惠价|品牌|規格|规格)\\s*[：:]/.test(text)) return false;
    return true;
  };
  const isJapaneseSubtitlePlausibleForTitle = (candidate, titleText) => {
    const text = normalizeJapaneseSubtitle(candidate);
    if (!text) return false;
    // ブログ來の商品詳細ページから直接スクレイピングしたタイトル下のテキストは極めて信頼性が高いため、
    // 外部検索（MangaUpdates）のような別作品の誤一致リスク検証（漢字の重なりチェックなど）は不要。
    // ひらがな・カタカナ・長音記号などの日本語信号が含まれていれば妥当とみなして採用する。
    return hasJapaneseSubtitleSignal(text);
  };
  const getJapaneseSubtitle = titleText => {
    const titleEl = document.querySelector('h1.BD_NAME, h1[itemprop=name], .book_title h1, h1');
    if (!titleEl) return '';

    // 【強力なショートカット】h1の直後（隣）に h2 があり、そこに日本語が含まれていればダイレクトに採用する！
    const nextEl = titleEl.nextElementSibling;
    if (nextEl && String(nextEl.tagName).toUpperCase() === 'H2') {
      const text = normalizeJapaneseSubtitle(getText(nextEl) || nextEl.innerText || '');
      if (hasJapaneseSubtitleSignal(text) && text !== toSingleLine(titleText)) {
        return text;
      }
    }
    
    // 兄弟要素の中にある h2 をダイレクトに探す
    const parent = titleEl.parentElement;
    if (parent) {
      const h2El = parent.querySelector('h2');
      if (h2El && h2El !== titleEl) {
        const text = normalizeJapaneseSubtitle(getText(h2El) || h2El.innerText || '');
        if (hasJapaneseSubtitleSignal(text) && text !== toSingleLine(titleText)) {
          return text;
        }
      }
    }

    const candidates = [];
    const pushCandidate = raw => {
      const text = normalizeJapaneseSubtitle(raw);
      if (!isJapaneseSubtitleCandidate(text, titleText)) return;
      if (!isJapaneseSubtitlePlausibleForTitle(text, titleText)) return;
      if (!candidates.includes(text)) candidates.push(text);
    };

    if (parent) {
      let afterTitle = false;
      for (const node of Array.from(parent.childNodes)) {
        if (node === titleEl) {
          afterTitle = true;
          continue;
        }
        if (!afterTitle) continue;

        if (node.nodeType === Node.TEXT_NODE) {
          pushCandidate(node.textContent || '');
        } else if (node.nodeType === Node.ELEMENT_NODE) {
          const el = node;
          const tag = String(el.tagName || '').toUpperCase();
          if (/^(UL|OL|TABLE|DL|NAV)$/.test(tag)) break;
          pushCandidate(getText(el));
        }
        if (candidates.length) break;
      }

      Array.from(parent.querySelectorAll('h2, .subtitle, .sub-title, .subTitle, .title02, .book_subtitle, p, div, span'))
        .slice(0, 12)
        .forEach(el => {
          if (el === titleEl || el.contains(titleEl) || titleEl.contains(el)) return;
          pushCandidate(getText(el));
        });
    }

    return candidates[0] || '';
  };
  const parseSizeFromText = raw => {
    const text = String(raw || '');
    const sizeMatch = text.match(/(\d+(?:\.\d+)?\s*[x×＊*]\s*\d+(?:\.\d+)?(?:\s*[x×＊*]\s*\d+(?:\.\d+)?)?\s*(?:cm|mm|公分|厘米))/i);
    if (!sizeMatch?.[1]) return '';
    return toSingleLine(sizeMatch[1])
      .replace(/[×＊*]/g, 'x')
      .replace(/\s*x\s*/g, ' x ');
  };

  const getFormatInfo = () => {
    const format = getLabeledValue(['規格', '规格']);
    let size = parseSizeFromText(format);

    if (!size) {
      const detailBlock = Array.from(document.querySelectorAll('.mod_b')).find(mod =>
        /詳細資料|详细资料/.test(getText(mod.querySelector('h3')))
      );
      size = parseSizeFromText(getText(detailBlock));
    }

    return {
      format: toSingleLine(format),
      size,
    };
  };

  const getImageSlotIndex = imageUrl => {
    const text = String(imageUrl || '');
    const match = text.match(new RegExp(`${productCodeRegexSafe}(?:_(?:b|t)_(\\d+))?\\.(?:jpe?g|png|webp)`, 'i'));
    if (!match) return Number.MAX_SAFE_INTEGER;
    if (!match[1]) return 0;
    const idx = parseInt(match[1], 10);
    return Number.isFinite(idx) ? idx : Number.MAX_SAFE_INTEGER;
  };

  const getImageOrderIndex = imageUrl => getImageSlotIndex(imageUrl);

  /** メイン表示用URL: サムネイル(_t_01等)を優先。なければギャラリー先頭 */
  const pickMainDisplayImageUrl = urls => {
    const list = uniq((urls || []).map(normalizeDisplayImageUrl).filter(Boolean));
    if (!list.length) return '';

    const thumbCandidates = list
      .filter(url => /_t_\d+\.(?:jpe?g|png|webp)(?:[?&#]|$)/i.test(url))
      .sort((a, b) => getImageSlotIndex(a) - getImageSlotIndex(b));
    if (thumbCandidates.length) return thumbCandidates[0];

    const sorted = [...list].sort((a, b) => getImageSlotIndex(a) - getImageSlotIndex(b));
    return sorted[0];
  };

  /** ダウンロード用メイン: 表紙スロットの _b_ 高解像度を優先（_t_ サムネは2枚目に回るのを防ぐ） */
  const pickMainBookDownloadImageUrl = urls => {
    const list = uniq((urls || []).map(normalizeDownloadImageUrl).filter(Boolean));
    if (!list.length) return '';

    const coverCandidates = list
      .filter(url => /_b_\d+\.(?:jpe?g|png|webp)(?:[?&#]|$)/i.test(url))
      .sort((a, b) => {
        const order = getImageSlotIndex(a) - getImageSlotIndex(b);
        if (order !== 0) return order;
        return getImageQualityScore(b) - getImageQualityScore(a);
      });
    if (coverCandidates.length) return coverCandidates[0];

    const sorted = [...list].sort((a, b) => getImageSlotIndex(a) - getImageSlotIndex(b));
    return sorted[0] || '';
  };

  const isSameBookImageSlot = (leftUrl, rightUrl) => {
    if (!leftUrl || !rightUrl) return false;
    const leftSlot = getImageSlotIndex(leftUrl);
    const rightSlot = getImageSlotIndex(rightUrl);
    if (leftSlot < Number.MAX_SAFE_INTEGER && leftSlot === rightSlot) return true;
    return normalizeDownloadImageUrl(leftUrl) === normalizeDownloadImageUrl(rightUrl);
  };

    const getImageQualityScore = imageUrl => {
  const text = String(imageUrl || '');
  let score = 0;

  if (/\/img\/[A-Z]\d{2}\//i.test(text)) score += 1400;
  if (/\/image\/getImage/i.test(text)) score -= 600;

  if (keepGoodsThumbVariant) {
    if (/_t_\d+\.(?:jpe?g|png|webp)(?:[?&#]|$)/i.test(text)) score += 900;
    if (/_b_\d+\.(?:jpe?g|png|webp)(?:[?&#]|$)/i.test(text)) score += 300;
  } else {
    if (/_b_\d+\.(?:jpe?g|png|webp)(?:[?&#]|$)/i.test(text)) score += 900;
    if (/_t_\d+\.(?:jpe?g|png|webp)(?:[?&#]|$)/i.test(text)) score += 200;
  }

  score += scoreEdgeDistance(extractRequestedEdge(text));
  return score;
};

  const sortProductImageUrls = (urls, maxCount = 10) => uniq(urls)
    .sort((a, b) => {
      const orderDiff = getImageOrderIndex(a) - getImageOrderIndex(b);
      if (orderDiff !== 0) return orderDiff;
      return getImageQualityScore(b) - getImageQualityScore(a);
    })
    .slice(0, maxCount);

  const hasCurrentProductCode = text => new RegExp(
    `${productCodeRegexSafe}(?:_(?:b|t)_\\d+|\\.(?:jpe?g|png|webp))`,
    'i'
  ).test(String(text || ''));
  const isImageLikeUrl = imageUrl => /(?:\.jpe?g|\.png|\.webp)(?:[?&#]|$)|\/fancybox\/getImage\.php\?|\/image\/getImage\?/i.test(String(imageUrl || ''));
  const isNoiseImageUrl = imageUrl => keepGoodsThumbVariant
    ? /(?:\/G\/ADbanner\/|\/G\/prod\/comingsoon|languageTest|esqsm|\/reco\/|\/banner\/|\/recommend\/|\/ad\/|listE\.php|listN\.php|\/activity\/|\/combo\/)/i.test(String(imageUrl || ''))
    : /(?:\/G\/ADbanner\/|\/G\/prod\/comingsoon|languageTest|esqsm|\/reco\/|\/banner\/|\/recommend\/|\/ad\/|_t_\d+\.(?:jpg|jpeg|png|webp)|listE\.php|listN\.php|\/activity\/|\/combo\/)/i.test(String(imageUrl || ''));
  const isProductBonusBannerImage = imageUrl => {
    const text = String(imageUrl || '');
    return hasCurrentProductCode(text) && /(?:addons\.books\.com\.tw\/G\/ADbanner\/|\/G\/ADbanner\/|\/banner\/)/i.test(text);
  };
  const isGiftPromoBannerImage = imageUrl => {
    const text = String(imageUrl || '');
    if (!isImageLikeUrl(text)) return false;
    if (isProductBonusBannerImage(text)) return true;
    if (/(?:addons\.books\.com\.tw\/G\/ADbanner\/|\/G\/ADbanner\/|\/image\/getImage\?|fancybox\/getImage\.php)/i.test(text)) {
      return !/(?:\/reco\/|\/recommend\/|listE\.php|listN\.php|\/activity\/combo)/i.test(text);
    }
    if (hasCurrentProductCode(text) && /_(?:b|t)_\d+\./i.test(text)) return false;
    if (/(?:addons\.books\.com\.tw|activity\.books\.com\.tw|im\d+\.book\.com\.tw)/i.test(text)) {
      return !/(?:\/reco\/|\/recommend\/|comingsoon|\/img\/[A-Z]\d{2}\/)/i.test(text);
    }
    return false;
  };
  const isCurrentProductImageUrl = imageUrl => {
    const text = String(imageUrl || '');
    return isImageLikeUrl(text) && hasCurrentProductCode(text) && !isNoiseImageUrl(text);
  };
  const isCurrentProductDetailImage = imageUrl => {
    const text = String(imageUrl || '');
    return isImageLikeUrl(text) && hasCurrentProductCode(text) && (!isNoiseImageUrl(text) || isProductBonusBannerImage(text));
  };
  const collectImageUrlsFromGalleryNode = node => {
    if (!node || !isGalleryImageNode(node)) return [];
    const urls = [];
    const pushUrl = raw => {
      const normalized = normalizeImageUrl(raw);
      if (normalized) urls.push(normalized);
    };

    Array.from(node.querySelectorAll('img'))
      .filter(img => isGalleryImageNode(img))
      .forEach(img => {
        const linkedHref = img.closest('a[href]')?.getAttribute('href') || '';
        if (/\/products\//i.test(linkedHref) && !new RegExp(productCodeRegexSafe, 'i').test(linkedHref)) {
          return;
        }
        pushUrl(img.getAttribute('data-original'));
        pushUrl(img.getAttribute('data-src'));
        pushUrl(img.getAttribute('data-large'));
        pushUrl(img.currentSrc);
        pushUrl(img.src);
        const srcset = img.getAttribute('srcset') || img.getAttribute('data-srcset') || '';
        if (srcset) {
          srcset.split(',').forEach(part => {
            pushUrl(part.trim().split(/\s+/)[0]);
          });
        }
      });

    Array.from(node.querySelectorAll('a[href]'))
      .filter(a => isGalleryImageNode(a))
      .forEach(a => {
        const href = a.getAttribute('href') || '';
        if (/\/products\//i.test(href) && !new RegExp(productCodeRegexSafe, 'i').test(href)) {
          return;
        }
        if (/\.(?:jpe?g|webp|png)(?:\?|$)/i.test(href) || /getImage/i.test(href)) {
          pushUrl(href);
        }
      });

    Array.from(node.querySelectorAll('[style*="background-image"]'))
      .filter(el => isGalleryImageNode(el))
      .forEach(el => {
        const style = el.getAttribute('style') || '';
        Array.from(style.matchAll(/url\((['"]?)(.*?)\1\)/gi)).forEach(match => {
          pushUrl(match[2]);
        });
      });

    const html = getSanitizedHtml(node);
    Array.from(
      html.matchAll(/https?:\/\/[^"'\\s<)]+?\.(?:jpg|jpeg|webp|png)(?:\?[^"'\\s<)]*)?/gi)
    ).forEach(match => {
      const normalized = normalizeImageUrl(match[0]);
      if (!normalized) return;
      if (!new RegExp(productCodeRegexSafe, 'i').test(normalized)) return;
      pushUrl(normalized);
    });
    Array.from(
      html.matchAll(/https?:\/\/www\.books\.com\.tw\/fancybox\/getImage\.php\?[^"'\\s<)]*/gi)
    ).forEach(match => {
      pushUrl(match[0]);
    });

    return uniq(urls);
  };

  const collectImageUrlsFromGiftScope = scope => {
    if (!scope) return [];
    const urls = [];
    const pushUrl = raw => {
      const normalized = normalizeImageUrl(raw);
      if (normalized) urls.push(normalized);
    };

    const isGiftScopeNode = node => (
      !!node
      && !isInsideImageViewerOverlay(node)
      && !isInsideExcludedImageRegion(node)
    );

    Array.from(scope.querySelectorAll('img'))
      .filter(img => isGiftScopeNode(img))
      .forEach(img => {
        pushUrl(img.getAttribute('data-original'));
        pushUrl(img.getAttribute('data-src'));
        pushUrl(img.getAttribute('data-large'));
        pushUrl(img.currentSrc);
        pushUrl(img.src);
      });

    Array.from(scope.querySelectorAll('a[href]'))
      .filter(a => isGiftScopeNode(a))
      .forEach(a => {
        const href = a.getAttribute('href') || '';
        if (/\.(?:jpe?g|webp|png)(?:\?|$)/i.test(href) || /getImage/i.test(href)) {
          pushUrl(href);
        }
      });

    const html = getSanitizedHtml(scope);
    Array.from(
      html.matchAll(/https?:\/\/[^"'\\s<)]+?\.(?:jpg|jpeg|webp|png)(?:\?[^"'\\s<)]*)?/gi)
    ).forEach(match => pushUrl(match[0]));
    Array.from(
      html.matchAll(/https?:\/\/www\.books\.com\.tw\/fancybox\/getImage\.php\?[^"'\\s<)]*/gi)
    ).forEach(match => pushUrl(match[0]));

    return uniq(urls).filter(url => {
      if (!isImageLikeUrl(url)) return false;
      if (isGiftPromoBannerImage(url)) return true;
      const text = String(url || '');
      if (hasCurrentProductCode(text) && /_(?:b|t)_\d+\./i.test(text)) return false;
      return /(?:addons\.books\.com\.tw|activity\.books\.com\.tw)/i.test(text)
        && !/(?:\/reco\/|\/recommend\/|comingsoon)/i.test(text);
    });
  };

  const collectGiftSectionImages = scope => {
    if (!scope) return [];
    return sortProductImageUrls(
      collectImageUrlsFromGiftScope(scope),
      BOOK_GIFT_MAX_IMAGES
    );
  };

  const isBooksGallerySlotUrl = imageUrl => {
    const text = String(imageUrl || '');
    if (!hasCurrentProductCode(text) || !isImageLikeUrl(text)) return false;
    if (isNoiseImageUrl(text) && !isProductBonusBannerImage(text)) return false;
    return /_(?:b|t)_\d+\.(?:jpe?g|png|webp)(?:[?&#]|$)/i.test(text)
      || new RegExp(`/img/[A-Z]\\d{2}/[^"'\\s<)]*${productCodeRegexSafe}\\.(?:jpe?g|png|webp)`, 'i').test(text)
      || new RegExp(`/img/[A-Z]\\d{2}/[^"'\\s<)]*${productCodeRegexSafe}(?:_(?:b|t)_\\d+)?\\.`, 'i').test(text);
  };

  const isBookSectionDetailImage = imageUrl => {
    if (isCurrentProductDetailImage(imageUrl)) return true;
    const text = String(imageUrl || '');
    return isImageLikeUrl(text)
      && /(?:addons\.books\.com\.tw\/G\/ADbanner\/|\/G\/ADbanner\/)/i.test(text)
      && !/(?:\/reco\/|\/recommend\/|\/activity\/|\/combo\/)/i.test(text);
  };

  const isProductGalleryImage = imageUrl => isBooksGallerySlotUrl(imageUrl);
  const isDetailSectionImage = imageUrl => {
    const text = String(imageUrl || '');
    return isImageLikeUrl(text) && (!isNoiseImageUrl(text) || isProductBonusBannerImage(text));
  };
  const getPrice = () => {
    // まずは構造化メタから取得（最も安定）
    const metaPrice =
      document.querySelector('meta[property="product:price:amount"]')?.getAttribute('content') ||
      document.querySelector('meta[itemprop="price"]')?.getAttribute('content') ||
      '';
    if (metaPrice) return metaPrice.replace(/[^\d]/g, '');

    // 「優惠價：85折195元」のような表示は最後の数字を採用
    const promoText = getLabeledValue(['優惠價', '优惠价', '特價', '特价']);
    if (promoText) {
      const nums = promoText.match(/\d[\d,]*/g);
      if (nums?.length) return nums[nums.length - 1].replace(/[^\d]/g, '');
    }

    // DOM候補からフォールバック（割引率85を拾わないよう最後の数字を採用）
    const selectors = [
      '.prod_cont_b strong.price01 b',
      '.price_box .price01 b',
      '.price01 b',
      'strong.price01 b',
      '.price01',
      '.price em',
      '.price',
    ];

    for (const selector of selectors) {
      const text = getText(document.querySelector(selector));
      if (!text) continue;
      const nums = text.match(/\d[\d,]*/g);
      if (nums?.length) return nums[nums.length - 1].replace(/[^\d]/g, '');
    }

    return '';
  };

  const getBookDescriptionBlocks = () => {
    const headingPattern = /^(?:內容簡介|内容简介|商品簡介|商品介紹|商品介绍|內容介紹)$/;
    const blocks = Array.from(document.querySelectorAll('.mod_b, .mod, .prod_cont')).filter(block => {
      const headingText = toSingleLine(getText(block.querySelector('h3, h2, h1')));
      return headingPattern.test(headingText);
    });
    if (blocks.length) return blocks;

    const fallback = document.querySelector('[itemprop="description"], #summary, .prod_summary, .bd .content');
    return fallback ? [fallback] : [];
  };

  const getBookDescription = () => {
    const blocks = getBookDescriptionBlocks();
    const texts = blocks
      .map(block => toMultiline(getText(
        block.querySelector?.('.bd .content, .content, [itemprop="description"]') || block
      )))
      .filter(Boolean);
    if (texts.length) return mergeTextBlocks(texts);

    return toMultiline(getText(
      document.querySelector('[itemprop="description"], #summary, .prod_summary, .bd .content')
    ));
  };

  const getBookDescriptionContentInfo = () => {
    const blocks = getBookDescriptionBlocks();
    if (!blocks.length) return { text: '', images: [] };

    const images = sortProductImageUrls(
      blocks.flatMap(block =>
        collectImageUrlsFromNode(block).filter(isBookSectionDetailImage)
      ),
      BOOK_DETAIL_MAX_IMAGES
    );

    return {
      text: getBookDescription(),
      images,
    };
  };

  const getBookPreviewContentInfo = () => {
    // 「內頁簡介」「試閱」など、書籍プレビュー専用セクションだけを抽出する。
    // 以前は span/p/div/td/strong まで幅広く拾っていて、無関係な領域を picking してしまっていた。
    const sectionSelector = '.mod_a, .mod_b, .mod, .prod_cont, .type01, .type02, .type02_p001, .type02_p002, .type02_p003';
    const headingSelector = 'h1, h2, h3, h4, .title, .maintitle, dt';
    // 見出しテキスト全体が「試閱」等の完全一致 or それに近い場合だけ採用（ノイズ抑止）
    const previewHeadingPattern = /^(?:\s*)(?:內頁簡介|内頁簡介|內頁介绍|内页简介|內页简介|內頁圖|内頁圖|內页图|內頁試閱|内頁試阅|試閱|试阅)(?:\s*[:：]?\s*)$/;
    // 関連商品・おすすめ・暢銷ランキング等のセクションに侵入しないよう停止パターンを追加
    const stopHeadingPattern = /^(?:特惠贈品|特惠赠品|首刷贈品|首刷赠品|書籍延伸內容|书籍延伸内容|商品簡介|商品介绍|商品介紹|內容簡介|内容简介|詳細資料|详细资料|規格|规格|作者|譯者|译者|出版社|主題活動|主题活动|百貨商品推薦|百货商品推荐|買了此商品的人.*|买了此商品的人.*|看了此商品的人.*|看过此商品的人.*|猜你喜歡.*|猜你喜欢.*|你可能會喜歡.*|你可能会喜欢.*|相關商品.*|相关商品.*|同類商品.*|同类商品.*|暢銷排行.*|畅销排行.*|本月強推.*|本月强推.*|延伸閱讀.*|延伸阅读.*)$/;
    const headingNodes = Array.from(
      document.querySelectorAll(headingSelector)
    ).filter(node => previewHeadingPattern.test(toSingleLine(getText(node))));
    const seenSections = new Set();
    const scopes = [];
    const MAX_SIBLING_WALK = 8; // 無制限に nextElementSibling を辿らない

    for (const headingNode of headingNodes) {
      const section = headingNode.closest(sectionSelector);
      if (!section || seenSections.has(section)) continue;
      seenSections.add(section);

      let anchor = headingNode;
      while (anchor?.parentElement && anchor.parentElement !== section) {
        anchor = anchor.parentElement;
      }

      if (anchor) scopes.push(anchor);
      let current = anchor?.nextElementSibling || null;
      let walks = 0;
      while (current && walks < MAX_SIBLING_WALK) {
        const currentHeading = current.matches?.(headingSelector)
          ? toSingleLine(getText(current))
          : toSingleLine(getText(current.querySelector?.(headingSelector)));
        if (previewHeadingPattern.test(currentHeading) || stopHeadingPattern.test(currentHeading)) break;
        scopes.push(current);
        current = current.nextElementSibling;
        walks += 1;
      }

      const fallbackScope = section.querySelector('.bd .content, .content, .cont, .bd');
      if (fallbackScope) scopes.push(fallbackScope);
    }

    const targetScopes = uniq(scopes).filter(Boolean);
    if (!targetScopes.length) return { text: '', images: [] };

    const texts = targetScopes
      .map(scope => toMultiline(getText(scope)).replace(previewHeadingPattern, '').trim())
      .filter(Boolean);
    // 商品コードを含まない画像は別商品の画像なので必ず除外
    const images = sortProductImageUrls(
      targetScopes.flatMap(scope =>
        collectImageUrlsFromNode(scope).filter(isBookSectionDetailImage)
      ),
      BOOK_DETAIL_MAX_IMAGES
    );

    return {
      text: mergeTextBlocks(texts),
      images,
    };
  };

  const cleanGoodsSectionText = raw => toMultiline(raw)
    .replace(/^(?:產品說明|产品说明|商品簡介|商品介绍|商品介紹|商品說明|商品说明)\s*$/gmu, '')
    .trim();

  const getGoodsContentInfo = () => {
    const sectionSelector = '.mod_a, .mod_b, .mod, .prod_cont, .type01, .type02, .type02_p001, .type02_p002, .type02_p003';
    const headingPattern = /產品說明|产品说明|商品簡介|商品介绍|商品介紹|商品說明|商品说明/;
    const stopPattern = /詳細資料|详细资料|規格|规格|運送|購物說明|购物说明|注意事項|注意事项/;
    const blocks = Array.from(document.querySelectorAll(sectionSelector));
    const sections = [];
    let collectFollowing = false;

    for (const block of blocks) {
      const heading = toSingleLine(getText(block.querySelector('h1, h2, h3, h4, .title, .maintitle, strong, dt')));
      const detailImages = collectImageUrlsFromNode(block).filter(isDetailSectionImage);

      if (headingPattern.test(heading)) {
        collectFollowing = true;
        sections.push(block);
        continue;
      }

      if (!collectFollowing) continue;
      if (stopPattern.test(heading)) break;
      if (detailImages.length || cleanGoodsSectionText(getText(block))) sections.push(block);
    }

    const fallbackSection =
      blocks.find(block => /產品說明|产品说明/.test(getText(block.querySelector('h3, h2, h1')))) ||
      blocks.find(block => /商品簡介|商品介绍|商品介紹|商品說明|商品说明/.test(getText(block.querySelector('h3, h2, h1')))) ||
      null;

    const targetSections = sections.length ? uniq(sections) : (fallbackSection ? [fallbackSection] : []);
    const textBlocks = targetSections.map(section => {
      const body = cleanGoodsSectionText(getText(
        section.querySelector('.type01_content, .bd .type01_content, .bd .content, .content, .bd, .cont') || section
      ));
      return body;
    }).filter(Boolean);

    const sectionImages = sortProductImageUrls(
      targetSections.flatMap(section => collectImageUrlsFromNode(section).filter(isDetailSectionImage)),
      20
    );

    const html = getSanitizedHtml(document.documentElement);
    const htmlDetailUrls = uniq(
      Array.from(html.matchAll(
        new RegExp(String.raw`https?:\/\/(?:www\.books\.com\.tw\/img|im\d+\.book\.com\.tw\/image\/getImage\?i=https:\/\/www\.books\.com\.tw\/img)\/(?:N|M)\d{2}\/\d{3}\/\d{2}\/${productCodeRegexSafe}(?:_t_\d+|_b_\d+)?\.(?:jpg|jpeg|png|webp)(?:\?[^"'\s<]*)?`, 'gi')
      ))
        .map(match => normalizeImageUrl(match[0]))
        .filter(isCurrentProductDetailImage)
    );

    return {
      text: mergeTextBlocks(textBlocks),
      images: sortProductImageUrls([...sectionImages, ...htmlDetailUrls], 20),
    };
  };
  const getExtendedContentInfo = () => {
    const bonusTextPattern = /首刷|贈品|赠品|買就送|买就送|限定特典|書卡|书卡|特典卡/;
    const sectionSelector = '.mod_a, .mod_b, .mod, .prod_cont, .type01, .type02, .type02_p001, .type02_p002, .type02_p003';
    const exactHeadingPattern = /^(?:書籍延伸內容|书籍延伸内容|內容延伸|内容延伸)$/;
    const rawHeadingNodes = Array.from(
      document.querySelectorAll('h1, h2, h3, h4, .title, .maintitle, strong, dt, th, td, div, span, p')
    ).filter(node => exactHeadingPattern.test(toSingleLine(getText(node))));
    const headingNodes = rawHeadingNodes.filter(node => !Array.from(node.children || []).some(child =>
      exactHeadingPattern.test(toSingleLine(getText(child)))
    ));
    const seenSections = new Set();
    const sectionEntries = headingNodes.map(node => {
      const section = node.closest(sectionSelector) || node.parentElement;
      if (!section || seenSections.has(section)) return null;
      seenSections.add(section);
      return { headingNode: node, section };
    }).filter(Boolean);
    if (!sectionEntries.length) return { text: '', images: [] };

    const headingSelector = 'h1, h2, h3, h4, .title, .maintitle, strong, dt';
    // 関連商品・おすすめセクションに侵入しないよう停止パターンを追加
    const stopHeadingPattern = /^(?:特惠贈品|特惠赠品|首刷贈品|首刷赠品|商品簡介|商品介绍|商品介紹|內容簡介|内容简介|詳細資料|详细资料|規格|规格|作者|譯者|译者|出版社|主題活動|主题活动|百貨商品推薦|百货商品推荐|買了此商品的人.*|买了此商品的人.*|看了此商品的人.*|看过此商品的人.*|猜你喜歡.*|猜你喜欢.*|你可能會喜歡.*|你可能会喜欢.*|相關商品.*|相关商品.*|同類商品.*|同类商品.*|暢銷排行.*|畅销排行.*|本月強推.*|本月强推.*|延伸閱讀.*|延伸阅读.*)$/;
    const MAX_SIBLING_WALK = 8; // 無制限に nextElementSibling を辿らない（関連商品セクションへの侵入防止）
    const collectExtendedContentScopes = (section, headingNode) => {
      let anchor = headingNode;
      while (anchor?.parentElement && anchor.parentElement !== section) {
        anchor = anchor.parentElement;
      }

      const scopes = [];
      const anchorText = toSingleLine(getText(anchor));
      const headingText = toSingleLine(getText(headingNode));
      const anchorHasImage = !!anchor?.querySelector?.('img, [style*="background-image"], a[href*="getImage"], a[href$=".jpg"], a[href$=".jpeg"], a[href$=".png"], a[href$=".webp"]');
      const anchorHasBodyText = !!anchorText.replace(headingText, '').trim();
      if (anchor && (anchorHasImage || anchorHasBodyText)) {
        scopes.push(anchor);
      }
      let current = anchor?.nextElementSibling || null;
      let walks = 0;
      while (current && walks < MAX_SIBLING_WALK) {
        const currentHeading = current.matches?.(headingSelector)
          ? toSingleLine(getText(current))
          : toSingleLine(getText(current.querySelector?.(headingSelector)));
        if (exactHeadingPattern.test(currentHeading) || stopHeadingPattern.test(currentHeading)) break;
        scopes.push(current);
        current = current.nextElementSibling;
        walks += 1;
      }

      if (!scopes.length) {
        const fallbackScope = section.querySelector('.bd .content, .content, .cont, .bd');
        if (fallbackScope) scopes.push(fallbackScope);
      }
      return scopes;
    };

    const texts = [];
    const imageUrls = [];

    for (const entry of sectionEntries) {
      const section = entry.section;
      const scopes = collectExtendedContentScopes(section, entry.headingNode);
      if (!scopes.length) continue;
      const sectionTitle = toSingleLine(getText(
        entry.headingNode
      ));
      const parts = [];

      for (const scope of scopes) {
        imageUrls.push(...collectImageUrlsFromNode(scope).filter(isBookSectionDetailImage));
        const items = Array.from(scope.children || []).filter(child => child.matches?.('.item'));
        const targetItems = items.length ? items : [scope];

        for (const item of targetItems) {
          const itemTitle = toSingleLine(getText(item.querySelector?.('h4')));
          const contentEl = item.querySelector?.('.cont') || item.querySelector?.('.content') || item;
          let itemBody = toMultiline(getText(contentEl));
          const itemImages = collectImageUrlsFromNode(item).filter(isBookSectionDetailImage);
          if (itemImages.length) {
            imageUrls.push(...itemImages);
          } else {
            imageUrls.push(...collectImageUrlsFromNode(contentEl || item).filter(isBookSectionDetailImage));
          }

          if (bonusTextPattern.test(`${itemTitle}\n${itemBody}`)) {
            continue;
          }

          if (itemTitle) {
            parts.push(itemTitle);
            if (itemBody.startsWith(itemTitle)) {
              itemBody = itemBody.slice(itemTitle.length).trim();
            }
          }
          if (itemBody) parts.push(itemBody);
        }
      }

      if (parts.length) {
        texts.push(mergeTextBlocks([sectionTitle, ...parts]));
      }
    }

    return {
      text: mergeTextBlocks(texts),
      images: sortProductImageUrls(uniq(imageUrls), BOOK_DETAIL_MAX_IMAGES),
    };
  };

  const getFirstEditionBonusInfo = () => {
    const giftHeadingPattern = /特惠贈品|特惠赠品|首刷贈品|首刷赠品|滿額送|满额送|書籍延伸內容|书籍延伸内容/;
    const giftHeadingSelector = 'h1, h2, h3, h4, h5, p, span, strong, dt, th, .title, .maintitle, .hd, .tit';
    const sectionSelector = '.mod_a, .mod_b, .mod, .prod_cont, .type01, .type02, .type02_p001, .type02_p002, .type02_p003';
    const genericGiftTextPattern = /贈品|赠品|滿額送|满额送|買就送|买就送|贈送條件|赠送条件|活動說明|活动说明|注意事項|注意事项|寄送地區限制|寄送地区限制|送完為止|送完为止|已贈完|已赠完|此商品贈品|此商品赠品|剩餘數量|剩余数量|限量|滿\d+元|满\d+元/;

    const resolveGiftScopeHeading = (scope, fallbackHeading = '') => {
      const headingCandidates = Array.from(scope?.querySelectorAll?.(giftHeadingSelector) || [])
        .map(el => toSingleLine(getText(el)))
        .filter(Boolean);
      const preferred = headingCandidates.find(text => /特惠贈品|特惠赠品/.test(text))
        || headingCandidates.find(text => /滿額送|满额送/.test(text))
        || headingCandidates.find(text => giftHeadingPattern.test(text))
        || toSingleLine(fallbackHeading);
      if (preferred) return preferred;
      const scopeLead = toSingleLine(getText(scope)).slice(0, 120);
      if (/特惠贈品|特惠赠品/.test(scopeLead)) return '特惠贈品';
      if (/滿額送|满额送/.test(scopeLead)) return '滿額送';
      return '';
    };

    const buildGiftSectionLabel = (heading, genericBlocks) => {
      if (/特惠贈品|特惠赠品/.test(heading)) return '特惠贈品';
      if (/滿額送|满额送/.test(heading)) return '滿額送';
      if (genericBlocks.some(block => /滿額送|满额送/.test(block))) return '特惠贈品';
      return '';
    };

    const extractBuyGiftSummary = rawText => {
      const collapsed = toSingleLine(rawText);
      const matched = collapsed.match(/買就送\s*[（(]?\s*贈品\s*[)）]?\s*.*?限定特典卡/)?.[0] || '';
      return matched
        .replace(/買就送\s*[（(]\s*贈品\s*[)）]\s*/g, '買就送(贈品)')
        .replace(/\s+/g, ' ')
        .trim();
    };

    const extractFirstEditionDetail = rawText => {
      const collapsed = toSingleLine(rawText);
      return collapsed.match(/首刷限定\s*[：:].*?(?:送完為止|送完为止)(?:\s*[（(][^）)]*[)）])?/)?.[0] || '';
    };

    const cleanupGiftBlockText = rawText => toMultiline(rawText)
      .replace(/^(?:特惠贈品|特惠赠品|首刷贈品|首刷赠品)\s*$/gmu, '')
      .trim();

    const extractGenericGiftBlocks = (scope, heading) => {
      const giftInfoRoot = scope.querySelector?.('[id*="getGiftInfo"]');
      const extractionRoot = giftInfoRoot || scope;
      const items = Array.from(extractionRoot.querySelectorAll?.('.item') || []);
      const targetItems = items.length ? items : [extractionRoot];
      const blocks = [];

      for (const item of targetItems) {
        const itemTitle = toSingleLine(getText(item.querySelector?.('h4, h5, .title, strong, dt')));
        const contentEl = item.querySelector?.('.cont, .content, .bd, .box_2, .box') || item;
        let itemBody = cleanupGiftBlockText(getText(contentEl));

        if (itemTitle && itemBody.startsWith(itemTitle)) {
          itemBody = itemBody.slice(itemTitle.length).trim();
        }

        const itemText = `${itemTitle}\n${itemBody}`.trim();
        const itemImages = collectGiftSectionImages(item);
        if (!genericGiftTextPattern.test(itemText) && !itemImages.length) continue;

        const block = mergeTextBlocks([
          itemTitle,
          itemBody,
        ]);
        if (block) blocks.push(block);
      }

      if (blocks.length) {
        return blocks;
      }

      const fallbackText = cleanupGiftBlockText(getText(scope));
      if (!genericGiftTextPattern.test(`${heading}\n${fallbackText}`)) {
        return [];
      }

      return [fallbackText];
    };

    const headingScopes = Array.from(
      document.querySelectorAll(giftHeadingSelector)
    )
      .filter(el => giftHeadingPattern.test(toSingleLine(getText(el))))
      .map(el => el.closest(sectionSelector) || el.closest('[id*="getGiftInfo"]')?.parentElement || el.parentElement)
      .filter(Boolean);
    const ajaxScopes = Array.from(document.querySelectorAll('[id*="getGiftInfo"]'))
      .map(el => el.closest(sectionSelector) || el.parentElement)
      .filter(Boolean);
    const itemTitleScopes = Array.from(document.querySelectorAll('.item h4, .item h5, .item strong, .item .title, .item dt'))
      .filter(el => /首刷贈品|首刷赠品|滿額送|满额送|買就送|买就送/.test(toSingleLine(getText(el))))
      .map(el => el.closest(sectionSelector) || el.closest('[id*="getGiftInfo"]')?.parentElement || el.parentElement)
      .filter(Boolean);

    const scopes = uniq([
      ...headingScopes,
      ...ajaxScopes,
      ...itemTitleScopes,
    ]);

    const blocks = [];
    const images = [];
    for (const scope of scopes) {
      const heading = resolveGiftScopeHeading(scope);
      const scopeText = toMultiline(getText(scope));
      const scopeSingleLine = toSingleLine(getText(scope));
      if (!scopeText) continue;

      const directBlocks = [];
      const scopeImages = collectGiftSectionImages(scope);

      const buyGiftSummary = extractBuyGiftSummary(scopeSingleLine);
      if (buyGiftSummary) {
        directBlocks.push(buyGiftSummary);
      }

      const firstEditionDetail = extractFirstEditionDetail(scopeSingleLine);
      if (firstEditionDetail) {
        const hasFirstEditionTitle = /首刷贈品|首刷赠品/.test(`${heading}\n${scopeSingleLine}`);
        directBlocks.push(mergeTextBlocks([
          hasFirstEditionTitle ? '首刷贈品' : '',
          firstEditionDetail,
        ]));
      }

      const genericBlocks = extractGenericGiftBlocks(scope, heading);
      const sectionLabel = buildGiftSectionLabel(heading, genericBlocks);
      const blockText = mergeTextBlocks([
        sectionLabel && genericBlocks.length ? sectionLabel : '',
        ...directBlocks,
        ...genericBlocks,
      ]);
      if (!blockText && !scopeImages.length) continue;

      if (blockText) blocks.push(blockText);
      images.push(...scopeImages);
    }

    return {
      text: mergeTextBlocks(blocks),
      images: uniq(images),
    };
  };

  const collectCoverFallbackUrls = () => {
    const coverImg = queryFirstOutsideImageViewer(
      '#item-img-content img',
      '.cnt_prod_img001 .cover_img .cover',
      '.cnt_prod_img001 .cover_img img',
      '.cover img',
      'img[id*="item-img"]'
    );
    if (!coverImg) return [];

    const urls = [];
    const push = raw => {
      const normalized = normalizeDownloadImageUrl(raw);
      if (normalized && isCurrentProductImageUrl(normalized)) urls.push(normalized);
    };
    push(coverImg.getAttribute('data-original'));
    push(coverImg.getAttribute('data-src'));
    push(coverImg.getAttribute('data-large'));
    push(coverImg.closest('a[href]')?.getAttribute('href'));
    push(coverImg.currentSrc);
    push(coverImg.src);
    return uniq(urls);
  };

  const collectGallerySlotUrls = () => {
    const bySlot = new Map();

    const pushCandidate = raw => {
      const normalized = normalizeDownloadImageUrl(raw);
      if (!normalized) return;
      const isGalleryUrl = isBooksGallerySlotUrl(normalized) || isCurrentProductImageUrl(normalized);
      if (!isGalleryUrl) return;
      const slot = getImageSlotIndex(normalized);
      const existing = bySlot.get(slot);
      if (!existing || (/_b_\d+\./i.test(normalized) && !/_b_\d+\./i.test(existing))) {
        bySlot.set(slot, normalized);
      }
    };

    const pushFromNode = node => {
      if (!node || !isGalleryImageNode(node)) return;
      [
        'data-original',
        'data-src',
        'data-large',
        'data-img',
        'data-image',
        'data-l',
      ].forEach(attr => pushCandidate(node.getAttribute?.(attr)));
      collectImageUrlsFromGalleryNode(node).forEach(pushCandidate);
    };

    const galleryRoot = getGalleryImageRoot();
    const thumbSelectors = '#thumbnail li, .li_box li, .each_box li, .items li, #thumbnail a, .li_box a';
    const thumbScopes = galleryRoot
      ? Array.from(galleryRoot.querySelectorAll(thumbSelectors)).filter(isGalleryImageNode)
      : queryWithinGalleryRoot(thumbSelectors);

    thumbScopes.forEach(pushFromNode);

    queryWithinGalleryRoot(
      '#thumbnail img, .li_box img, .each_box img, .items img, #thumbnail li img'
    ).forEach(img => {
      pushCandidate(img.getAttribute('data-original'));
      pushCandidate(img.getAttribute('data-src'));
      pushCandidate(img.getAttribute('data-large'));
      pushCandidate(img.currentSrc);
      pushCandidate(img.src);
      pushCandidate(img.closest('a[href]')?.getAttribute('href'));
    });

    queryWithinGalleryRoot('#thumbnail, .li_box, .each_box, .items').forEach(scope => {
      pushFromNode(scope);
      scope.querySelectorAll('li, a, img').forEach(pushFromNode);
    });

    queryWithinGalleryRoot(
      '.cover_img .cover, .cover_img img, img.cover, img[id*="item-img"]'
    ).forEach(img => {
      pushCandidate(img.getAttribute('data-original'));
      pushCandidate(img.getAttribute('data-src'));
      pushCandidate(img.getAttribute('data-large'));
      pushCandidate(img.currentSrc);
      pushCandidate(img.src);
      pushCandidate(img.closest('a[href]')?.getAttribute('href'));
    });

    if (!bySlot.size) {
      collectCoverFallbackUrls().forEach(pushCandidate);
    }

    return [...bySlot.entries()]
      .sort((a, b) => a[0] - b[0])
      .map(([, url]) => url)
      .slice(0, BOOK_GALLERY_MAX_IMAGES);
  };

  const getGalleryImageUrls = () => collectGallerySlotUrls();

  const dedupeBookImageUrls = (urls, maxCount = BOOK_MAX_TOTAL_IMAGES) => {
    const byKey = new Map();
    for (const url of uniq(urls || [])) {
      if (!url) continue;
      const slot = getImageSlotIndex(url);
      const isGiftBanner = isGiftPromoBannerImage(url);
      const key = isGiftBanner
        ? `gift:${String(url).replace(/[?#].*$/, '').toLowerCase()}`
        : (slot < Number.MAX_SAFE_INTEGER
          ? `slot:${slot}`
          : String(url).replace(/[?#].*$/, '').toLowerCase());
      const existing = byKey.get(key);
      if (!existing || getImageQualityScore(url) > getImageQualityScore(existing)) {
        byKey.set(key, url);
      }
    }
    return sortProductImageUrls([...byKey.values()], maxCount);
  };

  const getBookGiftPromoImages = (giftImages = []) => sortProductImageUrls(
    uniq((giftImages || []).map(normalizeDownloadImageUrl).filter(isGiftPromoBannerImage)),
    BOOK_GIFT_MAX_IMAGES
  );

  const getBookImages = (options = {}) => {
    const galleryUrls = dedupeBookImageUrls(
      getGalleryImageUrls(),
      BOOK_GALLERY_MAX_IMAGES
    );
    const giftUrls = getBookGiftPromoImages(options.giftImages)
      .filter(url => !galleryUrls.some(galleryUrl => isSameBookImageSlot(galleryUrl, url)));

    const pageCoverDisplay = normalizeDisplayImageUrl(options.displayCoverUrl || '')
      || normalizeDisplayImageUrl(collectCoverFallbackUrls()[0] || '');
    const pageCoverDownload = normalizeDownloadImageUrl(options.displayCoverUrl || '')
      || normalizeDownloadImageUrl(collectCoverFallbackUrls()[0] || '');

    const resolvedMainDownload = pageCoverDownload
      || pickMainBookDownloadImageUrl(galleryUrls)
      || normalizeDownloadImageUrl(galleryUrls[0])
      || '';
    const resolvedMainDisplay = pageCoverDisplay
      || normalizeDisplayImageUrl(resolvedMainDownload)
      || pickMainDisplayImageUrl(galleryUrls)
      || '';

    const galleryAdditional = galleryUrls
      .filter(url => url && !isSameBookImageSlot(url, resolvedMainDownload))
      .sort((a, b) => getImageOrderIndex(a) - getImageOrderIndex(b));

    const maxGalleryAdditional = Math.max(0, BOOK_MAX_TOTAL_IMAGES - 1 - giftUrls.length);
    const additionalImageUrls = [
      ...galleryAdditional.slice(0, maxGalleryAdditional),
      ...giftUrls,
    ].filter(Boolean);

    const all = dedupeBookImageUrls([
      ...(resolvedMainDownload ? [resolvedMainDownload] : []),
      ...additionalImageUrls,
    ], BOOK_MAX_TOTAL_IMAGES);

    return {
      main: resolvedMainDownload,
      displayMain: resolvedMainDisplay,
      gallery: galleryUrls,
      gift: giftUrls,
      all,
      additional: additionalImageUrls,
    };
  };

  const getGoodsImages = (extraDetailUrls = []) => {
    const galleryUrls = getGalleryImageUrls();
    const detailUrls = sortProductImageUrls(extraDetailUrls, 20)
      .filter(url => !galleryUrls.includes(url));
    const main = pickMainDisplayImageUrl(galleryUrls)
      || normalizeDisplayImageUrl(galleryUrls[0])
      || detailUrls[0]
      || '';

    return {
      main,
      gallery: galleryUrls,
      detail: detailUrls,
      all: uniq([
        ...(main ? [main] : []),
        ...galleryUrls.filter(url => url !== main),
        ...detailUrls.filter(url => url !== main),
      ]),
    };
  };

  // ===== 共通項目 =====

  // 商品名
  const titleEl = document.querySelector('h1.BD_NAME, h1[itemprop=name], .book_title h1, h1');
  const name = titleEl?.innerText?.trim() || '';
  const japaneseSubtitle = getJapaneseSubtitle(name);

  // 価格（割引後優先）
  const price = getPrice();

  // 画像URL（メイン画像）
  const imgEl = queryFirstOutsideImageViewer(
    '#item-img-content img',
    '.cnt_prod_img001 .cover_img .cover',
    '.cnt_prod_img001 .cover_img img',
    '.cover img',
    'img[id*="item-img"]'
  );
  const fallbackImageUrl = normalizeDisplayImageUrl(
    imgEl?.getAttribute('data-original') ||
    imgEl?.getAttribute('data-src') ||
    imgEl?.getAttribute('data-large') ||
    imgEl?.closest('a[href]')?.getAttribute('href') ||
    imgEl?.currentSrc ||
    imgEl?.src ||
    ''
  );

  // 発売日
  let releaseDate = getLabeledValue(['出版日期', '出版时间', '發行日期', '発売日']);
  const releaseMatch = releaseDate.match(/(\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}|\d{4}年\d{1,2}月\d{1,2}日)/);
  if (releaseMatch) releaseDate = releaseMatch[1];

  await waitForGiftSectionContent();

  let description = '';
  let mainImageUrl = fallbackImageUrl;
  let additionalImageUrls = [];
  let bonusInfoText = '';
  let formatSize = '';
  let originalTitle = extractOriginalTitle(name);
  let goodsWorkTitle = '';

  if (isBook) {
    const baseDescription = getBookDescription();
    const firstEditionBonus = getFirstEditionBonusInfo();
    const formatInfo = getFormatInfo();
    const bookImages = getBookImages({
      giftImages: firstEditionBonus.images,
      displayCoverUrl: fallbackImageUrl,
    });

    description = baseDescription;
    bonusInfoText = firstEditionBonus.text;
    formatSize = formatInfo.format || formatInfo.size;
    originalTitle = extractOriginalTitle(name);
    mainImageUrl = bookImages.displayMain
      || fallbackImageUrl
      || normalizeDisplayImageUrl(bookImages.main)
      || bookImages.main
      || '';
    additionalImageUrls = (bookImages.additional || [])
      .filter(imgUrl => imgUrl && imgUrl !== mainImageUrl)
      .slice(0, BOOK_MAX_TOTAL_IMAGES - 1);
  } else {
    const brand = getLabeledValue(['品牌']);
    const goodsContent = getGoodsContentInfo();
    const goodsImages = getGoodsImages(goodsContent.images);

    goodsWorkTitle = extractGoodsWorkTitle(name);
    originalTitle = goodsWorkTitle || extractOriginalTitle(name);
    description = mergeTextBlocks([brand ? `品牌：${brand}` : '', goodsContent.text]);
    mainImageUrl = goodsImages.main || goodsContent.images[0] || fallbackImageUrl;
    additionalImageUrls = uniq([
      ...goodsImages.gallery.filter(imgUrl => imgUrl && imgUrl !== mainImageUrl),
      ...goodsImages.detail.filter(imgUrl => imgUrl && imgUrl !== mainImageUrl),
      ...goodsContent.images.filter(imgUrl => imgUrl && imgUrl !== mainImageUrl),
    ]).slice(0, 20);
  }

  const language = getLabeledValue(['語言', '语言']);
  const categoryPath = getCategoryPath();
  const rawPageTitle = document.title || '';
  const pageTitle = rawPageTitle
    .replace(/[-|｜–—]\s*博客來.*$/i, '')
    .replace(/[-|｜–—]\s*books\.com\.tw.*$/i, '')
    .trim() || name;

  const result = {
    商品コード: productCode,
    商品名: name,
    ページタイトル: pageTitle,
    日本語タイトル: japaneseSubtitle,
    原題タイトル: originalTitle,
    作品名原題: goodsWorkTitle,
    価格: price,
    画像URL: mainImageUrl,
    追加画像URL: additionalImageUrls.join(';'),
    発売日: releaseDate,
    商品説明: description,
    特典情報: bonusInfoText,
    規格サイズ: formatSize,
    URL: url,
    種別: isBook ? '書籍' : 'グッズ',
    言語: language,
    カテゴリ: categoryPath,
  };

  // ===== 書籍のみ =====
  if (isBook) {
    // ISBN
    const isbn = getIsbn();
    result['ISBN'] = isbn;

    // 著者
    let author = '';
    const authorEl = document.querySelector('a[href*="search=author"], .BD_AUTHOR a, li a[href*="author"]');
    if (authorEl) {
      author = authorEl.innerText.trim();
    } else {
      author = getLabeledValue(['作者', '著者', '原著']);
    }
    result['著者'] = author;

    // 翻訳者
    const translator = getLabeledValue(['譯者', '译者', '翻譯者', '翻訳者']);
    result['翻訳者'] = translator;

    // イラストレーター
    const illustrator = getLabeledValue(['繪者', '绘者', '插畫', '插画', 'イラスト']);
    result['イラストレーター'] = illustrator;

    // 出版社
    let publisher = getPublisher();
    result['出版社'] = publisher;

    const bookDescriptionMetadata = mergeTextBlocks([
      author ? '作者：' + author : '',
      translator ? '譯者：' + translator : '',
      illustrator ? 'イラストレーター：' + illustrator : '',
      publisher ? '出版社：' + publisher : '',
      releaseDate ? '出版日期：' + releaseDate : '',
      formatSize ? '規格：' + formatSize : '',
    ]);
    result['商品説明'] = mergeTextBlocks([
      result['商品説明'] || '',
      bookDescriptionMetadata,
    ]);
  }

  return result;
}

// popup.jsからのメッセージを受け取る
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'getProductInfo') {
    Promise.resolve(getProductInfo())
      .then(info => {
        sendResponse({ success: true, data: info });
      })
      .catch(e => {
        sendResponse({ success: false, error: e.message });
      });
  }
  return true;
});





















