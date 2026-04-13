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
      return upgradeBooksImageVariant(originalUrl, { keepThumbVariant: keepGoodsThumbVariant });
    }
    return upgradeBooksImageVariant(parsed.href, { keepThumbVariant: keepGoodsThumbVariant });
  } catch {
    return upgradeBooksImageVariant(urlText, { keepThumbVariant: keepGoodsThumbVariant });
  }
};

const normalizeImageUrl = rawUrl => {
  if (!rawUrl) return '';
  return promoteLargeImageUrl(rawUrl);
};

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
    if (isInsideImageViewerOverlay(node)) return [];
    const urls = [];
    const pushUrl = raw => {
      const normalized = normalizeImageUrl(raw);
      if (normalized) urls.push(normalized);
    };

    Array.from(node.querySelectorAll('img'))
      .filter(img => !isInsideImageViewerOverlay(img))
      .forEach(img => {
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
      .filter(a => !isInsideImageViewerOverlay(a))
      .forEach(a => {
      const href = a.getAttribute('href') || '';
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
    ).forEach(match => pushUrl(match[0]));
    Array.from(
      html.matchAll(/https?:\/\/www\.books\.com\.tw\/fancybox\/getImage\.php\?[^"'\\s<)]*/gi)
    ).forEach(match => pushUrl(match[0]));

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
        return /特惠贈品|特惠赠品|首刷贈品|首刷赠品|買就送|买就送|限定特典卡|書籍延伸內容|书籍延伸内容/.test(`${heading}\n${text}`) ||
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
  const extractOriginalTitle = rawTitle => {
    const fallback = toSingleLine(rawTitle);
    if (!fallback) return '';

    const normalized = fallback
      .replace(/^[台臺][湾灣]版\s*/u, '')
      .replace(/\s*【[^】]*(?:首刷|初回|限定|特裝|特装|通常|普通)[^】]*】\s*/gu, ' ')
      .replace(/\s*\([^)]*(?:首刷|初回|限定|特裝|特装|通常|普通)[^)]*\)\s*/gu, ' ')
      .replace(/\s*\[[^\]]*(?:首刷|初回|限定|特裝|特装|通常|普通)[^\]]*\]\s*/gu, ' ')
      .replace(/\s*（[^）]*(?:首刷|初回|限定|特裝|特装|通常|普通)[^）]*）\s*/gu, ' ')
      .replace(/\s*[-/／]\s*(?:首刷|初回|限定|特裝|特装|通常|普通)[^ ]*\s*/gu, ' ')
      .replace(/\s*第?\s*\d+\s*(?:巻|卷|冊|册|集|話|话|部)\s*$/u, '')
      .replace(/\s+\d+\s*$/u, '')
      .replace(/\s+/g, ' ')
      .trim();

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
  const getJapaneseSubtitle = titleText => {
    const titleEl = document.querySelector('h1.BD_NAME, h1[itemprop=name], .book_title h1, h1');
    if (!titleEl) return '';

    const candidates = [];
    const pushCandidate = raw => {
      const text = normalizeJapaneseSubtitle(raw);
      if (!isJapaneseSubtitleCandidate(text, titleText)) return;
      if (!candidates.includes(text)) candidates.push(text);
    };

    const parent = titleEl.parentElement;
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

  const getImageOrderIndex = imageUrl => {
    const text = String(imageUrl || '');
    const match = text.match(new RegExp(`${productCodeRegexSafe}(?:_(?:b|t)_(\\d+))?\\.(?:jpe?g|png|webp)`, 'i'));
    if (!match) return Number.MAX_SAFE_INTEGER;
    if (!match[1]) return 0;
    const idx = parseInt(match[1], 10);
    return Number.isFinite(idx) ? idx : Number.MAX_SAFE_INTEGER;
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

  const hasCurrentProductCode = text => new RegExp(productCodeRegexSafe, 'i').test(String(text || ''));
  const isImageLikeUrl = imageUrl => /(?:\.jpe?g|\.png|\.webp)(?:[?&#]|$)|\/fancybox\/getImage\.php\?|\/image\/getImage\?/i.test(String(imageUrl || ''));
  const isNoiseImageUrl = imageUrl => keepGoodsThumbVariant
    ? /(?:\/G\/ADbanner\/|\/G\/prod\/comingsoon|languageTest|esqsm|\/reco\/|\/banner\/|\/recommend\/|\/ad\/|listE\.php|listN\.php|\/activity\/|\/combo\/)/i.test(String(imageUrl || ''))
    : /(?:\/G\/ADbanner\/|\/G\/prod\/comingsoon|languageTest|esqsm|\/reco\/|\/banner\/|\/recommend\/|\/ad\/|_t_\d+\.(?:jpg|jpeg|png|webp)|listE\.php|listN\.php|\/activity\/|\/combo\/)/i.test(String(imageUrl || ''));
  const isProductBonusBannerImage = imageUrl => {
    const text = String(imageUrl || '');
    return hasCurrentProductCode(text) && /(?:addons\.books\.com\.tw\/G\/ADbanner\/|\/G\/ADbanner\/|\/banner\/)/i.test(text);
  };
  const isCurrentProductImageUrl = imageUrl => {
    const text = String(imageUrl || '');
    return isImageLikeUrl(text) && hasCurrentProductCode(text) && !isNoiseImageUrl(text);
  };
  const isCurrentProductDetailImage = imageUrl => {
    const text = String(imageUrl || '');
    return isImageLikeUrl(text) && hasCurrentProductCode(text) && (!isNoiseImageUrl(text) || isProductBonusBannerImage(text));
  };
  const collectGiftSectionImages = scope => {
    if (!scope) return [];

    const productSpecific = collectImageUrlsFromNode(scope).filter(isCurrentProductDetailImage);
    if (productSpecific.length) return productSpecific;

    const candidates = [];
    const pushCandidate = (rawUrl, width = 0, height = 0) => {
      const normalized = normalizeImageUrl(rawUrl);
      if (!normalized || !isImageLikeUrl(normalized)) return;

      const w = Number(width) || 0;
      const h = Number(height) || 0;
      if (Math.max(w, h) < 100 || Math.min(w, h) < 60) return;
      if (/(?:logo|icon|spinner|loading|blank|share)/i.test(normalized)) return;
      if (isNoiseImageUrl(normalized) && !/(?:addons\.books\.com\.tw\/G\/ADbanner\/|\/G\/ADbanner\/)/i.test(normalized)) return;

      candidates.push({ url: normalized, area: w * h });
    };

    Array.from(scope.querySelectorAll('img')).forEach(img => {
      const width = parseInt(img.getAttribute('width'), 10) || img.naturalWidth || img.width || 0;
      const height = parseInt(img.getAttribute('height'), 10) || img.naturalHeight || img.height || 0;
      pushCandidate(img.getAttribute('data-original'), width, height);
      pushCandidate(img.getAttribute('data-src'), width, height);
      pushCandidate(img.getAttribute('data-large'), width, height);
      pushCandidate(img.currentSrc, width, height);
      pushCandidate(img.src, width, height);
    });

    return uniq(
      candidates
        .sort((a, b) => b.area - a.area)
        .map(item => item.url)
    ).slice(0, 3);
  };

  const isProductGalleryImage = imageUrl => {
    const text = String(imageUrl || '');
    return !isNoiseImageUrl(text) &&
      new RegExp(`${productCodeRegexSafe}(?:_(?:b|t)_\\d+)?\\.(?:jpe?g|png|webp)(?:[?&#]|$)`, 'i').test(text);
  };
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

  const getBookDescription = () => {
    const headingPattern = /內容簡介|内容简介|商品簡介|商品介紹|商品介绍|內容介紹/;
    const blocks = Array.from(document.querySelectorAll('.mod_b, .mod, .prod_cont')).filter(block => {
      const headingText = getText(block.querySelector('h3, h2'));
      return headingPattern.test(headingText);
    });

    const texts = blocks
      .map(block => toMultiline(getText(
        block.querySelector('.bd .content, .content, [itemprop="description"]')
      )))
      .filter(Boolean);
    if (texts.length) return mergeTextBlocks(texts);

    return toMultiline(getText(
      document.querySelector('[itemprop="description"], #summary, .prod_summary, .bd .content')
    ));
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
    const sections = Array.from(document.querySelectorAll('.mod_b, .mod, .prod_cont')).filter(mod =>
      /書籍延伸內容|内容延伸/.test(getText(mod.querySelector('h3, h2')))
    );
    if (!sections.length) return { text: '', images: [] };

    const texts = [];
    const imageUrls = [];

    for (const section of sections) {
      const sectionTitle = toSingleLine(getText(section.querySelector('h3, h2')));
      const parts = [];
      const items = Array.from(section.querySelectorAll('.item'));

      for (const item of items) {
        const itemTitle = toSingleLine(getText(item.querySelector('h4')));
        const contentEl = item.querySelector('.cont') || item.querySelector('.content') || item;
        let itemBody = toMultiline(getText(contentEl));

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

      if (!parts.length) {
        const fallback = toMultiline(getText(section.querySelector('.bd .content, .content, .cont, .bd')));
        if (fallback && !bonusTextPattern.test(fallback)) parts.push(fallback);
      }

      if (parts.length) {
        texts.push(mergeTextBlocks([sectionTitle, ...parts]));
      }
      imageUrls.push(...collectImageUrlsFromNode(section).filter(isCurrentProductDetailImage));
    }

    return {
      text: mergeTextBlocks(texts),
      images: uniq(imageUrls),
    };
  };

  const getFirstEditionBonusInfo = () => {
    const giftHeadingPattern = /特惠贈品|特惠赠品|首刷贈品|首刷赠品|書籍延伸內容|书籍延伸内容/;
    const sectionSelector = '.mod_a, .mod_b, .mod, .prod_cont, .type01, .type02, .type02_p001, .type02_p002, .type02_p003';

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

    const headingScopes = Array.from(
      document.querySelectorAll('h1, h2, h3, h4, .title, .maintitle, strong, dt')
    )
      .filter(el => giftHeadingPattern.test(toSingleLine(getText(el))))
      .map(el => el.closest(sectionSelector) || el.parentElement)
      .filter(Boolean);
    const ajaxScopes = Array.from(document.querySelectorAll('[id*="getGiftInfo"]'))
      .map(el => el.closest(sectionSelector) || el.parentElement)
      .filter(Boolean);
    const itemTitleScopes = Array.from(document.querySelectorAll('.item h4, .item strong, .item .title'))
      .filter(el => /首刷贈品|首刷赠品|買就送|买就送/.test(toSingleLine(getText(el))))
      .map(el => el.closest(sectionSelector) || el.parentElement)
      .filter(Boolean);

    const scopes = uniq([
      ...headingScopes,
      ...ajaxScopes,
      ...itemTitleScopes,
    ]);

    const blocks = [];
    const images = [];
    for (const scope of scopes) {
      const heading = toSingleLine(getText(scope.querySelector('h1, h2, h3, h4, .title, .maintitle, strong, dt')));
      const scopeText = toMultiline(getText(scope));
      const scopeSingleLine = toSingleLine(getText(scope));
      if (!scopeText) continue;

      const directBlocks = [];

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

      if (!directBlocks.length) continue;

      const blockText = mergeTextBlocks(directBlocks);
      if (!blockText) continue;

      blocks.push(blockText);
      images.push(...collectGiftSectionImages(scope));
    }

    return {
      text: mergeTextBlocks(blocks),
      images: uniq(images),
    };
  };

  const getGalleryImageUrls = () => {
    const MAX_GALLERY_IMAGES = 10;
    const sortAndLimit = urls => sortProductImageUrls(urls, MAX_GALLERY_IMAGES);

    const thumbImgs = queryAllOutsideImageViewer(
      '#thumbnail img, .cnt_prod_img001 #thumbnail img, .cnt_prod_img001 .li_box img, .cnt_prod_img001 .each_box img, .cnt_prod_img001 .items img, .prod_img #thumbnail img'
    );
    const thumbUrls = uniq(
      thumbImgs
        .map(img => img?.getAttribute('data-original') || img?.getAttribute('data-src') || img?.currentSrc || img?.src || '')
        .map(normalizeImageUrl)
    ).filter(isProductGalleryImage);

    const thumbScopes = queryAllOutsideImageViewer(
      '#thumbnail, .cnt_prod_img001 #thumbnail, .cnt_prod_img001 .li_box, .cnt_prod_img001 .each_box, .cnt_prod_img001 .items, .prod_img #thumbnail'
    );
    const thumbScopeUrls = uniq(
      thumbScopes.flatMap(scope => collectImageUrlsFromNode(scope))
    ).filter(isProductGalleryImage);

    const coverImgs = queryAllOutsideImageViewer(
      '.cnt_prod_img001 .cover_img .cover, .cnt_prod_img001 .cover_img img, #item-img-content img.cover, .prod_img img.cover, .cover_img img'
    );
    const coverUrls = uniq(
      coverImgs
        .map(img => img?.getAttribute('data-original') || img?.getAttribute('data-src') || img?.currentSrc || img?.src || '')
        .map(normalizeImageUrl)
    ).filter(isProductGalleryImage);

    const domGalleryUrls = sortAndLimit([...thumbUrls, ...thumbScopeUrls, ...coverUrls]);
    if (domGalleryUrls.length > 1) {
      return domGalleryUrls;
    }

    const html = getSanitizedHtml(document.documentElement);
    const rawHtmlUrls = Array.from(
      html.matchAll(new RegExp(`https?:\\/\\/(?:www\\.books\\.com\\.tw\\/img|im\\d+\\.book\\.com\\.tw\\/image\\/getImage\\?i=https:\\/\\/www\\.books\\.com\\.tw\\/img)\\/[^"'\\s<)]*${productCodeRegexSafe}(?:_(?:b|t)_\\d+)?\\.(?:jpg|jpeg|png|webp)(?:\\?[^"'\\s<)]*)?`, 'gi'))
    ).map(match => normalizeImageUrl(match[0]));
    const htmlUrls = sortAndLimit(rawHtmlUrls.filter(isProductGalleryImage));
    if (htmlUrls.length > 1) {
      return htmlUrls;
    }

    const gallerySection = queryFirstOutsideImageViewer(
      '.cnt_prod_img001',
      '#item-img-content',
      '.prod_img_box',
      '.item_img'
    );
    const sectionUrls = gallerySection
      ? collectImageUrlsFromNode(gallerySection).filter(isProductGalleryImage)
      : [];

    const sectionMerged = sortAndLimit([...domGalleryUrls, ...htmlUrls, ...sectionUrls]);
    if (sectionMerged.length > 1) {
      return sectionMerged;
    }

    return sectionMerged;
  };
  const getBookImages = () => {
    const galleryUrls = getGalleryImageUrls();
    const detailUrls = [];

    const allJpgUrls = uniq([...galleryUrls, ...detailUrls]);

    return {
      main: galleryUrls[0] || detailUrls[0] || '',
      gallery: galleryUrls,
      detail: detailUrls,
      all: allJpgUrls,
    };
  };

  const getGoodsImages = (extraDetailUrls = []) => {
    const galleryUrls = getGalleryImageUrls();
    const detailUrls = sortProductImageUrls(extraDetailUrls, 20)
      .filter(url => !galleryUrls.includes(url));
    const main = galleryUrls[0] || detailUrls[0] || '';

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
  const fallbackImageUrl = normalizeImageUrl(
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
    const extendedContent = getExtendedContentInfo();
    const firstEditionBonus = getFirstEditionBonusInfo();
    const formatInfo = getFormatInfo();
    const bookImages = getBookImages();

    description = baseDescription;
    bonusInfoText = firstEditionBonus.text;
    formatSize = formatInfo.format || formatInfo.size;
    originalTitle = extractOriginalTitle(name);
    mainImageUrl = bookImages.main || fallbackImageUrl;
    additionalImageUrls = uniq([
      ...bookImages.gallery.slice(1),
      ...bookImages.detail,
      ...extendedContent.images,
      ...firstEditionBonus.images,
    ]).filter(imgUrl => imgUrl && imgUrl !== mainImageUrl);
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





















