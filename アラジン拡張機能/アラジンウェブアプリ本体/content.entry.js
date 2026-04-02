(() => {
  const state = globalThis.__ALADIN_CONTENT__;
  if (!state) {
    return;
  }
  if (globalThis.__ALADIN_SCRAPER__?.version === state.version) {
    return;
  }

  state.SCRAPER_RULES.default = {
    descriptionSelectors: state.SHARED_DESCRIPTION_SELECTORS,
    descriptionPlans: [
      {
        selectors: ['.pContent[id$="_Introduce"]'],
        minLength: 20,
        maxLength: 2000
      },
      {
        selectors: state.SHARED_DESCRIPTION_SELECTORS,
        excludeSectionLabels: state.DESCRIPTION_EXCLUDE_SECTION_LABELS,
        minLength: 20,
        maxLength: 2000
      }
    ],
    basicInfoSelectors: state.SHARED_BASIC_INFO_SELECTORS,
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
        excludeSectionLabels: state.DEFAULT_EXCLUDE_SECTION_LABELS,
        maxCount: 4
      }
    ],
    excludeAncestorSelectors: state.SHARED_IMAGE_EXCLUDE_ANCESTORS,
    excludeSectionLabels: state.DEFAULT_EXCLUDE_SECTION_LABELS,
    allowGenericFallback: true,
    fallbackMaxCount: 4
  };

  function normalizeGenre(genre) {
    return state.SCRAPER_RULES[genre] ? genre : 'default';
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
    const rule = state.SCRAPER_RULES[genre] || state.SCRAPER_RULES.default;

    await state.waitForSelectors(rule.waitSelectors || []);

    const descriptionSelectors = state.uniqueItems(rule.descriptionSelectors || [], state.SHARED_DESCRIPTION_SELECTORS);
    const basicInfoSelectors = state.uniqueItems(rule.basicInfoSelectors || [], state.SHARED_BASIC_INFO_SELECTORS);

    const additionalImages = await state.collectImagesByRule(rule, options.coverUrl || '');
    const description = state.cleanDescriptionText(
      state.pickTextByPlans(rule.descriptionPlans || [], descriptionSelectors, { minLength: 20, maxLength: 2000 }),
      { keepStructuredSections: genre === '音楽映像' || genre === 'マンガ' }
    );
    const basicInfo = state.cleanBasicInfoText(state.pickBasicInfoText(basicInfoSelectors));

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
    version: state.version,
    scrapePage
  };
})();
