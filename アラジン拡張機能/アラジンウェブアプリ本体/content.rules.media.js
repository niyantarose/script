(() => {
  const state = globalThis.__ALADIN_CONTENT__;
  if (!state) {
    return;
  }

  state.SCRAPER_RULES.音楽映像 = {
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
        excludeSectionLabels: state.DESCRIPTION_EXCLUDE_SECTION_LABELS,
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
        excludeSectionLabels: state.DEFAULT_EXCLUDE_SECTION_LABELS,
        maxCount: 10
      },
      {
        kind: 'introduce',
        selectors: ['.pContent[id$="_Introduce"] img'],
        maxCount: 6
      }
    ],
    excludeAncestorSelectors: state.SHARED_IMAGE_EXCLUDE_ANCESTORS,
    excludeSectionLabels: state.DEFAULT_EXCLUDE_SECTION_LABELS,
    allowGenericFallback: true,
    fallbackMaxCount: 6
  };
})();
