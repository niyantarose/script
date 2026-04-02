(() => {
  const state = globalThis.__ALADIN_CONTENT__;
  if (!state) {
    return;
  }

  state.SCRAPER_RULES.マンガ = {
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
        excludeSectionLabels: state.DESCRIPTION_EXCLUDE_SECTION_LABELS,
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
        maxCount: 8
      },
      {
        kind: 'section',
        selectors: ['#bookDescriptionToggle img', '#productDescription img', '.Ere_prod_mconts_box img'],
        includeSectionLabels: [/소개/i, /책소개/i, /출판사/i],
        excludeSectionLabels: state.DEFAULT_EXCLUDE_SECTION_LABELS,
        excludeAncestorSelectors: ['.pContent[id$="_Introduce"]'],
        maxCount: 4
      }
    ],
    excludeAncestorSelectors: state.SHARED_IMAGE_EXCLUDE_ANCESTORS,
    excludeSectionLabels: state.DEFAULT_EXCLUDE_SECTION_LABELS,
    allowGenericFallback: true,
    fallbackMaxCount: 2
  };
})();
