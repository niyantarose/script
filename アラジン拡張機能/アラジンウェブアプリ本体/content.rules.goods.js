(() => {
  const state = globalThis.__ALADIN_CONTENT__;
  if (!state) {
    return;
  }

  state.SCRAPER_RULES.グッズ = {
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
        excludeSectionLabels: state.DESCRIPTION_EXCLUDE_SECTION_LABELS,
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
        excludeSectionLabels: state.DEFAULT_EXCLUDE_SECTION_LABELS,
        maxCount: 12
      },
      {
        kind: 'introduce',
        selectors: ['.pContent[id$="_Introduce"] img'],
        maxCount: 4
      }
    ],
    excludeAncestorSelectors: state.SHARED_IMAGE_EXCLUDE_ANCESTORS,
    excludeSectionLabels: state.DEFAULT_EXCLUDE_SECTION_LABELS,
    allowGenericFallback: true,
    fallbackMaxCount: 8
  };
})();
