export const ProductPayloadSchemaExample = {
  source: {
    country: "aladin",
    siteKey: "aladin_book",
    pageUrl: "https://www.aladin.co.kr/shop/wproduct.aspx?ItemId=0"
  },
  product: {
    title: "Sample Title",
    siteProductCode: "123456789",
    folderNameStrategyValue: "123456789__Sample_Title",
    genre: "manga",
    subGenre: null
  },
  sections: {
    basicInfo: {},
    description: {},
    authorPublisher: {},
    bonusInfo: {},
    volumeInfo: {}
  },
  images: {
    main: [],
    detail: [],
    bonus: [],
    sample: []
  },
  raw: {
    html: "",
    extractedAt: "2026-03-23T12:00:00+09:00"
  }
};
