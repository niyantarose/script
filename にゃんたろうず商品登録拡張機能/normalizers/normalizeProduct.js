export function normalizeProduct({ adapter, raw, genreProfile }) {
  const normalized = genreProfile.normalize(raw);

  return {
    source: {
      country: adapter.country,
      siteKey: adapter.siteKey,
      pageUrl: raw.pageUrl
    },
    product: {
      title: raw.rawFields?.title || raw.titleCandidate || "",
      siteProductCode: raw.siteProductCode || "",
      folderNameStrategyValue: "",
      genre: normalized.genre,
      subGenre: normalized.subGenre || null
    },
    sections: normalized.sections,
    meta: {
      warnings: normalized.warnings || []
    }
  };
}
