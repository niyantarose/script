import { uniqueImageCandidatesByUrl } from "../../../core/dedupe.js";

export function normalizeAladinBookRaw(raw = {}) {
  return {
    pageUrl: String(raw.pageUrl || "").trim(),
    titleCandidate: String(raw.titleCandidate || raw.rawFields?.title || "").trim(),
    siteProductCode: String(raw.siteProductCode || "").trim(),
    rawFields: {
      title: String(raw.rawFields?.title || raw.titleCandidate || "").trim(),
      author: String(raw.rawFields?.author || "").trim(),
      publisher: String(raw.rawFields?.publisher || "").trim(),
      publishDate: String(raw.rawFields?.publishDate || "").trim(),
      isbn: String(raw.rawFields?.isbn || "").trim(),
      price: String(raw.rawFields?.price || "").trim(),
      coverUrl: String(raw.rawFields?.coverUrl || "").trim(),
      categoryText: String(raw.rawFields?.categoryText || "").trim()
    },
    rawSections: {
      basicInfo: String(raw.rawSections?.basicInfo || "").trim(),
      description: String(raw.rawSections?.description || "").trim()
    },
    imageCandidates: uniqueImageCandidatesByUrl(Array.isArray(raw.imageCandidates) ? raw.imageCandidates : []),
    rawHtml: String(raw.rawHtml || "")
  };
}
