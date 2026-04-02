import { buildAuthorPublisherSection } from "../sections/authorPublisher.js";
import { buildBasicInfoSection } from "../sections/basicInfo.js";
import { buildBonusInfoSection } from "../sections/bonusInfo.js";
import { buildDescriptionSection } from "../sections/description.js";

function extractVolumeInfo(title) {
  const matched = String(title || "").match(/(\d+\s*(?:권|巻))/i);
  return {
    text: matched ? matched[1] : ""
  };
}

export const mangaGenreProfile = {
  id: "manga",
  label: "Manga",
  requiredSections: ["basicInfo", "description", "authorPublisher", "bonusInfo", "volumeInfo"],
  imageBuckets: ["main", "detail", "bonus", "sample"],
  dataFields: ["title", "author", "publisher", "publishDate", "isbn", "price", "pageUrl"],
  match(raw) {
    const combined = `${raw.rawFields?.categoryText || ""} ${raw.rawFields?.title || ""}`;
    return /만화|코믹|comic|webtoon|웹툰/i.test(combined);
  },
  normalize(raw) {
    return {
      genre: "manga",
      subGenre: null,
      sections: {
        basicInfo: buildBasicInfoSection(raw),
        description: buildDescriptionSection(raw),
        authorPublisher: buildAuthorPublisherSection(raw),
        bonusInfo: buildBonusInfoSection(raw),
        volumeInfo: extractVolumeInfo(raw.rawFields?.title)
      },
      warnings: []
    };
  }
};
