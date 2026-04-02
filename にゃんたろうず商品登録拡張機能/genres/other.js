import { buildAuthorPublisherSection } from "../sections/authorPublisher.js";
import { buildBasicInfoSection } from "../sections/basicInfo.js";
import { buildDescriptionSection } from "../sections/description.js";

export const otherGenreProfile = {
  id: "other",
  label: "Other",
  requiredSections: ["basicInfo", "description", "authorPublisher"],
  imageBuckets: ["main", "detail"],
  dataFields: ["title", "author", "publisher", "publishDate", "isbn", "price", "pageUrl"],
  match() {
    return true;
  },
  normalize(raw) {
    return {
      genre: "other",
      subGenre: null,
      sections: {
        basicInfo: buildBasicInfoSection(raw),
        description: buildDescriptionSection(raw),
        authorPublisher: buildAuthorPublisherSection(raw)
      },
      warnings: ["Phase 1 では manga 以外は簡易プロファイルで保存します"]
    };
  }
};
