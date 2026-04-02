import { resolveFolderNameStrategyValue } from "../core/pathPolicy.js";
import { normalizeImages } from "./normalizeImages.js";
import { normalizeProduct } from "./normalizeProduct.js";

export function buildPayload({ adapter, raw, genreProfile, settings }) {
  const productPayload = normalizeProduct({ adapter, raw, genreProfile });
  const folderNameStrategyValue = resolveFolderNameStrategyValue(productPayload, settings);

  return {
    ...productPayload,
    product: {
      ...productPayload.product,
      folderNameStrategyValue
    },
    images: normalizeImages({ adapter, raw, genreProfile }),
    raw: {
      html: raw.rawHtml || "",
      extractedAt: new Date().toISOString()
    }
  };
}
