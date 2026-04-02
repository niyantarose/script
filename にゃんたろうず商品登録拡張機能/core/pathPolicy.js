import { ROOT_DOWNLOAD_FOLDER } from "./constants.js";
import { sanitizeFolderSegment } from "./sanitize.js";

export function resolveFolderNameStrategyValue(payload, settings) {
  const titleSegment = sanitizeFolderSegment(payload.product.title || payload.product.siteProductCode || "item");
  const codeSegment = sanitizeFolderSegment(payload.product.siteProductCode || payload.product.title || "item");
  const strategy = settings.folderNameStrategy || "siteProductCode__productTitle";

  if (strategy === "productTitle") {
    return titleSegment;
  }

  if (strategy === "siteProductCode") {
    return codeSegment;
  }

  return sanitizeFolderSegment(`${codeSegment}__${titleSegment}`, codeSegment || titleSegment || "item");
}

export function buildBaseFolder(payload, settings) {
  return `${ROOT_DOWNLOAD_FOLDER}/${resolveFolderNameStrategyValue(payload, settings)}`;
}

export function buildSiteFolder(payload) {
  return payload.source.siteKey;
}

export function buildGenreFolder(payload) {
  return payload.product.genre;
}

export function buildImageFolder(payload, settings) {
  return buildBaseFolder(payload, settings);
}

export function buildDataFolder(payload, settings) {
  return buildBaseFolder(payload, settings);
}

export function buildImagePath(payload, settings, bucket, filename) {
  return `${buildImageFolder(payload, settings)}/${filename}`;
}

export function buildDataPath(payload, settings, filename) {
  return `${buildDataFolder(payload, settings)}/${filename}`;
}

export function buildFolderPreview(payload, settings) {
  return buildBaseFolder(payload, settings);
}
