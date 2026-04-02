import { sanitizeFileExtension } from "./sanitize.js";

export function detectImageExtension(url, fallback = "jpg") {
  try {
    const pathname = new URL(String(url || "")).pathname || "";
    const rawExtension = pathname.split(".").pop() || "";
    if (/^jpe?g$/i.test(rawExtension)) {
      return "jpg";
    }
    return sanitizeFileExtension(rawExtension, fallback);
  } catch (_error) {
    return fallback;
  }
}

export function buildImageFilename(sequence, extension) {
  return `${Number(sequence)}.${sanitizeFileExtension(extension).toUpperCase()}`;
}

export function buildDataFilename(kind) {
  if (kind === "raw") {
    return "raw.html";
  }
  return `${kind}.json`;
}

export function resolveConflictAction(reRunPolicy) {
  return reRunPolicy === "continue" ? "uniquify" : "overwrite";
}
