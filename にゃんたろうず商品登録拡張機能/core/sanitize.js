const RESERVED_NAMES = new Set([
  "CON", "PRN", "AUX", "NUL",
  "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
  "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
]);

export function normalizeWhitespace(value) {
  return String(value || "")
    .replace(/\s+/g, " ")
    .trim();
}

export function sanitizeFolderSegment(value, fallback = "item") {
  let cleaned = normalizeWhitespace(value)
    .replace(/[\\/:*?"<>|]/g, "_")
    .replace(/[. ]+$/g, "")
    .replace(/\s+/g, "_")
    .slice(0, 96);

  if (!cleaned) {
    cleaned = fallback;
  }

  if (RESERVED_NAMES.has(cleaned.toUpperCase())) {
    cleaned = `${cleaned}_item`;
  }

  return cleaned || fallback;
}

export function sanitizeFileExtension(value, fallback = "jpg") {
  const normalized = String(value || "")
    .replace(/^[.]+/, "")
    .replace(/[^a-z0-9]/gi, "")
    .toLowerCase();
  return normalized || fallback;
}
