export function uniqueImageCandidatesByUrl(items = []) {
  const seen = new Set();
  const result = [];

  for (const item of items) {
    const url = String(item?.url || "").trim();
    if (!url || seen.has(url)) {
      continue;
    }
    seen.add(url);
    result.push(item);
  }

  return result;
}
