export function buildAuthorPublisherSection(raw) {
  return {
    author: String(raw.rawFields?.author || "").trim(),
    publisher: String(raw.rawFields?.publisher || "").trim(),
    publishDate: String(raw.rawFields?.publishDate || "").trim(),
    isbn: String(raw.rawFields?.isbn || "").trim()
  };
}
