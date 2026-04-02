export function buildIssueInfoSection(raw) {
  return {
    publishDate: String(raw.rawFields?.publishDate || "").trim()
  };
}
