export function buildBasicInfoSection(raw) {
  return {
    text: String(raw.rawSections?.basicInfo || "").trim()
  };
}
