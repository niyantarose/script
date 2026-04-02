export function buildSpecInfoSection(raw) {
  return {
    text: String(raw.rawSections?.basicInfo || "").trim()
  };
}
