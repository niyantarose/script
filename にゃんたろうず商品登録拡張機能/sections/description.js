export function buildDescriptionSection(raw) {
  return {
    text: String(raw.rawSections?.description || "").trim()
  };
}
