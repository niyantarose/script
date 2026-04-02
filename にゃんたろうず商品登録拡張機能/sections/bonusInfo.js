export function buildBonusInfoSection(raw) {
  const description = String(raw.rawSections?.description || "");
  const matched = description.match(/(?:特典|증정|부록).{0,120}/i);
  return {
    text: matched ? matched[0].trim() : ""
  };
}
