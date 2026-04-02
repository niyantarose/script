export function buildPriceInfoSection(raw) {
  return {
    price: String(raw.rawFields?.price || "").trim()
  };
}
