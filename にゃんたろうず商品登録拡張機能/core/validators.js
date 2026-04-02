export function ensureSupportedPageContext(context) {
  if (!context?.supported) {
    throw new Error(context?.reason || "対応サイトではありません");
  }
  return context;
}

export function ensureProbePayload(raw) {
  if (!raw || typeof raw !== "object") {
    throw new Error("ページから商品情報を取得できませんでした");
  }
  if (!raw.pageUrl) {
    throw new Error("ページ URL を取得できませんでした");
  }
  return raw;
}
