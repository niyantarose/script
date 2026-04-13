function fetchStoreCategoryKeys() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const tokenSheet  = ss.getSheetByName("トークン");
  const accessToken = tokenSheet.getRange("A2").getValue();
  if (!accessToken) throw new Error("アクセストークンがありません。");

  const sellerId = "niyantarose";  // ← ご自分のストアIDに変更
  // １階層目のカテゴリを取得（page_key省略）
  const url = [
    "https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/stCategoryList",
    "?seller_id=", encodeURIComponent(sellerId)
  ].join("");

  const resp = UrlFetchApp.fetch(url, {
    headers: { "Authorization": "Bearer " + accessToken }
  }).getContentText();

  // XMLをパースしてカテゴリ名とページキーをログ出力
  const root    = XmlService.parse(resp).getRootElement();
  const results = root.getChildren("Result");
  results.forEach(r => {
    const name    = r.getChildText("Name");
    const pageKey = r.getChildText("PageKey");
    Logger.log("カテゴリ名：%s  — ページキー：%s", name, pageKey);
  });
}
