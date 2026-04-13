function listStoreCategories() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const token    = ss.getSheetByName("トークン").getRange("A2").getValue();
  const sellerId = "niyantarose";
  const url      = 
    "https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/stCategoryList" +
    "?seller_id=" + encodeURIComponent(sellerId);

  const res = UrlFetchApp.fetch(url, {
    headers: { "Authorization": "Bearer " + token }
  });
  Logger.log(res.getContentText());
}
