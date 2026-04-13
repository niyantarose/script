function setDropdowns() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("シート1");
  var range = sheet.getRange("B2:B1000"); // 商品ID列をチェック
  var values = range.getValues();
  
  for (var i = 0; i < values.length; i++) {
    var row = i + 2;
    if (values[i][0] !== "") {
      // 商品IDがある行 → ドロップダウン設定
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(["即納在庫あり", "お取り寄せ在庫あり", "在庫なし"], true)
        .build();
      sheet.getRange("H" + row).setDataValidation(rule);
    } else {
      // 商品IDが空欄の行 → ドロップダウン解除
      sheet.getRange("H" + row).clearDataValidations();
    }
  }
}
