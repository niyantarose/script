/**
 * resetRawSheet()
 * Qoo10在庫_raw シートをまるごと消して空シートを再作成します。
 */
function resetRawSheet() {
  var ss    = SpreadsheetApp.getActive();
  var name  = 'Qoo10在庫_raw';
  var existing = ss.getSheetByName(name);
  if (existing) ss.deleteSheet(existing);
  ss.insertSheet(name);
}
