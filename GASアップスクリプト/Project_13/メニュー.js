/******************************************************
 * メニュー（Amazon / Qoo10 共通）
 ******************************************************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // ===== Amazon メニュー =====
  ui.createMenu('在庫管理(Amazon)')
    .addItem('Amazon在庫照合', 'Amazon在庫照合')
    .addItem('差異/未登録レポートのみ更新', 'Amazonレポートシート更新_メニュー用')
    .addSeparator()
    .addItem('Amazonチェック全選択', 'Amazon_チェック全選択')
    .addItem('Amazonチェック全解除', 'Amazon_チェック全解除')
    .addSeparator()
    .addItem('差異だけチェックON', 'Amazon差異_全選択')
    .addItem('差異チェック反転', 'Amazon差異_逆選択')
    .addSeparator()
    .addItem('SKU送信（チェック行を転送）', 'Amazon_SKU送信')
    .addToUi();

  // ===== Qoo10 メニュー =====
  ui.createMenu('在庫管理(Qoo10)')
    .addItem('Qoo10在庫照合', 'Qoo10在庫照合')
    .addSeparator()
    .addItem('差異一覧/未登録一覧を更新', 'Qoo10レポートだけ更新')
    .addSeparator()
    .addItem('☑ チェック全選択', 'Qoo10_チェック全選択')
    .addItem('☐ チェック全解除', 'Qoo10_チェック全解除')
    .addSeparator()
    .addItem('差異だけチェックON', 'Qoo10差異_全選択')
    .addItem('差異チェック反転', 'Qoo10差異_逆選択')
    .addSeparator()
    .addItem('▶ Qoo10SKUを送信', 'チェック済みをQoo10価格在庫変更へ送る')
    .addToUi();
}
