/******************************************************
 * 01_menu.gs - メニュー
 ******************************************************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('📌 商品入力シート用')
    .addItem('配送テキストを反映', '配送テキストを反映_商品入力')
    .addItem('商品説明を整形', '商品説明を整形_商品入力')
    .addItem('商品画像URLを挿入（S=メイン / T=詳細）', '商品画像URLを挿入_ST_商品入力')
    .addItem('画像URLをサーバーにコピー（S+T）', 'ST列画像URLをサーバー変換_商品入力')
    .addSeparator()
    .addItem('【一括】商品入力_一括実行', '商品入力_一括実行')
    .addToUi();

  ui.createMenu('📌 Yahoo商品登録用')
    .addItem('配送テキストを反映', '配送テキストを反映_Yahoo商品登録')
    .addItem('商品説明を整形', '商品説明を整形_Yahoo商品登録')
    .addItem('商品画像URLを挿入（S=メイン / T=詳細）', '商品画像URLを挿入_ST_Yahoo商品登録')
    .addItem('画像URLをサーバーにコピー（S+T）', 'ST列画像URLをサーバー変換_Yahoo商品登録')
    .addToUi();
}
