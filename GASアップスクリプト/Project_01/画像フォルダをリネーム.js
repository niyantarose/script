function 全シート_フォルダをSKUにリネーム() {
  const FOLDER_ID = '1Vk1cVgrXM5CNFqfBvAaBUJpheDn4f9Lz';
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();

  // 全シートからサイト商品コード→SKUマップを構築
  const コードマップ = {};
  const 対象シート名 = ['台湾まんが', '台湾書籍その他', '台湾グッズ', '台湾雑誌'];

  for (const シート名 of 対象シート名) {
  const sh = ss.getSheetByName(シート名);
  if (!sh) continue;
  const 列マップ = _kyoutuu.列番号を取得(sh);
  const サイト商品コード列 = 列マップ['サイト商品コード'];
  
  // シートによってSKU列名が違う
  const SKU列 = 列マップ['商品コード(SKU)'] || 列マップ['親コード'];
  if (!サイト商品コード列 || !SKU列) continue;

  const 最終行 = sh.getLastRow();
  if (最終行 < 2) continue;
  const data = sh.getRange(2, 1, 最終行 - 1, sh.getLastColumn()).getValues();
  for (const r of data) {
    const サイトコード = String(r[サイト商品コード列 - 1] || '').trim();
    const SKU = String(r[SKU列 - 1] || '').trim();
    if (サイトコード && SKU) コードマップ[サイトコード] = SKU;
  }
}

  // フォルダを走査してリネーム
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const subfolders = folder.getFolders();
  let リネーム数 = 0, スキップ数 = 0;

  while (subfolders.hasNext()) {
    const sub = subfolders.next();
    const 現在名 = sub.getName();
    if (コードマップ[現在名]) {
      const 新名前 = コードマップ[現在名];
      if (現在名 !== 新名前) {
        sub.setName(新名前);
        リネーム数++;
      }
    } else {
      スキップ数++;
    }
  }

  ui.alert(`✅ リネーム完了\nリネーム: ${リネーム数}件\nマッチなし: ${スキップ数}件`);
}