function 全シート_フォルダをSKUにリネーム() {
  const FOLDER_ID = '1Vk1cVgrXM5CNFqfBvAaBUJpheDn4f9Lz';
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();

  // 全シートから「サイト商品コード / 博客來商品コード → SKU」マップを構築
  const コードマップ = {};
  const 対象シート名 = ['台湾まんが', '台湾書籍その他', '台湾グッズ', '台湾雑誌'];

  let 競合数 = 0;

  for (const シート名 of 対象シート名) {
    const sh = ss.getSheetByName(シート名);
    if (!sh) continue;

    const 列マップ = _kyoutuu.列番号を取得(sh);

    const サイト商品コード列 =
      列マップ['サイト商品コード'] || 0;

    const 博客來商品コード列 =
      列マップ['博客來商品コード'] || 0;

    // シートによってSKU列名が違うので候補を広めに見る
    const SKU列 =
      列マップ['商品コード（SKU）'] ||
      列マップ['商品コード(SKU)'] ||
      列マップ['親コード'] ||
      列マップ['商品コード'] ||
      0;

    if (!(SKU列 && (サイト商品コード列 || 博客來商品コード列))) continue;

    const 最終行 = sh.getLastRow();
    if (最終行 < 2) continue;

    const data = sh.getRange(2, 1, 最終行 - 1, sh.getLastColumn()).getValues();

    for (const r of data) {
      const SKU = String(r[SKU列 - 1] || '').trim();
      if (!SKU) continue;

      const 候補コード一覧 = [
        サイト商品コード列 ? String(r[サイト商品コード列 - 1] || '').trim() : '',
        博客來商品コード列 ? String(r[博客來商品コード列 - 1] || '').trim() : ''
      ].filter(Boolean);

      for (const コード of 候補コード一覧) {
        if (!コード) continue;

        if (コードマップ[コード] && コードマップ[コード] !== SKU) {
          // 同じコードに別SKUがぶつかった時は数だけ記録
          競合数++;
          continue;
        }

        コードマップ[コード] = SKU;
      }
    }
  }

  // フォルダを走査してリネーム
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const subfolders = folder.getFolders();

  let リネーム数 = 0;
  let マッチ済みそのまま = 0;
  let スキップ数 = 0;

  while (subfolders.hasNext()) {
    const sub = subfolders.next();
    const 現在名 = String(sub.getName() || '').trim();

    const 新名前 = コードマップ[現在名];
    if (!新名前) {
      スキップ数++;
      continue;
    }

    if (現在名 === 新名前) {
      マッチ済みそのまま++;
      continue;
    }

    sub.setName(新名前);
    リネーム数++;
  }

  ui.alert(
    `✅ リネーム完了\n` +
    `リネーム: ${リネーム数}件\n` +
    `一致したが変更なし: ${マッチ済みそのまま}件\n` +
    `マッチなし: ${スキップ数}件\n` +
    `コード競合: ${競合数}件`
  );
}