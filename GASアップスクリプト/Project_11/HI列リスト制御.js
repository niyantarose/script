/** 全角→半角（英数字／ハイフン／アンダーバー） */
function TOHALFWIDTH(text) {
  if (text === null || text === "") return "";
  return text.toString()
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
    .replace(/[ー＿]/g, s => s === "ー" ? "-" : s === "＿" ? "_" : s);
}

/** onEdit：作成月Mを初回セット＋H/IのプルダウンをBあり行に付与 */
function onEdit(e) {
  if (!e || !e.range) return; // 手動実行で落ちない用
  const sh = e.range.getSheet();
  if (sh.getName() !== "在庫マスター") return; // ←タブ名を合わせる
  const row = e.range.getRow(), col = e.range.getColumn();
  if (row <= 1) return; // ヘッダ除外

  // 1) B列（商品ID）編集時の処理
  if (col === 2) {
    const has = String(e.range.getValue()).trim() !== "";

    // ★作成月（M列=13）を初めて値が入ったときだけ当月1日に固定
    const mCell = sh.getRange(row, 13);
    if (has && !mCell.getValue()) {
      const now = new Date();
      const firstOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
      mCell.setValue(firstOfMonth);
    }

    // ★H/I プルダウンを付け外し（空行なら消す）
    const ruleH = SpreadsheetApp.newDataValidation()
      .requireValueInList(["即納在庫あり","お取り寄せ在庫あり","在庫なし"], true)
      .setAllowInvalid(false).build();
    const ruleI = SpreadsheetApp.newDataValidation()
      .requireValueInList(["予約待ち","発注中","EMS入庫待ち","入庫済み"], true)
      .setAllowInvalid(false).build();

    const hCell = sh.getRange(row, 8);  // H
    const iCell = sh.getRange(row, 9);  // I
    if (has) {
      hCell.setDataValidation(ruleH);
      iCell.setDataValidation(ruleI);
    } else {
      hCell.clearDataValidations(); hCell.clearContent();
      iCell.clearDataValidations(); iCell.clearContent();
    }
  }
}

/** 既存データ整備：列一括ルールを消したあと、Bが入ってる行にH/Iを一括再付与（手動実行OK） */
function applyDropdownsForExistingRows() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("在庫マスター");
  const last = sh.getLastRow();
  if (last < 2) return;
  const ids = sh.getRange(2, 2, last - 1, 1).getValues(); // B2:B

  const ruleH = SpreadsheetApp.newDataValidation()
    .requireValueInList(["即納在庫あり","お取り寄せ在庫あり","在庫なし"], true)
    .setAllowInvalid(false).build();
  const ruleI = SpreadsheetApp.newDataValidation()
    .requireValueInList(["予約待ち","発注中","EMS入庫待ち","入庫済み"], true)
    .setAllowInvalid(false).build();

  for (let i = 0; i < ids.length; i++) {
    const row = i + 2;
    const has = String(ids[i][0]).trim() !== "";
    const h = sh.getRange(row, 8), ii = sh.getRange(row, 9);
    if (has) {
      h.setDataValidation(ruleH);
      ii.setDataValidation(ruleI);
    } else {
      h.clearDataValidations(); h.clearContent();
      ii.clearDataValidations(); ii.clearContent();
    }
  }
}
