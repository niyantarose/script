function アンダーバーをハイフンに置換() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();

  // 商品コードが入ってる列だけ指定(A=1,B=2…)
  const LOCS = [
    { sheet:'発注',       col:12, headerRow:6 }, // L列
    { sheet:'商品マスタ', col:2,  headerRow:6 }, // B列
    { sheet:'EMSリスト',  col:9,  headerRow:6 }, // I列
  ];

  const res = ui.alert(
    '商品コードの「_」を「-」に置換する?',
    '対象は各シートの商品コード列だけ。購入No.やEMS番号は触らへん。',
    ui.ButtonSet.OK_CANCEL
  );
  if (res !== ui.Button.OK) return;

  let total = 0;
  const log = [];
  LOCS.forEach(loc => {
    const sh = ss.getSheetByName(loc.sheet);
    if (!sh) { log.push(`${loc.sheet}:シート無し`); return; }
    const last = sh.getLastRow();
    if (last <= loc.headerRow) { log.push(`${loc.sheet}:データ無し`); return; }
    const rng = sh.getRange(loc.headerRow + 1, loc.col, last - loc.headerRow, 1);
    const n = rng.createTextFinder('_').matchEntireCell(false).replaceAllWith('-');
    total += n;
    log.push(`${loc.sheet}:${n}件`);
  });

  ui.alert(`置換完了(計${total}件)\n` + log.join('\n'));
}