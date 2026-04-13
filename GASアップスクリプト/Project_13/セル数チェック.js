function セル数チェック_リアルタイム() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const LIMIT = 10000000;

  console.log(`処理開始: 全 ${sheets.length} シートをチェックします...`);
  
  let totalCells = 0;

  // forEachではなくforループで1つずつ確実に処理
  for (let i = 0; i < sheets.length; i++) {
    const sh = sheets[i];
    const name = sh.getName();
    
    // ★チェック開始の合図
    console.log(`[${i + 1}/${sheets.length}] シート確認中: "${name}" ...`);
    
    try {
      const r = sh.getMaxRows();
      const c = sh.getMaxColumns();
      const cells = r * c;
      totalCells += cells;
      
      // ★結果をすぐに出す
      console.log(`   ✅ ${cells.toLocaleString()} セル (行:${r} x 列:${c})`);
      
    } catch (e) {
      console.error(`   ❌ エラー: "${name}" の取得に失敗 - ${e.message}`);
    }

    // ★ここで強制的にログを表示させる（これが重要）
    SpreadsheetApp.flush(); 
  }

  console.log("------------------------------------------------");
  console.log(`全てのチェック完了。合計: ${totalCells.toLocaleString()} / 10,000,000`);
  
  if (totalCells > LIMIT * 0.9) {
    ss.toast("警告: セル数が限界に近いです", "完了", -1);
  } else {
    ss.toast("チェック完了", "完了", 5);
  }
}