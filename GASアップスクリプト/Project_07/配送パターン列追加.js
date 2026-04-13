function 配送パターンマスターを作成() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const シート名 = '配送パターンマスター';
  if (ss.getSheetByName(シート名)) { ui.alert('「配送パターンマスター」シートは既に存在します'); return; }

  const sh = ss.insertSheet(シート名);
  sh.getRange(1, 1).setValue('配送パターン');
  sh.getRange(1, 1).setBackground('#4a86e8').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  const 初期値 = [['佐川'], ['ゆうパケ_1(450)'], ['ゆうパケ_2(950)'], ['ゆうパケ_3(無料)']];
  sh.getRange(2, 1, 初期値.length, 1).setValues(初期値);
  sh.setColumnWidth(1, 160);
  sh.setFrozenRows(1);

  ui.alert('✅ 配送パターンマスターを作成しました\n\nA列に配送パターンを追加・削除して管理できます');
}

function 配送パターンドロップダウン一括設定() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const マスター = ss.getSheetByName('配送パターンマスター');
  if (!マスター) {
    SpreadsheetApp.getUi().alert('❌ 配送パターンマスターが見つかりません\n先に「配送パターンマスターを作成」を実行してください');
    return;
  }

  // マスターシートのA列を参照するルール
  const 最終行マスター = Math.max(マスター.getLastRow(), 2);
  const masterRange = マスター.getRange(2, 1, 最終行マスター - 1, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(masterRange, true)
    .setAllowInvalid(false)
    .build();

  // 列が確定済みのシート
  const 固定シート = [
    { name: '台湾グッズ', col: 10 },
  ];

  // 粗利益率の右隣に列挿入するシート
  const 挿入シート名 = [
    '台湾_書籍（コミック/小説/設定集）',
    '韓国グッズ',
    '韓国マンガ',
    '韓国書籍',
    '韓国音楽映像',
  ];

  const results = [];

  // ① 固定シート：ドロップダウンだけ更新
  固定シート.forEach(({ name, col }) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) { results.push('❌ シートなし: ' + name); return; }
    const lastRow = Math.max(sheet.getLastRow(), 100);
    sheet.getRange(2, col, lastRow - 1, 1).setDataValidation(rule);
    results.push('✅ ドロップダウン更新: ' + name);
  });

  // ② 挿入シート：ヘッダーから配送パターン列を探してドロップダウン更新
  挿入シート名.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) { results.push('❌ シートなし: ' + name); return; }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const 配送Col = headers.indexOf('配送パターン');

    if (配送Col === -1) {
      // 配送パターン列がなければ粗利益率の右隣に新規挿入
      const 粗利益率Idx = headers.indexOf('粗利益率');
      if (粗利益率Idx === -1) { results.push('⚠ 粗利益率列も配送パターン列もなし: ' + name); return; }

      const 挿入Col = 粗利益率Idx + 2;
      sheet.insertColumnAfter(粗利益率Idx + 1);
      const headerCell = sheet.getRange(1, 挿入Col);
      headerCell.setValue('配送パターン');
      headerCell.setBackground('#FF0000').setFontColor('#FFFFFF').setFontWeight('bold');
      const lastRow = Math.max(sheet.getLastRow(), 100);
      sheet.getRange(2, 挿入Col, lastRow - 1, 1).setDataValidation(rule);
      results.push('✅ 列追加＋ドロップダウン設定: ' + name + '（' + 挿入Col + '列目）');
    } else {
      // 既存の配送パターン列のドロップダウンを更新
      const lastRow = Math.max(sheet.getLastRow(), 100);
      sheet.getRange(2, 配送Col + 1, lastRow - 1, 1).setDataValidation(rule);
      results.push('✅ ドロップダウン更新（マスター参照）: ' + name);
    }
  });

  SpreadsheetApp.getUi().alert(results.join('\n'));
}