function 韓国書籍シートを作成() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const シート名 = '韓国書籍';
  if (ss.getSheetByName(シート名)) { ui.alert(`「${シート名}」シートは既に存在します`); return; }

  const sh = ss.insertSheet(シート名);
  const ヘッダー = [
    '登録状況', '管理コード', 'タイトル(自動)', '作品番号', '言語コード', '版種コード', 'カテゴリコード',
    '商品名(日本語)', '商品名(原題)', '付録情報', '言語', 'カテゴリ', '版種名',
    '単巻数', 'セット巻数開始', 'セット巻数終了', '特典メモ',
    '著者', '出版社', '発売日', 'ISBN',
    'アラジン商品コード', 'アラジンURL',
    'yes24商品コード', 'yes24URL',
    'Kyobo商品コード', 'KyoboURL',
    'メイン画像URL', '追加画像URL',
    '原価', '売価', '配送パターン',
    '重複チェック用正規化タイトル', '登録日',
    '登録者', '備考'
  ];

  sh.getRange(1, 1, 1, ヘッダー.length).setValues([ヘッダー]);

  const 自動列 = ['管理コード', 'タイトル(自動)', '作品番号', '言語コード', '版種コード', 'カテゴリコード', '重複チェック用正規化タイトル', '登録日'];
  const API列 = ['商品名(原題)', '著者', '出版社', '発売日', 'ISBN', '原価',
                 'アラジン商品コード', 'アラジンURL', 'yes24商品コード', 'yes24URL',
                 'Kyobo商品コード', 'KyoboURL', 'メイン画像URL', '追加画像URL'];

  for (let i = 0; i < ヘッダー.length; i++) {
    const cell = sh.getRange(1, i + 1);
    let 色 = '#4a86e8'; // 青（手動）
    if (自動列.includes(ヘッダー[i])) 色 = '#999999'; // グレー（自動）
    if (API列.includes(ヘッダー[i])) 色 = '#e69138';  // オレンジ（API）
    cell.setBackground(色);
    cell.setFontColor('#ffffff');
    cell.setFontWeight('bold');
    cell.setHorizontalAlignment('center');
  }

  sh.setRowHeight(1, 40);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);

  const 最終行 = 1000;
  sh.getRange(2, 1, 最終行 - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['未登録', '登録済み'], true).build()
  );

  const 全体範囲 = sh.getRange(2, 1, 最終行 - 1, ヘッダー.length);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2="未登録"')
      .setBackground('#ffff00')
      .setRanges([全体範囲])
      .build()
  ]);

  ui.alert('✅ 韓国書籍シートを作成しました\n\n韓国書籍シートをアクティブにした状態で\nメニュー「⑥プルダウン更新」を実行してください');
}