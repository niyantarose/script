function メニュー_韓国マンガシート作成() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  
  const シート名 = '韓国マンガ';
  if (ss.getSheetByName(シート名)) {
    ui.alert('「' + シート名 + '」シートは既に存在します');
    return;
  }

  const sh = ss.insertSheet(シート名);

  // ヘッダー定義
  const ヘッダー = [
    '登録状況',
    '管理コード',
    '言語コード',
    '版種コード',
    '作品番号',
    'カテゴリコード',
    '商品名(原題)',
    '商品名(日本語)',
    '著者',
    '出版社',
    '発売日',
    'ISBN',
    '言語',
    '版種名',
    '付録情報',
    '特典メモ',
    '単巻数',
    'セット巻数開始',
    'セット巻数終了',
    'アラジン商品コード',
    'アラジンURL',
    'yes24商品コード',
    'yes24URL',
    'Kyobo商品コード',
    'KyoboURL',
    'メイン画像URL',
    '追加画像URL',
    '原価',
    '売価',
    '配送パターン',
    '重複チェック用正規化タイトル',
    '登録日',
    '登録者',
    '備考'
  ];

  // ヘッダー行を書き込み
  sh.getRange(1, 1, 1, ヘッダー.length).setValues([ヘッダー]);

  // ヘッダー行の書式
  const ヘッダー範囲 = sh.getRange(1, 1, 1, ヘッダー.length);
  ヘッダー範囲.setBackground('#4a86e8');
  ヘッダー範囲.setFontColor('#ffffff');
  ヘッダー範囲.setFontWeight('bold');
  ヘッダー範囲.setHorizontalAlignment('center');

  // 行の高さ・列の固定
  sh.setRowHeight(1, 40);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2); // 登録状況・管理コードを固定

  // 登録状況列（A列）の条件付き書式：未登録=黄色・登録済み=色なし
  const 最終行 = 1000;
  const 全体範囲 = sh.getRange('A2:AH' + 最終行);
  const 登録状況ルール = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$A2="未登録"')
    .setBackground('#ffff00')
    .setRanges([全体範囲])
    .build();
  sh.setConditionalFormatRules([登録状況ルール]);

  // 登録状況列のプルダウン
  sh.getRange(2, 1, 最終行 - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['未登録', '登録済み'], true)
      .build()
  );

  // 店舗列グループ化（アラジン・yes24・Kyobo）
  // 列番号：アラジン=20-21、yes24=22-23、Kyobo=24-25
  sh.columnGroups || null; // グループ化APIはスプレッドシートUIから手動で設定推奨

  ui.alert('✅ 韓国マンガシートを作成しました\n\n次にメニュー「⑥プルダウン更新」を実行してください');
}