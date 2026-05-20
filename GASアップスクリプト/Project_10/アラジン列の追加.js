/**
 * アラジン列追加.gs
 * 各シートにアラジンAPIで取得できる項目の列を末尾に追加する
 * メニュー → 🔴 アラジン列追加 から実行
 */

// シートごとに追加する列の定義
const ALADIN_ADD_COLUMNS = {
  '韓国書籍': [
    'アラジンURL',
    'アラジン商品ID',
    '発売日',
    '原価',
    '出版社',
    'メイン画像URL',
    '追加画像URL',
    '商品説明',
  ],
  '韓国マンガ': [
    'アラジンURL',
    'アラジン商品ID',
    '発売日',
    '原価',
    '出版社',
    'メイン画像URL',
    '追加画像URL',
    '商品説明',
  ],
  '韓国音楽映像': [
    'アラジン商品ID',
    '商品説明',
  ],
  '韓国グッズ': [
    'アラジン商品ID',
    '発売日',
    '商品説明',
  ],
};

function アラジン列追加() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();

  const 対象シート = Object.keys(ALADIN_ADD_COLUMNS);
  const 追加結果 = [];

  for (const shName of 対象シート) {
    const sh = ss.getSheetByName(shName);
    if (!sh) {
      追加結果.push(`⚠ ${shName}：シートが見つかりません`);
      continue;
    }

    const lastCol = sh.getLastColumn();
    const ヘッダー = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());

    const 追加列 = ALADIN_ADD_COLUMNS[shName];
    let 追加数 = 0;

    for (const colName of 追加列) {
      // 既存列があればスキップ
      if (ヘッダー.includes(colName)) continue;

      // 末尾に追加
      const newCol = sh.getLastColumn() + 1;
      sh.getRange(1, newCol).setValue(colName);

      // ヘッダーセルのスタイルを隣のセルに合わせる（薄いグレー背景）
      const headerCell = sh.getRange(1, newCol);
      headerCell.setBackground('#f3f3f3');
      headerCell.setFontWeight('bold');
      headerCell.setBorder(true, true, true, true, false, false,
        '#cccccc', SpreadsheetApp.BorderStyle.SOLID);

      ヘッダー.push(colName); // 重複チェック用に更新
      追加数++;
    }

    追加結果.push(`✅ ${shName}：${追加数}列追加`);
  }

  // ALADIN_COLUMN_MAPを新しい列名に対応させるため更新内容を表示
  ui.alert(
    'アラジン列追加完了',
    追加結果.join('\n') + '\n\nアラジン.gsのCOLUMN_MAPを更新してください。\n（メッセージを閉じた後にスクリプトエディタを確認）',
    ui.ButtonSet.OK
  );

  // 新しいCOLUMN_MAPをログに出力
  Logger.log('=== 更新後の ALADIN_COLUMN_MAP ===\n' + 新マップを生成_());
}

// 追加後の列名に対応したCOLUMN_MAPをログ出力用に生成
function 新マップを生成_() {
  return `const ALADIN_COLUMN_MAP = {
  '韓国書籍': {
    trigger:      ['ISBN', 'アラジンURL'],
    url:          'アラジンURL',
    isbn:         'ISBN',
    title:        '商品名(原題)',
    author:       '著者',
    publisher:    '出版社',
    pubDate:      '発売日',
    price:        '原価',
    cover:        'メイン画像URL',
    additionalImages: '追加画像URL',
    description:  '商品説明',
    categoryName: 'カテゴリ',
    itemId:       'アラジン商品ID',
  },
  '韓国マンガ': {
    trigger:      ['ISBN', 'アラジンURL'],
    url:          'アラジンURL',
    isbn:         'ISBN',
    title:        '商品名(原題)',
    author:       '著者',
    publisher:    '出版社',
    pubDate:      '発売日',
    price:        '原価',
    cover:        'メイン画像URL',
    additionalImages: '追加画像URL',
    description:  '商品説明',
    categoryName: 'カテゴリ',
    itemId:       'アラジン商品ID',
  },
  '韓国音楽映像': {
    trigger:      ['アラジンURL'],
    url:          'アラジンURL',
    isbn:         'JANコード',
    title:        '商品名(原題)',
    author:       'アーティスト名',
    pubDate:      '発売日',
    price:        '原価',
    cover:        'メイン画像URL',
    additionalImages: '追加画像URL',
    description:  '商品説明',
    categoryName: 'カテゴリ',
    itemId:       'アラジン商品ID',
  },
  '韓国グッズ': {
    trigger:      ['購入URL'],
    url:          '購入URL',
    title:        '商品名（原題）',
    pubDate:      '発売日',
    price:        '原価',
    cover:        'メイン画像URL',
    additionalImages: '追加画像URL',
    description:  '商品説明',
    itemId:       'アラジン商品ID',
  },
};`;
}