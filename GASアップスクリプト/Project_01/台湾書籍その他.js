const 設定_台湾書籍その他 = {
  マスターシート名: '台湾書籍その他',
  作品シート名:     'Works（書籍専用）',
  言語マスター名:   '言語マスター',
  カテゴリマスター名: 'カテゴリマスター',
  形態マスター名:   '形態マスターシート',
  マスターファイルID: '1ZAy5wtQCq1ixl47MMEIGnrVWeae-F_NfN34lSOUsN5M',

  作品ヘッダー: [
    'WorksKey', '作品ID', '日本語タイトル', '作者', '原題タイトル',
    '登録済み巻', '最新巻', '更新日時', '最新巻(予約込み)', '予約更新日時'
  ],
  作品列数: 10,

  色パレット: ['#b7e1cd', '#fce8b2', '#f4c7c3', '#c9daf8', '#d9d2e9', '#fce5cd', '#fff2cc', '#d9ead3', '#cfe2f3', '#f4cccc'],

  監視列: [
    '作者', '日本語タイトル', '原題タイトル', '原題商品タイトル',
    '言語', 'カテゴリ', '形態(通常/初回限定/特装)',
    '単巻数', 'セット巻数開始番号', 'セット巻数終了番号', '特典メモ',
    'ISBN', 'サイト商品コード', '売価', '原価'
  ],

  列名: {
    発行チェック:     '発番発行',
    商品コード:       '商品コード（SKU）',
    タイトル:         'タイトル',
    作者:             '作者',
    日本語タイトル:   '日本語タイトル',
    原題:             '原題タイトル',
    原題商品タイトル: '原題商品タイトル',
    リンク:           'リンク',
    形態:             '形態(通常/初回限定/特装)',
    言語:             '言語',
    カテゴリ:         'カテゴリ',
    単巻数:           '単巻数',
    セット開始:       'セット巻数開始番号',
    セット終了:       'セット巻数終了番号',
    特典メモ:         '特典メモ',
    ISBN:             'ISBN',
    作品ID:           '作品ID(W)(自動)',
    SKU自動:          'SKU(自動)',
    コードステータス: '商品コードステータス',
    登録状況:         '登録状況',
    配送パターン:     '配送パターン',
    博客來商品コード: 'サイト商品コード'
  }
};

/* ============================================================
 * シート作成
 * ============================================================ */

function 台湾書籍その他_ヘッダー一覧_() {
  return [
    '発番発行',
    '登録状況',
    '商品コード（SKU）',
    'タイトル',
    '作者',
    '日本語タイトル',
    '原題タイトル',
    '原題商品タイトル',
    'リンク',
    '形態(通常/初回限定/特装)',
    '言語',
    'カテゴリ',
    '単巻数',
    'セット巻数開始番号',
    'セット巻数終了番号',
    '特典メモ',
    'ISBN',
    '売価',
    '原価',
    '粗利益率',
    '配送パターン',
    '商品説明',
    'サイト商品コード',
    'メイン画像',
    '追加画像',
    '発売日',
    '作品ID(W)(自動)',
    'SKU(自動)',
    '商品コードステータス',
    '重複チェックキー',
    '登録日',
    '登録者',
    '備考'
  ];
}

function 台湾書籍その他シートを作成() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const シート名 = '台湾書籍その他';

  if (ss.getSheetByName(シート名)) {
    ui.alert(`「${シート名}」シートは既に存在します`);
    return;
  }

  const sh = ss.insertSheet(シート名);
  const ヘッダー = 台湾書籍その他_ヘッダー一覧_();

  sh.getRange(1, 1, 1, ヘッダー.length).setValues([ヘッダー]);

  const 自動列 = [
    '商品コード（SKU）', 'タイトル', '作品ID(W)(自動)', 'SKU(自動)',
    '商品コードステータス', '重複チェックキー', '登録日'
  ];
  const API列 = [
    '原題タイトル', '原題商品タイトル', 'リンク', 'サイト商品コード',
    'メイン画像', '追加画像', '発売日', '原価', '商品説明'
  ];

  for (let i = 0; i < ヘッダー.length; i++) {
    let 色 = '#4a86e8';
    if (自動列.includes(ヘッダー[i])) 色 = '#999999';
    if (API列.includes(ヘッダー[i]))  色 = '#e69138';

    sh.getRange(1, i + 1)
      .setBackground(色)
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  }

  sh.setRowHeight(1, 40);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(3);

  const 最終行 = 1000;

  sh.getRange(2, 1, 最終行 - 1, 1).insertCheckboxes();
  sh.getRange(2, 2, 最終行 - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['未登録', '登録済み'], true)
      .build()
  );

  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$B2="未登録"')
      .setBackground('#ffff00')
      .setRanges([sh.getRange(2, 1, 最終行 - 1, ヘッダー.length)])
      .build()
  ]);

  ui.alert(`✅ ${シート名}シートを作成しました\n\nメニュー「プルダウン更新」を実行してください`);
}

/* ============================================================
 * onEdit / 取込後補完
 * ============================================================ */

function 台湾書籍その他_onEdit(e) {
  台湾書籍系_onEdit_共通_(e, 設定_台湾書籍その他);
}

function 台湾書籍その他_取込後補完_(startRow, numRows) {
  const sh = SpreadsheetApp.getActive().getSheetByName('台湾書籍その他');
  if (!sh) return;
  if (!startRow || !numRows || numRows <= 0) return;

  台湾書籍系_追加行補完_共通_(sh, startRow, numRows, 設定_台湾書籍その他);
}

/* ============================================================
 * まんが → 書籍その他 コピー
 * ============================================================ */

function 台湾書籍その他_コピー元列候補_(先列名) {
  const map = {
    '商品コード（SKU）': ['商品コード（SKU）', '親コード'],
    'リンク': ['リンク', '博客來URL'],
    'メイン画像': ['メイン画像', 'メイン画像URL'],
    '追加画像': ['追加画像', '追加画像URL'],
    'サイト商品コード': ['サイト商品コード', '博客來商品コード'],
    '商品説明': ['商品説明', '備考'],
    '原題商品タイトル': ['原題商品タイトル', '原題タイトル']
  };
  return map[先列名] || [先列名];
}

function 台湾まんがから書籍その他へコピー_軽量版() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const 元 = ss.getSheetByName('台湾まんが');
  const 先 = ss.getSheetByName('台湾書籍その他');

  if (!元 || !先) {
    ss.toast('シートが見つかりません');
    return;
  }

  const 元最終行 = 元.getLastRow();
  const 元最終列 = 元.getLastColumn();
  const 先最終列 = 先.getLastColumn();

  if (元最終行 < 2) {
    ss.toast('台湾まんがにデータがありません');
    return;
  }

  const 元ヘッダー = 元.getRange(1, 1, 1, 元最終列).getValues()[0].map(v => String(v).trim());
  const 先ヘッダー = 先.getRange(1, 1, 1, 先最終列).getValues()[0].map(v => String(v).trim());
  const 元データ = 元.getRange(2, 1, 元最終行 - 1, 元最終列).getValues();

  const カテゴリ列 = 元ヘッダー.indexOf('カテゴリ');
  const タイトル列 = 元ヘッダー.indexOf('タイトル');

  const 出力 = [];

  for (let r = 0; r < 元データ.length; r++) {
    const row = 元データ[r];
    const カテゴリ = カテゴリ列 >= 0 ? String(row[カテゴリ列] || '').trim() : '';
    const タイトル = タイトル列 >= 0 ? String(row[タイトル列] || '').trim() : '';

    const isManga =
      /まんが|漫画/i.test(カテゴリ) ||
      /^台湾版\s*まんが/.test(タイトル);

    if (isManga) continue;

    const 新行 = new Array(先ヘッダー.length).fill('');

    for (let c = 0; c < 先ヘッダー.length; c++) {
      const 先列名 = 先ヘッダー[c];
      const 候補列名一覧 = 台湾書籍その他_コピー元列候補_(先列名);

      let 値 = '';
      for (const 候補列名 of 候補列名一覧) {
        const 元列番号 = 元ヘッダー.indexOf(候補列名);
        if (元列番号 >= 0) {
          値 = row[元列番号];
          break;
        }
      }
      新行[c] = 値;
    }

    出力.push(新行);
  }

  if (出力.length === 0) {
    ss.toast('移動対象データがありません');
    return;
  }

  const 開始行 = Math.max(先.getLastRow() + 1, 2);
  const 必要最終行 = 開始行 + 出力.length - 1;
  const 現在行数 = 先.getMaxRows();

  if (必要最終行 > 現在行数) {
    先.insertRowsAfter(現在行数, 必要最終行 - 現在行数);
  }

  先.getRange(開始行, 1, 出力.length, 先ヘッダー.length).setValues(出力);

  // コピー後に段階生成
  台湾書籍その他_取込後補完_(開始行, 出力.length);

  ss.toast(`✅ ${出力.length}件を台湾書籍その他へコピーしました`, '完了', 8);
}