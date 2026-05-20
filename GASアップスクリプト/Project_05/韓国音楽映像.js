/**
 * 韓国音楽映像.gs
 * 韓国音楽映像シート専用の設定とメニューラッパー
 *
 * 方針：
 * - Works は「作品・アルバム・映像タイトル単位」
 * - WorksKeyは内部的に「商品名(タイトル) + アーティスト名」で作る
 * - シートに「原題作者」という列は不要
 * - Yahooの商品タイトルには韓国語原題を出さない
 * - アーティスト/監督/出演などは「クレジット種別」列で制御する
 * - 商品重複は「アラジン商品コード / JAN / yes24 / Kyobo / 商品名(原題)」の優先順位で見る
 * - JANコードが無い商品でも止まらない
 * - 予約中は既存の作品ID(W)(自動)を優先して、Worksを増やさず上書き更新する
 * - 共通関数は _kyoutuu ライブラリを使用
 */

const 設定_韓国音楽映像 = {
  マスターシート名: '韓国音楽映像',
  作品シート名: 'Works（韓国音楽映像）',
  言語マスター名: '言語マスター',
  カテゴリマスター名: 'カテゴリマスター',
  形態マスター名: '形態マスターシート',
  マスターファイルID: '1ZAy5wtQCq1ixl47MMEIGnrVWeae-F_NfN34lSOUsN5M',

  作品ヘッダー: [
    'WorksKey',
    '作品ID',
    '日本語タイトル',
    'アーティスト',
    '原題タイトル',
    '登録済み数',
    '最新数',
    '更新日時',
    '最新数(予約込み)',
    '予約更新日時'
  ],
  作品列数: 10,

  /**
   * WorksKey設計
   *
   * これは列名ではなく、共通.gsへ渡す内部設定値。
   * 実際には「商品名(タイトル) + アーティスト名」でWorksKeyを作る。
   */
  WorksKey方式: '原題作者',
  WorksKeyにカテゴリを含める: false,

  /**
   * 同名疑いチェック
   *
   * 韓国音楽映像ではOST・Various Artists・ヴァリアス・アーティスト系が多く、
   * 本・漫画向けの「同名疑い」警告が誤爆しやすいためOFF。
   */
  同名疑いチェック: false,

  /**
   * タイトル表示設定
   *
   * クレジット種別が入っている場合：
   *   アーティスト：IVE
   *   監督：パク・チャヌク
   *   出演：〇〇
   *
   * クレジット種別が空欄の場合：
   *   LP/CD/OSTだけカテゴリ別設定で「アーティスト」として出す
   *   DVD/Blu-ray/映画系は誤表記防止のため人物を出さない
   */
  タイトルに作者を表示: true,
  タイトル作者ラベル: '',

  タイトル作者ラベルカテゴリ別: {
    'LP': 'アーティスト',
    'CD': 'アーティスト',
    'OST': 'アーティスト',
    'アルバム': 'アーティスト',
    '音楽': 'アーティスト'
  },

  // Yahooの商品タイトルでは韓国語原題が文字化けする可能性があるため表示しない。
  // 商品名(タイトル)列はWorksKey用、商品名(原題)列は商品そのもの・重複判定用として残す。
  タイトルに原題を表示: false,

  /**
   * 予約中のWorks運用
   */
  予約時既存作品ID優先: true,
  予約時Works上書き: true,
  確定済みWorks上書き: false,

  /**
   * 商品重複判定
   */
  商品重複キー列優先順位: [
    'アラジン商品コード',
    'JANコード',
    'yes24商品コード',
    'Kyobo商品コード',
    '商品名(原題)'
  ],
  商品重複キー出力列: '重複チェックキー',

  色パレット: [
    '#b7e1cd',
    '#fce8b2',
    '#f4c7c3',
    '#c9daf8',
    '#d9d2e9',
    '#fce5cd',
    '#fff2cc',
    '#d9ead3',
    '#cfe2f3',
    '#f4cccc'
  ],

  監視列: [
    'アーティスト名',
    'クレジット種別',
    '日本語タイトル',
    '商品名(タイトル)',
    '商品名(原題)',
    '特典メモ',
    '言語',
    'カテゴリ',
    '売価',
    '原価',
    '配送パターン',
    'JANコード',
    '発売日',
    'アラジン商品コード',
    'アラジンURL',
    'yes24商品コード',
    'yes24URL',
    'Kyobo商品コード',
    'KyoboURL'
  ],

  列名: {
    発行チェック:     '発番発行',
    登録状況:         '登録状況',

    商品コード:       '商品コード(SKU)',
    タイトル:         'タイトル',

    作者:             'アーティスト名',
    クレジット種別:   'クレジット種別',
    日本語タイトル:   '日本語タイトル',

    Worksタイトル:     '商品名(タイトル)',
    原題:             '商品名(タイトル)',
    原題商品タイトル: '商品名(原題)',

    形態:             '',

    言語:             '言語',
    カテゴリ:         'カテゴリ',

    単巻数:           '',
    セット開始:       '',
    セット終了:       '',

    特典メモ:         '特典メモ',

    ISBN:             'JANコード',

    SKU自動:          'SKU(自動)',
    作品ID:           '作品ID(W)(自動)',
    コードステータス: '商品コードステータス',

    配送パターン:     '配送パターン',
    商品説明:         '商品説明',

    売価:             '売価',
    原価:             '原価',
    粗利益率:         '粗利益率',

    JANコード:         'JANコード',
    発売日:           '発売日',

    サイト商品コード: 'アラジン商品コード',
    博客來商品コード: 'アラジン商品コード',

    アラジン商品コード: 'アラジン商品コード',
    アラジンURL:       'アラジンURL',

    yes24商品コード:  'yes24商品コード',
    yes24URL:         'yes24URL',

    Kyobo商品コード:  'Kyobo商品コード',
    KyoboURL:         'KyoboURL',

    メイン画像URL:     'メイン画像URL',
    追加画像URL:       '追加画像URL',

    重複チェックキー: '重複チェックキー',

    登録日:           '登録日',
    登録者:           '登録者',
    備考:             '備考'
  }
};

/* ============================================================
 * 韓国音楽映像 メニューラッパー
 * ============================================================ */

function 韓国音楽映像_確定発行() {
  _kyoutuu.確定発行を実行(設定_韓国音楽映像);
}

function 韓国音楽映像_削除() {
  _kyoutuu.削除を実行(設定_韓国音楽映像);
}

function 韓国音楽映像_一括更新() {
  _kyoutuu.一括更新を実行(設定_韓国音楽映像);
}

function 韓国音楽映像_プルダウン更新() {
  _kyoutuu.プルダウン更新を実行(設定_韓国音楽映像);
}

/**
 * 既存シート用：
 * 「アーティスト名」の右隣に「クレジット種別」列がなければ自動追加する。
 */
function 韓国音楽映像_クレジット種別列を追加() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(設定_韓国音楽映像.マスターシート名);
  const ui = SpreadsheetApp.getUi();

  if (!sh) {
    ui.alert('韓国音楽映像シートが見つかりません');
    return;
  }

  let headers = sh.getRange(1, 1, 1, sh.getLastColumn())
    .getValues()[0]
    .map(v => String(v || '').trim());

  if (headers.includes('クレジット種別')) {
    韓国音楽映像_クレジット種別プルダウン作成();
    ui.alert('✅ 既に「クレジット種別」列があります。プルダウンを更新しました。');
    return;
  }

  const artistCol = headers.indexOf('アーティスト名') + 1;

  if (!artistCol) {
    ui.alert('「アーティスト名」列が見つかりません');
    return;
  }

  sh.insertColumnAfter(artistCol);
  sh.getRange(1, artistCol + 1)
    .setValue('クレジット種別')
    .setBackground('#3c78d8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sh.setColumnWidth(artistCol + 1, 130);

  韓国音楽映像_クレジット種別プルダウン作成();

  ui.alert('✅ 「アーティスト名」の右隣に「クレジット種別」列を追加しました');
}

/**
 * クレジット種別プルダウン作成
 */
function 韓国音楽映像_クレジット種別プルダウン作成() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(設定_韓国音楽映像.マスターシート名);

  if (!sh) {
    SpreadsheetApp.getUi().alert('韓国音楽映像シートが見つかりません');
    return;
  }

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
    .getValues()[0]
    .map(v => String(v || '').trim());

  const col = headers.indexOf('クレジット種別') + 1;

  if (!col) {
    SpreadsheetApp.getUi().alert('「クレジット種別」列が見つかりません');
    return;
  }

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([
      'アーティスト',
      '監督',
      '出演',
      '著',
      '原作',
      '関連'
    ], true)
    .setAllowInvalid(false)
    .build();

  const lastRow = Math.max(sh.getMaxRows(), 1000);
  sh.getRange(2, col, lastRow - 1, 1).setDataValidation(rule);

  SpreadsheetApp.getUi().alert('✅ クレジット種別のプルダウンを作成しました');
}

/**
 * 拡張機能やAPIで行を書き込んだあとに、
 * 指定行だけ商品コード・タイトル・Worksを再生成したい場合に使う。
 */
function 韓国音楽映像_指定行を再生成(開始行, 行数) {
  if (!開始行 || 開始行 < 2) {
    SpreadsheetApp.getUi().alert('開始行は2行目以降を指定してください');
    return;
  }

  _kyoutuu.指定行を再生成(
    設定_韓国音楽映像,
    開始行,
    行数 || 1
  );
}

/**
 * 現在選択している行だけ再生成する手動用。
 */
function 韓国音楽映像_選択行を再生成() {
  const sh = SpreadsheetApp.getActiveSheet();

  if (sh.getName() !== 設定_韓国音楽映像.マスターシート名) {
    SpreadsheetApp.getUi().alert('韓国音楽映像シートで実行してください');
    return;
  }

  const range = sh.getActiveRange();
  const 開始行 = range.getRow();
  const 行数 = range.getNumRows();

  if (開始行 < 2) {
    SpreadsheetApp.getUi().alert('ヘッダー行は再生成できません');
    return;
  }

  _kyoutuu.指定行を再生成(
    設定_韓国音楽映像,
    開始行,
    行数
  );

  SpreadsheetApp.getUi().alert(`✅ ${行数}行を再生成しました`);
}

/**
 * 列名チェック
 */
function 韓国音楽映像_列名チェック() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(設定_韓国音楽映像.マスターシート名);

  if (!sh) {
    SpreadsheetApp.getUi().alert(
      `シートが見つかりません: ${設定_韓国音楽映像.マスターシート名}`
    );
    return;
  }

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
    .getValues()[0]
    .map(v => String(v || '').trim());

  const 必要列 = [
    ...設定_韓国音楽映像.監視列,
    ...Object.values(設定_韓国音楽映像.列名),
    ...(設定_韓国音楽映像.商品重複キー列優先順位 || [])
  ].filter(Boolean);

  const 不足列 = [...new Set(必要列)]
    .filter(name => !headers.includes(name));

  if (不足列.length) {
    SpreadsheetApp.getUi().alert(
      '❌ 不足している列があります\n\n' + 不足列.join('\n')
    );
  } else {
    SpreadsheetApp.getUi().alert(
      '✅ 韓国音楽映像：列名は設定と一致しています'
    );
  }
}

/**
 * 自己更新フラグ解除
 */
function 韓国音楽映像_自己更新フラグ解除() {
  PropertiesService.getScriptProperties().deleteProperty('KYOUTUU_UPDATING');
  SpreadsheetApp.getUi().alert('✅ 自己更新フラグを解除しました');
}

/* ============================================================
 * 韓国音楽映像シート作成
 * ============================================================ */

function 韓国音楽映像シートを作成() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const シート名 = 設定_韓国音楽映像.マスターシート名;

  if (ss.getSheetByName(シート名)) {
    ui.alert(`「${シート名}」シートは既に存在します`);
    return;
  }

  const sh = ss.insertSheet(シート名);

  const ヘッダー = [
    '発番発行',
    '登録状況',
    '商品コード(SKU)',
    'タイトル',

    'アーティスト名',
    'クレジット種別',
    '日本語タイトル',
    '商品名(タイトル)',
    '商品名(原題)',
    '特典メモ',

    '言語',
    'カテゴリ',

    '売価',
    '原価',
    '粗利益率',
    '配送パターン',

    '商品説明',
    'JANコード',
    '発売日',

    'アラジン商品コード',
    'アラジンURL',

    'yes24商品コード',
    'yes24URL',

    'Kyobo商品コード',
    'KyoboURL',

    'メイン画像URL',
    '追加画像URL',

    'SKU(自動)',
    '作品ID(W)(自動)',
    '商品コードステータス',
    '重複チェックキー',

    '登録日',
    '登録者',
    '備考'
  ];

  sh.getRange(1, 1, 1, ヘッダー.length).setValues([ヘッダー]);

  const 自動列 = [
    '商品コード(SKU)',
    'タイトル',
    'SKU(自動)',
    '作品ID(W)(自動)',
    '商品コードステータス',
    '重複チェックキー',
    '登録日'
  ];

  const API列 = [
    '商品名(タイトル)',
    '商品名(原題)',
    '原価',
    'JANコード',
    '発売日',
    'アラジン商品コード',
    'アラジンURL',
    'yes24商品コード',
    'yes24URL',
    'Kyobo商品コード',
    'KyoboURL',
    'メイン画像URL',
    '追加画像URL'
  ];

  const 入力重要列 = [
    'アーティスト名',
    'クレジット種別',
    '日本語タイトル',
    '商品名(タイトル)',
    '特典メモ',
    '言語',
    'カテゴリ',
    '売価',
    '配送パターン'
  ];

  for (let i = 0; i < ヘッダー.length; i++) {
    const h = ヘッダー[i];
    let 色 = '#4a86e8';

    if (自動列.includes(h)) 色 = '#999999';
    if (API列.includes(h)) 色 = '#e69138';
    if (入力重要列.includes(h)) 色 = '#3c78d8';

    sh.getRange(1, i + 1)
      .setBackground(色)
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
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

  const creditCol = ヘッダー.indexOf('クレジット種別') + 1;
  const creditRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([
      'アーティスト',
      '監督',
      '出演',
      '著',
      '原作',
      '関連'
    ], true)
    .setAllowInvalid(false)
    .build();

  sh.getRange(2, creditCol, 最終行 - 1, 1).setDataValidation(creditRule);

  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$B2="未登録"')
      .setBackground('#ffff00')
      .setRanges([sh.getRange(2, 1, 最終行 - 1, ヘッダー.length)])
      .build()
  ]);

  const 幅設定 = {
    1: 80,    // 発番発行
    2: 90,    // 登録状況
    3: 130,   // 商品コード(SKU)
    4: 420,   // タイトル

    5: 180,   // アーティスト名
    6: 130,   // クレジット種別
    7: 220,   // 日本語タイトル
    8: 240,   // 商品名(タイトル)
    9: 300,   // 商品名(原題)
    10: 180,  // 特典メモ

    11: 90,   // 言語
    12: 100,  // カテゴリ

    13: 90,   // 売価
    14: 90,   // 原価
    15: 100,  // 粗利益率
    16: 130,  // 配送パターン

    17: 300,  // 商品説明
    18: 130,  // JANコード
    19: 120,  // 発売日

    20: 140,  // アラジン商品コード
    21: 220,  // アラジンURL

    22: 140,  // yes24商品コード
    23: 220,  // yes24URL

    24: 140,  // Kyobo商品コード
    25: 220,  // KyoboURL

    26: 220,  // メイン画像URL
    27: 260,  // 追加画像URL

    28: 130,  // SKU(自動)
    29: 130,  // 作品ID(W)(自動)
    30: 170,  // 商品コードステータス
    31: 190,  // 重複チェックキー

    32: 120,  // 登録日
    33: 100,  // 登録者
    34: 200   // 備考
  };

  Object.keys(幅設定).forEach(col => {
    sh.setColumnWidth(Number(col), 幅設定[col]);
  });

  sh.getRange(1, 1, 最終行, ヘッダー.length)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    .setVerticalAlignment('middle');

  ui.alert(
    '✅ 韓国音楽映像シートを作成しました\n\n' +
    '次にメニュー「韓国音楽映像 → プルダウン更新」を実行してください'
  );
}

