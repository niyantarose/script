/**
 * 韓国マンガ.gs
 * 韓国マンガシート専用の設定とメニューラッパー
 *
 * 方針：
 * - 台湾まんが・韓国書籍と同じ標準列名に寄せる
 * - Yahooの商品タイトルには韓国語原題を出さない
 * - 原題タイトルはWorksKey用・内部管理用として残す
 * - 原題商品タイトルは商品単位の重複判定用として残す
 * - 予約中は既存の作品ID(W)(自動)を優先して、Worksを増やさず上書き更新する
 * - 共通関数は _kyoutuu ライブラリを使用
 *
 * 注意：
 * - onEdit は Onopen.gs 側の onEdit_インストール型(e) を使用する
 * - このファイルには simple onEdit(e) を置かない
 */

const 設定_韓国マンガ = {
  マスターシート名: '韓国マンガ',
  作品シート名: 'Works（韓国マンガ）',
  言語マスター名: '言語マスター',
  カテゴリマスター名: 'カテゴリマスター',
  形態マスター名: '形態マスターシート',
  マスターファイルID: '1ZAy5wtQCq1ixl47MMEIGnrVWeae-F_NfN34lSOUsN5M',

  作品ヘッダー: [
    'WorksKey',
    '作品ID',
    '日本語タイトル',
    '作者',
    '原題タイトル',
    '登録済み巻',
    '最新巻',
    '更新日時',
    '最新巻(予約込み)',
    '予約更新日時'
  ],
  作品列数: 10,

  /**
   * タイトル表示設定
   *
   * Yahooの商品タイトルでは韓国語原題を表示しない。
   * 原題タイトル・原題商品タイトルは内部管理用として残す。
   */
  タイトルに作者を表示: true,
  タイトル作者ラベル: '著',
  タイトルに原題を表示: false,

  /**
   * 予約中のWorks運用
   *
   * 予約中に日本語タイトル・作者・原題タイトルを修正しても、
   * 既に作品ID(W)(自動)が入っている場合は、その作品IDを優先する。
   */
  予約時既存作品ID優先: true,
  予約時Works上書き: true,
  確定済みWorks上書き: false,

  /**
   * 商品重複判定
   *
   * WorksKeyとは別。
   * 同じ商品を二重登録していないかを見るための優先順位。
   */
  商品重複キー列優先順位: [
    'ISBN',
    'アラジン商品コード',
    'yes24商品コード',
    'Kyobo商品コード',
    'サイト商品コード',
    '原題商品タイトル',
    'リンク'
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

  /**
   * onEditで監視する列
   */
  監視列: [
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
    '配送パターン',
    'サイト商品コード',
    'アラジン商品コード',
    'アラジンURL',
    'yes24商品コード',
    'yes24URL',
    'Kyobo商品コード',
    'KyoboURL'
  ],

  /**
   * 共通処理が参照する列名マップ
   */
  列名: {
    発行チェック:     '発番発行',
    登録状況:         '登録状況',

    商品コード:       '商品コード(SKU)',
    タイトル:         'タイトル',

    作者:             '作者',
    日本語タイトル:   '日本語タイトル',

    // 作品・シリーズ単位の原題。
    // 同じシリーズとしてまとめたい場合は、この列を同じ値にする。
    原題:             '原題タイトル',

    // 商品単位の原題名。重複チェックなどで使う。
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

    売価:             '売価',
    原価:             '原価',
    粗利益率:         '粗利益率',
    配送パターン:     '配送パターン',
    商品説明:         '商品説明',

    サイト商品コード: 'サイト商品コード',

    アラジン商品コード: 'アラジン商品コード',
    アラジンURL:       'アラジンURL',

    yes24商品コード:  'yes24商品コード',
    yes24URL:         'yes24URL',

    Kyobo商品コード:  'Kyobo商品コード',
    KyoboURL:         'KyoboURL',

    メイン画像:       'メイン画像',
    追加画像:         '追加画像',
    発売日:           '発売日',

    作品ID:           '作品ID(W)(自動)',
    SKU自動:          'SKU(自動)',
    コードステータス: '商品コードステータス',
    重複チェックキー: '重複チェックキー',

    登録日:           '登録日',
    登録者:           '登録者',
    備考:             '備考'
  }
};

/* ============================================================
 * 韓国マンガ メニューラッパー
 * ============================================================ */

function 韓国マンガ_確定発行() {
  _kyoutuu.確定発行を実行(設定_韓国マンガ);
}

function 韓国マンガ_削除() {
  _kyoutuu.削除を実行(設定_韓国マンガ);
}

function 韓国マンガ_一括更新() {
  _kyoutuu.一括更新を実行(設定_韓国マンガ);
}

function 韓国マンガ_プルダウン更新() {
  _kyoutuu.プルダウン更新を実行(設定_韓国マンガ);
}

/* ============================================================
 * Works管理
 * ============================================================ */

function 韓国マンガ_Works危険操作を停止_(操作名) {
  SpreadsheetApp.getUi().alert(
    '🔒 この操作は封印されています',
    `「${操作名}」は作品IDずれ事故（登録済み商品やYahoo上のコードと食い違う）の原因になるため、\n` +
    '2026-07-15に停止しました（台湾側と同じ運用）。\n\n' +
    '・作品IDは永久ID（欠番はそのまま。採番はハイウォーターで安全に継続します）\n' +
    '・どうしても必要な場合は管理者に相談してください',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function 韓国マンガ_Works初期化() {
  // 【封印】Works全削除は、商品行に残る作品ID・SKUを宙に浮かせる破壊操作のため停止。
  韓国マンガ_Works危険操作を停止_('Works初期化（全削除）');
}

function 韓国マンガ_WorksKey再正規化() {
  const ui = SpreadsheetApp.getUi();
  const 作品シート = SpreadsheetApp.getActive()
    .getSheetByName(設定_韓国マンガ.作品シート名);

  if (!作品シート || 作品シート.getLastRow() < 2) {
    ui.alert('Worksにデータがありません');
    return;
  }

  if (
    ui.alert(
      '確認',
      '全WorksKeyを再計算し重複を統合します。続行しますか？',
      ui.ButtonSet.OK_CANCEL
    ) !== ui.Button.OK
  ) {
    return;
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const r = _kyoutuu.WorksKey再正規化を実行(
      作品シート,
      設定_韓国マンガ
    );

    ui.alert(
      `✅ 完了\nキー更新: ${r.キー更新数}\n重複統合: ${r.統合数}\n削除: ${r.削除行数}`
    );
  } finally {
    lock.releaseLock();
  }
}

function 韓国マンガ_重複統合() {
  const ui = SpreadsheetApp.getUi();
  const 作品シート = SpreadsheetApp.getActive()
    .getSheetByName(設定_韓国マンガ.作品シート名);

  if (!作品シート || 作品シート.getLastRow() < 2) {
    ui.alert('Worksにデータがありません');
    return;
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const 重複リスト = _kyoutuu.Works重複を検出(
      作品シート,
      設定_韓国マンガ
    );

    if (重複リスト.length === 0) {
      ui.alert('✅ 重複はありません！');
      return;
    }

    let レポート = `🔍 重複検出: ${重複リスト.length}件\n\n`;

    for (const dup of 重複リスト) {
      レポート += `【${dup.行配列[0].データ[2]}】 著：${dup.行配列[0].データ[3]}\n`;

      for (const item of dup.行配列) {
        レポート += `  ID:${item.データ[1]} (行${item.行})\n`;
      }

      レポート += '\n';
    }

    if (
      ui.alert(
        '重複チェック結果',
        レポート + '\n自動統合しますか？',
        ui.ButtonSet.YES_NO
      ) !== ui.Button.YES
    ) {
      return;
    }

    const r = _kyoutuu.WorksKey再正規化を実行(
      作品シート,
      設定_韓国マンガ
    );

    ui.alert(`✅ 統合完了\n統合: ${r.統合数}件, 削除: ${r.削除行数}行`);
  } finally {
    lock.releaseLock();
  }
}

function 韓国マンガ_ID振り直し() {
  // 【封印】ID振り直しは登録済み商品・外部コードと食い違う「ずれ」の直接原因のため停止。
  // （ライブラリ側の WorksID振り直しを実行 も無効化済み）
  韓国マンガ_Works危険操作を停止_('Works ID振り直し');
}

function 韓国マンガ_孤立削除() {
  // 【封印】Works行の削除操作は台湾側の封印と揃えて停止（採番はハイウォーターで保護済みだが、
  // 作品メタデータの喪失と誤削除のリスクが残るため）。
  韓国マンガ_Works危険操作を停止_('Works孤立エントリー削除');
}

/* ============================================================
 * 韓国マンガシート作成
 * ============================================================ */

function 韓国マンガシートを作成() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const シート名 = 設定_韓国マンガ.マスターシート名;

  if (ss.getSheetByName(シート名)) {
    ui.alert(
      `「${シート名}」シートは既に存在します。\n\n` +
      '作り直す場合は、先に既存シート名を「韓国マンガ_旧」などに変更してください。'
    );
    return;
  }

  const sh = ss.insertSheet(シート名);

  const ヘッダー = [
    '発番発行',
    '登録状況',
    '商品コード(SKU)',
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

    'アラジン商品コード',
    'アラジンURL',
    'yes24商品コード',
    'yes24URL',
    'Kyobo商品コード',
    'KyoboURL',

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

  sh.getRange(1, 1, 1, ヘッダー.length).setValues([ヘッダー]);

  const 自動列 = [
    '商品コード(SKU)',
    'タイトル',
    '作品ID(W)(自動)',
    'SKU(自動)',
    '商品コードステータス',
    '重複チェックキー',
    '登録日'
  ];

  const API列 = [
    '原題商品タイトル',
    'リンク',
    'ISBN',
    '原価',
    'サイト商品コード',
    'アラジン商品コード',
    'アラジンURL',
    'yes24商品コード',
    'yes24URL',
    'Kyobo商品コード',
    'KyoboURL',
    'メイン画像',
    '追加画像',
    '発売日'
  ];

  const 入力重要列 = [
    '作者',
    '日本語タイトル',
    '原題タイトル',
    '形態(通常/初回限定/特装)',
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

  // 発番発行チェックボックス
  sh.getRange(2, 1, 最終行 - 1, 1).insertCheckboxes();

  // 登録状況プルダウン
  sh.getRange(2, 2, 最終行 - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['未登録', '登録済み'], true)
      .build()
  );

  // 未登録行の背景色
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$B2="未登録"')
      .setBackground('#ffff00')
      .setRanges([sh.getRange(2, 1, 最終行 - 1, ヘッダー.length)])
      .build()
  ]);

  const 幅設定 = {
    1: 80,
    2: 90,
    3: 130,
    4: 420,

    5: 180,
    6: 260,
    7: 240,
    8: 300,
    9: 220,

    10: 180,
    11: 90,
    12: 120,

    13: 90,
    14: 150,
    15: 150,

    16: 180,
    17: 150,

    18: 90,
    19: 90,
    20: 100,
    21: 130,

    22: 300,
    23: 150,

    24: 150,
    25: 220,
    26: 150,
    27: 220,
    28: 150,
    29: 220,

    30: 220,
    31: 260,
    32: 120,

    33: 130,
    34: 130,
    35: 170,
    36: 190,

    37: 120,
    38: 100,
    39: 200
  };

  Object.keys(幅設定).forEach(col => {
    sh.setColumnWidth(Number(col), 幅設定[col]);
  });

  sh.getRange(1, 1, 最終行, ヘッダー.length)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    .setVerticalAlignment('middle');

  ui.alert(
    '✅ 韓国マンガシートを標準列名で作成しました\n\n' +
    '次にメニュー「韓国マンガ → プルダウン更新」を実行してください'
  );
}

/* ============================================================
 * 既存シートのヘッダー色更新
 * ============================================================ */

function 韓国マンガヘッダー色更新() {
  const sh = SpreadsheetApp.getActive().getSheetByName('韓国マンガ');
  if (!sh) return;

  const ヘッダー = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

  const 自動列 = [
    '商品コード(SKU)',
    'タイトル',
    '作品ID(W)(自動)',
    'SKU(自動)',
    '商品コードステータス',
    '重複チェックキー',
    '登録日'
  ];

  const API列 = [
    '原題商品タイトル',
    'リンク',
    'ISBN',
    '原価',
    'サイト商品コード',
    'アラジン商品コード',
    'アラジンURL',
    'yes24商品コード',
    'yes24URL',
    'Kyobo商品コード',
    'KyoboURL',
    'メイン画像',
    '追加画像',
    '発売日'
  ];

  const 入力重要列 = [
    '作者',
    '日本語タイトル',
    '原題タイトル',
    '形態(通常/初回限定/特装)',
    '言語',
    'カテゴリ',
    '売価',
    '配送パターン'
  ];

  for (let i = 0; i < ヘッダー.length; i++) {
    const h = String(ヘッダー[i] || '').trim();
    const cell = sh.getRange(1, i + 1);

    let 色 = '#4a86e8';

    if (自動列.includes(h)) 色 = '#999999';
    if (API列.includes(h)) 色 = '#e69138';
    if (入力重要列.includes(h)) 色 = '#3c78d8';

    cell
      .setBackground(色)
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  }

  Logger.log('✅ 韓国マンガヘッダー色を更新しました');
}