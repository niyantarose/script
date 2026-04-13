/**
 * 韓国グッズ.gs
 * 韓国グッズシート専用の設定・シート作成・onEdit・確定発行
 *
 * ★ フロー
 *   入力中（onEdit）
 *     → 親コード（予約）自動生成（JMEE-BLK-0001形式）
 *     → 作品略称をマスターから自動反映（未登録は ???）
 *     → 商品名（出品用）自動組み立て
 *     → コードステータス = '商品コード(予約)'
 *   確定発行ボタン
 *     → 親コードを確定
 *     → Works（韓国グッズ）に登録・集計
 *     → コードステータス = '商品コード(発行済み確定)'
 */

const 設定_韓国グッズ = {
  マスターシート名:   '韓国グッズ',
  作品シート名:       'Works（韓国グッズ）',
  作品略称マスター名: 'グッズ作品マスター',
  ショップマスター名: 'ショップマスター',
  ジャンルマスター名: 'グッズジャンルマスター',

  作品ヘッダー: ['WorksKey', '作品ID', '作品名', '作品略称', '登録数', '更新日時'],
  作品列数: 6,

  色パレット: ['#b7e1cd', '#fce8b2', '#f4c7c3', '#c9daf8', '#d9d2e9', '#fce5cd'],

  // 共通onEditのdispatcherが参照する（韓国グッズは専用処理のため空）
  監視列: [],

  列名: {
    登録状況:         '登録状況',
    ショップ名:       'ショップ名',
    親コード:         '親コード',
    コードステータス: 'コードステータス',
    商品名出品用:     '商品名（出品用）',
    作品名:           '作品名',
    ジャンル:         'ジャンル',
    作品略称:         '作品略称',
    商品名日本語:     '商品名（日本語）',
    補足項目:         '補足項目',
    売価:             '売価',
    配送パターン:     '配送パターン',
    登録者:           '登録者',
    備考:             '備考',
    商品名原題:       '商品名（原題）',
    購入URL:          '購入URL',
    メイン画像URL:    'メイン画像URL',
    追加画像URL:      '追加画像URL',
    原価:             '原価',
    ショップ商品ID:   'ショップ商品ID',
    重複チェックキー: '重複チェックキー',
    登録日:           '登録日'
  }
};

/* ============================================================
 * グッズジャンルマスター作成
 * ============================================================ */
function グッズジャンルマスターを作成() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const シート名 = 設定_韓国グッズ.ジャンルマスター名;
  if (ss.getSheetByName(シート名)) { ui.alert(`「${シート名}」シートは既に存在します`); return; }

  const sh = ss.insertSheet(シート名);
  const ヘッダー = ['ジャンル', '備考'];
  sh.getRange(1, 1, 1, ヘッダー.length).setValues([ヘッダー]);

  const 初期値 = [
    ['コミック', ''],
    ['ゲーム', ''],
    ['ノベル', ''],
    ['アニメ', ''],
    ['アイドル', ''],
    ['映画', ''],
    ['ドラマ', ''],
  ];
  sh.getRange(2, 1, 初期値.length, 2).setValues(初期値);

  const hr = sh.getRange(1, 1, 1, ヘッダー.length);
  hr.setBackground('#4a86e8').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 120);
  sh.setColumnWidth(2, 200);

  ui.alert('✅ グッズジャンルマスターを作成しました\n\n※ ジャンルは自由に追加・削除できます');
}

/* ============================================================
 * ショップマスター作成
 * ============================================================ */
function ショップマスターを作成() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const シート名 = 設定_韓国グッズ.ショップマスター名;
  if (ss.getSheetByName(シート名)) { ui.alert(`「${シート名}」シートは既に存在します`); return; }

  const sh = ss.insertSheet(シート名);
  const ヘッダー = ['ショップ名', 'コード', '備考'];
  sh.getRange(1, 1, 1, ヘッダー.length).setValues([ヘッダー]);

  const 初期値 = [
    ['JMEE', 'JMEE', ''],
    ['Weverse', 'WV', ''],
  ];
  sh.getRange(2, 1, 初期値.length, 3).setValues(初期値);

  const hr = sh.getRange(1, 1, 1, ヘッダー.length);
  hr.setBackground('#4a86e8').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, 3, 150);

  ui.alert('✅ ショップマスターを作成しました');
}

/* ============================================================
 * グッズ作品マスター作成
 * ============================================================ */
function グッズ作品マスターを作成() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const シート名 = 設定_韓国グッズ.作品略称マスター名;
  if (ss.getSheetByName(シート名)) { ui.alert(`「${シート名}」シートは既に存在します`); return; }

  const sh = ss.insertSheet(シート名);
  const ヘッダー = ['作品名', '作品略称', '備考'];
  sh.getRange(1, 1, 1, ヘッダー.length).setValues([ヘッダー]);

  const 初期値 = [
    ['枯れた花に涙を', 'KRH', ''],
    ['ブルーロック', 'BLK', ''],
  ];
  sh.getRange(2, 1, 初期値.length, 3).setValues(初期値);

  const hr = sh.getRange(1, 1, 1, ヘッダー.length);
  hr.setBackground('#4a86e8').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, 3, 160);

  ui.alert('✅ グッズ作品マスターを作成しました\n\n※ 新しい作品を登録する前にここに略称を追加してください');
}

/* ============================================================
 * Works（韓国グッズ）シート作成
 * ============================================================ */
function 韓国グッズWorksシートを作成() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const シート名 = 設定_韓国グッズ.作品シート名;
  if (ss.getSheetByName(シート名)) { ui.alert(`「${シート名}」シートは既に存在します`); return; }

  const sh = ss.insertSheet(シート名);
  sh.getRange(1, 1, 1, 設定_韓国グッズ.作品列数).setValues([設定_韓国グッズ.作品ヘッダー]);

  const hr = sh.getRange(1, 1, 1, 設定_韓国グッズ.作品列数);
  hr.setBackground('#cc0000').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  [150, 120, 200, 100, 80, 150].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  ui.alert('✅ Works（韓国グッズ）シートを作成しました');
}

/* ============================================================
 * 韓国グッズシート作成
 * ============================================================ */
function 韓国グッズシートを作成() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const シート名 = '韓国グッズ';
  if (ss.getSheetByName(シート名)) { ui.alert(`「${シート名}」シートは既に存在します`); return; }

  const sh = ss.insertSheet(シート名);

const ヘッダー = [
  // 基本・自動ゾーン
  '登録状況', 'ショップ名', '親コード', 'コードステータス', '商品名（出品用）',
  // 手動入力ゾーン
  '作品名', 'ジャンル', '作品略称', '商品名（日本語）', '補足項目', '特典メモ',
  '売価', '配送パターン', '登録者', '備考',
  // API/手動ゾーン
  '商品名（原題）', '購入URL', 'メイン画像URL', '追加画像URL', '原価',
  // 自動ゾーン
  'ショップ商品ID', '重複チェックキー', '登録日'
];

  sh.getRange(1, 1, 1, ヘッダー.length).setValues([ヘッダー]);

  const 手動列 = ['登録状況', 'ショップ名', '作品名', 'ジャンル', '商品名（日本語）', '補足項目', '特典メモ', '売価', '配送パターン', '登録者', '備考'];
  const 自動列 = ['親コード', 'コードステータス', '商品名（出品用）', '作品略称', 'ショップ商品ID', '重複チェックキー', '登録日'];
  const API列 =  ['商品名（原題）', '購入URL', 'メイン画像URL', '追加画像URL', '原価'];

  for (let i = 0; i < ヘッダー.length; i++) {
    const cell = sh.getRange(1, i + 1);
    let 色 = '#4a86e8';
    if (自動列.includes(ヘッダー[i])) 色 = '#999999';
    if (API列.includes(ヘッダー[i]))  色 = '#e69138';
    cell.setBackground(色).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  }

  sh.setRowHeight(1, 40);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(5); // 登録状況〜商品名（出品用）まで固定

  const 最終行 = 1000;

  // 登録状況プルダウン
  sh.getRange(2, 1, 最終行 - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['未登録', '登録済み', '売り切れ'], true).build()
  );

  // ショップ名プルダウン
  const ショップシート = ss.getSheetByName(設定_韓国グッズ.ショップマスター名);
  sh.getRange(2, 2, 最終行 - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(
        ショップシート ? ショップシート.getRange('A2:A100') : sh.getRange('A1'), true
      ).build()
  );

  // ジャンルプルダウン（マスターシート参照）
  const ジャンルシート = ss.getSheetByName(設定_韓国グッズ.ジャンルマスター名);
  const ジャンル列 = ヘッダー.indexOf('ジャンル') + 1;
  sh.getRange(2, ジャンル列, 最終行 - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(
        ジャンルシート ? ジャンルシート.getRange('A2:A100') : sh.getRange('A1'), true
      ).build()
  );

  // 配送パターンプルダウン
  const 配送列 = ヘッダー.indexOf('配送パターン') + 1;
  sh.getRange(2, 配送列, 最終行 - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['通常', '大型', '冷蔵', '冷凍'], true).build()
  );

  // 条件付き書式（未登録=黄色 / 予約=水色 / 確定=緑）
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2="未登録"')
      .setBackground('#ffff00')
      .setRanges([sh.getRange(2, 1, 最終行 - 1, ヘッダー.length)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$D2="商品コード(予約)"')
      .setBackground('#cfe2f3')
      .setRanges([sh.getRange(2, 1, 最終行 - 1, ヘッダー.length)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$D2="商品コード(発行済み確定)"')
      .setBackground('#d9ead3')
      .setRanges([sh.getRange(2, 1, 最終行 - 1, ヘッダー.length)])
      .build(),
  ]);

  // 列幅
  const 列幅マップ = {
  '登録状況': 80, 'ショップ名': 100, '親コード': 150, 'コードステータス': 160, '商品名（出品用）': 320,
  '作品名': 180, 'ジャンル': 100, '作品略称': 80, '商品名（日本語）': 180, '補足項目': 100,
  '特典メモ': 150, // ← 追加
  '売価': 80, '配送パターン': 100, '登録者': 80, '備考': 150,
  '商品名（原題）': 200, '購入URL': 220, 'メイン画像URL': 200, '追加画像URL': 200, '原価': 80,
  'ショップ商品ID': 140, '重複チェックキー': 150, '登録日': 120
};
  ヘッダー.forEach((h, i) => { if (列幅マップ[h]) sh.setColumnWidth(i + 1, 列幅マップ[h]); });

  ui.alert('✅ 韓国グッズシートを作成しました\n\n【色の意味】\n🟡 黄色 = 未登録\n🔵 水色 = 商品コード予約済み\n🟢 緑 = 確定発行済み\n\n【手順】\n1. ショップマスター・グッズ作品マスターに登録\n2. ショップ名・作品名・商品名を入力\n3. ① 商品コード確定発行 で確定');
}

/* ============================================================
 * 内部ヘルパー：マスター検索
 * ============================================================ */
function 韓国グッズ_作品マスターから取得_(作品名) {
  const ss = SpreadsheetApp.getActive();
  const マスター = ss.getSheetByName(設定_韓国グッズ.作品略称マスター名);
  if (!マスター || マスター.getLastRow() < 2) return null;
  const データ = マスター.getRange(2, 1, マスター.getLastRow() - 1, 3).getValues();
  for (const row of データ) {
    if (String(row[0]).trim() === String(作品名).trim()) {
      return { 作品略称: String(row[1]).trim() };
    }
  }
  return null;
}

function 韓国グッズ_ショップマスターから取得_(ショップ名) {
  const ss = SpreadsheetApp.getActive();
  const マスター = ss.getSheetByName(設定_韓国グッズ.ショップマスター名);
  if (!マスター || マスター.getLastRow() < 2) return null;
  const データ = マスター.getRange(2, 1, マスター.getLastRow() - 1, 2).getValues();
  for (const row of データ) {
    if (String(row[0]).trim() === String(ショップ名).trim()) {
      return { コード: String(row[1]).trim() };
    }
  }
  return null;
}

function 韓国グッズ_親コードを生成_(ショップ名, 作品名, 現在行) {
  const ショップ情報 = 韓国グッズ_ショップマスターから取得_(ショップ名);
  const 作品情報    = 韓国グッズ_作品マスターから取得_(作品名);
  if (!ショップ情報 || !作品情報) return null;

  const prefix = `${ショップ情報.コード}-${作品情報.作品略称}-`;
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('韓国グッズ');
  if (!sh || sh.getLastRow() < 2) return `${prefix}0001`;

  const 既存コード = sh.getRange(2, 3, sh.getLastRow() - 1, 1).getValues().flat();
  let 最大番号 = 0;
  既存コード.forEach((code, i) => {
    if (i + 2 === 現在行) return; // 現在行は除外
    if (code && String(code).startsWith(prefix)) {
      const n = parseInt(String(code).replace(prefix, ''), 10);
      if (!isNaN(n) && n > 最大番号) 最大番号 = n;
    }
  });
  return `${prefix}${String(最大番号 + 1).padStart(4, '0')}`;
}

function 韓国グッズ_出品用商品名を生成_(作品名, ジャンル, 商品名日本語, 補足項目, 商品名原題) {
  if (!作品名 || !商品名日本語) return '';
  const ジャンル部 = ジャンル ? ` ${String(ジャンル).trim()}` : '';
  let 名前 = `韓国${ジャンル部} グッズ 『${作品名} ${商品名日本語}`;
  if (補足項目 && String(補足項目).trim()) 名前 += ` ${String(補足項目).trim()}`;
  名前 += '』';
  if (商品名原題 && String(商品名原題).trim()) 名前 += ` ${String(商品名原題).trim()}`;
  return 名前;
}

/* ============================================================
 * onEdit（共通.gsのdispatcherから呼び出される）
 * ============================================================ */
function 韓国グッズ_onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== '韓国グッズ') return;
  const row = e.range.getRow();
  if (row < 2) return;
  if (自己更新中か_()) return;

  const lock = LockService.getDocumentLock();
  try { if (!lock.tryLock(5000)) return; } catch (_) { return; }

  try {
    自己更新を開始_();
    try {
      const ヘッダー = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      const 列 = {};
      ヘッダー.forEach((h, i) => { if (h) 列[String(h).trim()] = i + 1; });

      const col = e.range.getColumn();
      const get = (名前) => 列[名前] ? sh.getRange(row, 列[名前]).getValue() : '';
      const set = (名前, v) => { if (列[名前]) sh.getRange(row, 列[名前]).setValue(v); };

      // 確定済み行は触らない
      if (get('コードステータス') === '商品コード(発行済み確定)') return;

      const ショップ名   = get('ショップ名');
      const 作品名       = get('作品名');

      let 再生成フラグ = false;

      // ショップ名・作品名が変更されたとき
      if (col === 列['ショップ名'] || col === 列['作品名']) {
        if (作品名) {
          const 作品情報 = 韓国グッズ_作品マスターから取得_(作品名);
          set('作品略称', 作品情報 ? 作品情報.作品略称 : '???');
        }
        if (ショップ名 && 作品名) {
          const 親コード = 韓国グッズ_親コードを生成_(ショップ名, 作品名, row);
          if (親コード) {
            set('親コード', 親コード);
          } else {
            set('親コード', 'ERROR:マスター未登録');
            set('コードステータス', '入力中...');
          }
        }
        再生成フラグ = true;
      }

      // 商品名（日本語）・補足項目・ジャンル・特典メモが変更されたとき
      if (col === 列['商品名（日本語）'] || col === 列['補足項目'] ||
          col === 列['ジャンル']         || col === 列['特典メモ']) {
        再生成フラグ = true;
      }

      // 購入URLからショップ商品IDを自動抽出
      if (col === 列['購入URL']) {
        const url = get('購入URL');
        if (url) {
          const match = String(url).match(/product_no=(\d+)/);
          if (match) set('ショップ商品ID', match[1]);
          const shopInfo = 韓国グッズ_ショップマスターから取得_(ショップ名);
          const shopCode = shopInfo ? shopInfo.コード : ショップ名;
          if (match) set('重複チェックキー', `${shopCode}-${match[1]}`);
        }
      }

      // 出品用商品名・コードステータスを再生成
      if (再生成フラグ) {
        const 作品名最新       = get('作品名');
        const ジャンル最新     = get('ジャンル');
        const 商品名日本語最新 = get('商品名（日本語）');
        const 補足項目最新     = get('補足項目');
        const 商品名原題最新   = get('商品名（原題）');
        const 特典メモ最新     = get('特典メモ');
        const 親コード最新     = get('親コード');

        if (作品名最新 && 商品名日本語最新) {
          set('商品名（出品用）', 韓国グッズ_出品用商品名を生成_(
            作品名最新, ジャンル最新, 商品名日本語最新,
            補足項目最新, 商品名原題最新, 特典メモ最新
          ));
        }

        const ショップ名最新 = get('ショップ名');
        if (ショップ名最新 && 作品名最新 && 商品名日本語最新 && 親コード最新 && !String(親コード最新).startsWith('ERROR')) {
          set('コードステータス', '商品コード(予約)');
        } else {
          set('コードステータス', '入力中...');
        }
      }

      // 登録日（初回入力時のみ）
      if (作品名 && !get('登録日')) {
        set('登録日', new Date());
      }

    } finally {
      自己更新を終了_();
    }
  } finally {
    lock.releaseLock();
  }
}
/* ============================================================
 * ① 確定発行（韓国グッズ専用）
 * ============================================================ */
function 韓国グッズ_確定発行() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('韓国グッズ');
  if (!sh || sh.getLastRow() < 2) { ui.alert('データがありません'); return; }

  const ヘッダー = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const 列 = {};
  ヘッダー.forEach((h, i) => { if (h) 列[String(h).trim()] = i + 1; });

  const データ = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const Worksシート = 韓国グッズ_Worksシートを確保_(ss);

  const lock = LockService.getDocumentLock();
  lock.waitLock(60000);

  try {
    let 発行数 = 0;
    const Works更新Map = {}; // 作品名 → { 略称, 件数 }

    データ.forEach((row, i) => {
      const get = (名前) => 列[名前] ? row[列[名前] - 1] : '';
      const rowNum = i + 2;

      // 確定済みはWorks集計のみ
      if (get('コードステータス') === '商品コード(発行済み確定)') {
        const 作品名 = String(get('作品名')).trim();
        if (作品名) {
          if (!Works更新Map[作品名]) Works更新Map[作品名] = { 略称: String(get('作品略称')).trim(), 件数: 0 };
          Works更新Map[作品名].件数++;
        }
        return;
      }

      // 予約状態の行のみ確定発行
      if (get('コードステータス') !== '商品コード(予約)') return;

      const 親コード = String(get('親コード')).trim();
      if (!親コード || 親コード.startsWith('ERROR')) return;

      // ステータスを確定に更新
      sh.getRange(rowNum, 列['コードステータス']).setValue('商品コード(発行済み確定)');
      sh.getRange(rowNum, 列['登録状況']).setValue('未登録');
      発行数++;

      // Works集計用
      const 作品名 = String(get('作品名')).trim();
      if (作品名) {
        if (!Works更新Map[作品名]) {
          Works更新Map[作品名] = { 略称: String(get('作品略称')).trim(), 件数: 0 };
        }
        Works更新Map[作品名].件数++;
      }
    });

    // Works更新
    韓国グッズ_Worksを更新_(Worksシート, Works更新Map);

    ui.alert(`✅ 確定発行完了: ${発行数}件\n\n【色の意味】\n🟢 緑 = 確定発行済み`);
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
 * Works更新（内部）
 * ============================================================ */
function 韓国グッズ_Worksシートを確保_(ss) {
  let sh = ss.getSheetByName(設定_韓国グッズ.作品シート名);
  if (!sh) {
    sh = ss.insertSheet(設定_韓国グッズ.作品シート名);
    sh.getRange(1, 1, 1, 設定_韓国グッズ.作品列数).setValues([設定_韓国グッズ.作品ヘッダー]);
    const hr = sh.getRange(1, 1, 1, 設定_韓国グッズ.作品列数);
    hr.setBackground('#cc0000').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
    sh.setFrozenRows(1);
  }
  return sh;
}

function 韓国グッズ_Worksを更新_(Worksシート, Works更新Map) {
  const 既存Map = {};
  if (Worksシート.getLastRow() >= 2) {
    Worksシート.getRange(2, 1, Worksシート.getLastRow() - 1, 設定_韓国グッズ.作品列数).getValues()
      .forEach((row, i) => { if (row[2]) 既存Map[String(row[2]).trim()] = i + 2; });
  }

  Object.keys(Works更新Map).forEach(作品名 => {
    const info = Works更新Map[作品名];
    const WorksKey = 作品名.replace(/\s+/g, '').toLowerCase();

    if (既存Map[作品名]) {
      const 行 = 既存Map[作品名];
      Worksシート.getRange(行, 5).setValue(info.件数);
      Worksシート.getRange(行, 6).setValue(new Date());
    } else {
      const 新ID = 'KR-W-' + String(Worksシート.getLastRow()).padStart(4, '0');
      Worksシート.appendRow([WorksKey, 新ID, 作品名, info.略称, info.件数, new Date()]);
      既存Map[作品名] = Worksシート.getLastRow();
    }
  });
}

/* ============================================================
 * 重複チェック
 * ============================================================ */
function 韓国グッズ_重複チェック() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActive().getSheetByName('韓国グッズ');
  if (!sh || sh.getLastRow() < 2) { ui.alert('データがありません'); return; }

  const ヘッダー = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const キー列idx = ヘッダー.findIndex(h => h === '重複チェックキー');
  const データ = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();

  const キーMap = {};
  const 重複リスト = [];
  データ.forEach((row, i) => {
    const キー = String(row[キー列idx] || '').trim();
    if (!キー) return;
    if (キーMap[キー] !== undefined) {
      重複リスト.push({ キー, 行1: キーMap[キー] + 2, 行2: i + 2 });
    } else {
      キーMap[キー] = i;
    }
  });

  if (重複リスト.length === 0) {
    ui.alert('✅ 重複はありません！');
  } else {
    let msg = `⚠️ 重複が ${重複リスト.length} 件あります\n\n`;
    重複リスト.slice(0, 20).forEach(d => { msg += `・${d.キー}（${d.行1}行目 / ${d.行2}行目）\n`; });
    ui.alert(msg);
  }
}

/* ============================================================
 * Works初期化
 * ============================================================ */
function 韓国グッズ_Works初期化() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('警告', 'Works（韓国グッズ）を全削除します。続行？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(設定_韓国グッズ.作品シート名);
  if (sh) { const last = sh.getLastRow(); if (last > 1) sh.deleteRows(2, last - 1); }
  else { sh = ss.insertSheet(設定_韓国グッズ.作品シート名); }
  sh.getRange(1, 1, 1, 設定_韓国グッズ.作品列数).setValues([設定_韓国グッズ.作品ヘッダー]);
  ui.alert('✅ Works初期化完了（韓国グッズ）');
}

function 韓国グッズ_プルダウン更新() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('韓国グッズ');
  if (!sh) { SpreadsheetApp.getUi().alert('韓国グッズシートが見つかりません'); return; }

  const 最終行 = Math.max(sh.getLastRow() + 100, 200);
  const ヘッダー = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const 列 = {};
  ヘッダー.forEach((h, i) => { if (h) 列[String(h).trim()] = i + 1; });

  // ショップ名プルダウン（ショップマスター参照）
  const ショップシート = ss.getSheetByName(設定_韓国グッズ.ショップマスター名);
  if (ショップシート && 列['ショップ名']) {
    sh.getRange(2, 列['ショップ名'], 最終行 - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInRange(ショップシート.getRange('A2:A100'), true).build()
    );
  }

  // ジャンルプルダウン（グッズジャンルマスター参照）
  const ジャンルシート = ss.getSheetByName(設定_韓国グッズ.ジャンルマスター名);
  if (ジャンルシート && 列['ジャンル']) {
    sh.getRange(2, 列['ジャンル'], 最終行 - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInRange(ジャンルシート.getRange('A2:A100'), true).build()
    );
  }

  // 登録状況プルダウン
  if (列['登録状況']) {
    sh.getRange(2, 列['登録状況'], 最終行 - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['未登録', '登録済み', '売り切れ'], true).build()
    );
  }

  // 配送パターンプルダウン（マスター参照）
const 配送マスター = ss.getSheetByName('配送パターンマスター');
if (配送マスター && 列['配送パターン']) {
  const 最終行マスター = Math.max(配送マスター.getLastRow(), 2);
  sh.getRange(2, 列['配送パターン'], 最終行 - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(配送マスター.getRange(2, 1, 最終行マスター - 1, 1), true)
      .setAllowInvalid(false)
      .build()
  );
}

  SpreadsheetApp.getActiveSpreadsheet().toast('✅ 韓国グッズ プルダウン更新完了', '完了', 3);
}

function 韓国グッズ_出品用商品名を生成_(作品名, ジャンル, 商品名日本語, 補足項目, 商品名原題, 特典メモ) {
  if (!作品名 || !商品名日本語) return '';
  const ジャンル部 = ジャンル ? ` ${String(ジャンル).trim()}` : '';
  let 名前 = `韓国${ジャンル部} グッズ 『${作品名} ${商品名日本語}`;
  if (補足項目 && String(補足項目).trim()) 名前 += ` ${String(補足項目).trim()}`;
  名前 += '』';
  if (特典メモ && String(特典メモ).trim()) 名前 += ` ${String(特典メモ).trim()}`;
  if (商品名原題 && String(商品名原題).trim()) 名前 += ` ${String(商品名原題).trim()}`;
  return 名前;
}