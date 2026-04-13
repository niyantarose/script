/**
 * 台湾グッズ.gs
 * 台湾グッズシート専用の設定・シート作成・onEdit処理
 *
 * ★ 既存フレームワークとの違い
 * - 作品管理は 作品略称マスター（台湾）シートで行う
 * - 親コードは {言語}-{作品略称}-{連番4桁} 形式（例: TW-ZCR-0001）
 * - 共通のWorksフレームワーク（WorksKey等）は使用しない
 * - onEditは共通.gsのdispatcherから呼び出される専用関数
 */

const 設定_台湾グッズ = {
  マスターシート名:   '台湾グッズ',
  作品シート名:       'Works（台湾グッズ）',
  作品略称マスター名: '作品略称マスター（台湾）',

  言語リスト: ['TW', 'CN', 'TH', 'HK'],

  作品ヘッダー: ['WorksKey', '作品ID', '作品名（原題）', '作品名（日本語）', '作品略称', '登録数', '更新日時'],
  作品列数: 7,

  色パレット: ['#b7e1cd', '#fce8b2', '#f4c7c3', '#c9daf8', '#d9d2e9', '#fce5cd'],

  // 共通onEditのdispatcherが監視列を参照するが台湾グッズは専用処理のため空
  監視列: [],

  列名: {
    登録状況:         '登録状況',
    言語:             '言語',
    親コード:         '親コード',
    商品名出品用:     '商品名（出品用）',
    商品名日本語:     '商品名（日本語）',
    補足項目:         '補足項目',
    売価:             '売価',
    原価:             '原価',
    粗利益率: '粗利益率',
    配送パターン:     '配送パターン',
    登録者:           '登録者',
    備考:             '備考',
    作品名原題:       '作品名（原題）',
    作品名日本語sheet: '作品名（日本語）',
    商品名原題:       '商品名（原題）',
    博客來商品コード: '博客來商品コード',
    博客來URL:        '博客來URL',
    メイン画像URL:    'メイン画像URL',
    追加画像URL:      '追加画像URL',
    発売日:           '発売日',
    
    重複チェックキー: '重複チェックキー',
    登録日:           '登録日'
  }
};

/* ============================================================
 * 作品略称マスター（台湾）作成
 * ============================================================ */
function 台湾作品略称マスターを作成() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const シート名 = 設定_台湾グッズ.作品略称マスター名;
  if (ss.getSheetByName(シート名)) { ui.alert(`「${シート名}」シートは既に存在します`); return; }

  const sh = ss.insertSheet(シート名);
  const ヘッダー = ['作品名（原題）', '作品名（日本語）', '作品略称', '言語', '備考'];
  sh.getRange(1, 1, 1, ヘッダー.length).setValues([ヘッダー]);

  const 初期値 = [
    ['全知讀者視角', '全知的な読者の視点', 'ZCR', 'TW', ''],
    ['咒術迴戰',     '呪術廻戦',           'JJK', 'TW', ''],
  ];
  sh.getRange(2, 1, 初期値.length, ヘッダー.length).setValues(初期値);

  const ヘッダー範囲 = sh.getRange(1, 1, 1, ヘッダー.length);
  ヘッダー範囲.setBackground('#4a86e8').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, 5, 160);

  ui.alert('✅ 作品略称マスター（台湾）を作成しました\n\n※ 新しい作品を登録する前にここに略称を追加してください');
}

/* ============================================================
 * Works（台湾グッズ）シート作成
 * ============================================================ */
function 台湾グッズWorksシートを作成() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const シート名 = 設定_台湾グッズ.作品シート名;
  if (ss.getSheetByName(シート名)) { ui.alert(`「${シート名}」シートは既に存在します`); return; }

  const sh = ss.insertSheet(シート名);
  sh.getRange(1, 1, 1, 設定_台湾グッズ.作品列数).setValues([設定_台湾グッズ.作品ヘッダー]);

  const ヘッダー範囲 = sh.getRange(1, 1, 1, 設定_台湾グッズ.作品列数);
  ヘッダー範囲.setBackground('#6aa84f').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  [150, 120, 200, 200, 100, 80, 150].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  ui.alert('✅ Works（台湾グッズ）シートを作成しました');
}

/* ============================================================
 * 台湾グッズシート作成
 * ============================================================ */
function 台湾グッズシートを作成() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const シート名 = '台湾グッズ';
  if (ss.getSheetByName(シート名)) { ui.alert(`「${シート名}」シートは既に存在します`); return; }

  const sh = ss.insertSheet(シート名);

  const ヘッダー = [
    '登録状況', '言語', '親コード', '商品名（出品用）',
    '商品名（日本語）', '補足項目', '売価', '原価', '粗利益率','配送パターン', '登録者', '備考',
    '作品名（原題）', '作品名（日本語）', '商品名（原題）',
    '博客來商品コード', '博客來URL',
    'メイン画像URL', '追加画像URL', '発売日',
    '重複チェックキー', '登録日'
  ];

  sh.getRange(1, 1, 1, ヘッダー.length).setValues([ヘッダー]);

  const 手動列 =   ['登録状況', '言語', '商品名（日本語）', '補足項目', '売価', '配送パターン', '登録者', '備考'];
  const 自動列 =   ['親コード', '商品名（出品用）', '作品名（日本語）', '重複チェックキー', '登録日'];
  const Chrome列 = ['作品名（原題）', '商品名（原題）', '博客來商品コード', '博客來URL', 'メイン画像URL', '追加画像URL', '発売日', '原価'];

  for (let i = 0; i < ヘッダー.length; i++) {
    const cell = sh.getRange(1, i + 1);
    let 色 = '#4a86e8';
    if (自動列.includes(ヘッダー[i]))   色 = '#999999';
    if (Chrome列.includes(ヘッダー[i])) 色 = '#e69138';
    cell.setBackground(色).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  }

  sh.setRowHeight(1, 40);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(4);

  const 最終行 = 1000;

  sh.getRange(2, 1, 最終行 - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['未登録', '登録済み', '売り切れ'], true).build()
  );
  sh.getRange(2, 2, 最終行 - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(設定_台湾グッズ.言語リスト, true).build()
  );

  const 配送列 = ヘッダー.indexOf('配送パターン') + 1;
  sh.getRange(2, 配送列, 最終行 - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['通常', '大型', '冷蔵', '冷凍'], true).build()
  );

  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$A2="未登録"')
      .setBackground('#ffff00')
      .setRanges([sh.getRange(2, 1, 最終行 - 1, ヘッダー.length)])
      .build()
  ]);

  const 列幅マップ = {
    '登録状況': 80, '言語': 60, '親コード': 140, '商品名（出品用）': 320,
    '商品名（日本語）': 180, '補足項目': 100, '売価': 80, '配送パターン': 100,
    '登録者': 80, '備考': 150,
    '作品名（原題）': 200, '作品名（日本語）': 180, '商品名（原題）': 200,
    '博客來商品コード': 140, '博客來URL': 220,
    'メイン画像URL': 200, '追加画像URL': 200, '発売日': 100, '原価': 80,
    '重複チェックキー': 150, '登録日': 120
  };
  ヘッダー.forEach((h, i) => { if (列幅マップ[h]) sh.setColumnWidth(i + 1, 列幅マップ[h]); });

  ui.alert('✅ 台湾グッズシートを作成しました\n\n次の手順：\n1. 作品略称マスター（台湾）に作品を登録\n2. Works（台湾グッズ）シートを作成\n3. 入力開始！');
}

/* ============================================================
 * 内部ヘルパー：作品略称マスター検索
 * ============================================================ */
function 台湾_マスターから取得_(作品名原題) {
  const ss = SpreadsheetApp.getActive();
  const マスター = ss.getSheetByName(設定_台湾グッズ.作品略称マスター名);
  if (!マスター || マスター.getLastRow() < 2) return null;
  const データ = マスター.getRange(2, 1, マスター.getLastRow() - 1, 5).getValues();
  for (const row of データ) {
    if (String(row[0]).trim() === String(作品名原題).trim()) {
      return { 作品名日本語: String(row[1]).trim(), 略称: String(row[2]).trim(), 言語: String(row[3]).trim() };
    }
  }
  return null;
}

function 台湾グッズ_親コードを生成(言語, 作品名原題) {
  const マスター情報 = 台湾_マスターから取得_(作品名原題);
  if (!マスター情報 || !マスター情報.略称) return 'ERROR:略称未登録';

  const prefix = `${言語}-${マスター情報.略称}-`;
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('台湾グッズ');
  if (!sh || sh.getLastRow() < 2) return `${prefix}0001`;

  const 既存コード = sh.getRange(2, 3, sh.getLastRow() - 1, 1).getValues().flat();
  let 最大番号 = 0;
  既存コード.forEach(code => {
    if (code && String(code).startsWith(prefix)) {
      const n = parseInt(String(code).replace(prefix, ''), 10);
      if (!isNaN(n) && n > 最大番号) 最大番号 = n;
    }
  });
  return `${prefix}${String(最大番号 + 1).padStart(4, '0')}`;
}

function 台湾グッズ_出品用商品名を生成(作品名日本語, 商品名日本語, 補足項目, 商品名原題) {
  if (!作品名日本語 || 作品名日本語 === '???' || !商品名日本語) return '';
  let 出品名 = `台湾版 グッズ 『${作品名日本語} ${商品名日本語}`;
  if (補足項目 && String(補足項目).trim()) 出品名 += ` ${String(補足項目).trim()}`;
  出品名 += `』`;
  if (商品名原題 && String(商品名原題).trim()) 出品名 += ` ${String(商品名原題).trim()}`;
  return 出品名;
}

/* ============================================================
 * onEdit（共通.gsのdispatcherから呼び出される）
 * ============================================================ */
function 台湾グッズ_onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== '台湾グッズ') return;
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

      const 言語         = get('言語');
      const 作品名原題   = get('作品名（原題）');
      const 商品名日本語 = get('商品名（日本語）');
      const 補足項目     = get('補足項目');
      const 商品名原題v  = get('商品名（原題）');

      // 作品名（原題）または言語が入力されたとき
      if (col === 列['作品名（原題）'] || col === 列['言語']) {
        if (作品名原題) {
          const マスター = 台湾_マスターから取得_(作品名原題);
          const 作品名日本語 = マスター ? (マスター.作品名日本語 || '???') : '???';
          set('作品名（日本語）', 作品名日本語);

          if (言語 && !get('親コード')) {
            set('親コード', 台湾グッズ_親コードを生成(言語, 作品名原題));
          }

          if (商品名日本語) {
            set('商品名（出品用）', 台湾グッズ_出品用商品名を生成(作品名日本語, 商品名日本語, 補足項目, 商品名原題v));
          }
        }
      }

      // 商品名（日本語）または補足項目が変更されたとき
      if (col === 列['商品名（日本語）'] || col === 列['補足項目']) {
        const 作品名日本語 = get('作品名（日本語）');
        if (作品名日本語 && 作品名日本語 !== '???' && 商品名日本語) {
          set('商品名（出品用）', 台湾グッズ_出品用商品名を生成(作品名日本語, 商品名日本語, 補足項目, 商品名原題v));
        }
      }

      // 博客來商品コード入力 → 重複チェックキー生成
      if (col === 列['博客來商品コード']) {
        const コード = get('博客來商品コード');
        if (コード) set('重複チェックキー', String(コード).trim().toUpperCase());
      }

      // 登録日（作品名原題が入って登録日が空なら自動入力）
      if (作品名原題 && !get('登録日')) {
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
 * Works（台湾グッズ）更新
 * ============================================================ */
function 台湾グッズ_Worksを更新() {
  const ss = SpreadsheetApp.getActive();
  const グッズシート = ss.getSheetByName('台湾グッズ');
  const Worksシート  = ss.getSheetByName(設定_台湾グッズ.作品シート名);

  if (!グッズシート || !Worksシート) { SpreadsheetApp.getUi().alert('必要なシートが見つかりません'); return; }
  if (グッズシート.getLastRow() < 2) { SpreadsheetApp.getActiveSpreadsheet().toast('データがありません', '台湾グッズ', 3); return; }

  const ヘッダー = グッズシート.getRange(1, 1, 1, グッズシート.getLastColumn()).getValues()[0];
  const 作品原題idx  = ヘッダー.findIndex(h => h === '作品名（原題）');
  const 作品日本語idx = ヘッダー.findIndex(h => h === '作品名（日本語）');

  const データ = グッズシート.getRange(2, 1, グッズシート.getLastRow() - 1, グッズシート.getLastColumn()).getValues();
  const 作品Map = {};
  データ.forEach(row => {
    const 原題   = String(row[作品原題idx] || '').trim();
    const 日本語 = String(row[作品日本語idx] || '').trim();
    if (!原題) return;
    if (!作品Map[原題]) 作品Map[原題] = { 日本語, 件数: 0 };
    作品Map[原題].件数++;
  });

  const Works既存Map = {};
  if (Worksシート.getLastRow() >= 2) {
    Worksシート.getRange(2, 1, Worksシート.getLastRow() - 1, 設定_台湾グッズ.作品列数).getValues()
      .forEach((row, i) => { if (row[2]) Works既存Map[String(row[2]).trim()] = i + 2; });
  }

  Object.keys(作品Map).forEach(原題 => {
    const info = 作品Map[原題];
    if (Works既存Map[原題]) {
      const 行 = Works既存Map[原題];
      Worksシート.getRange(行, 6).setValue(info.件数);
      Worksシート.getRange(行, 7).setValue(new Date());
    } else {
      const 新ID = 'TW-W-' + String(Worksシート.getLastRow()).padStart(4, '0');
      const マスター情報 = 台湾_マスターから取得_(原題);
      Worksシート.appendRow([
        原題.replace(/\s+/g, '').toLowerCase(),
        新ID, 原題, info.日本語,
        マスター情報 ? マスター情報.略称 : '',
        info.件数, new Date()
      ]);
    }
  });

  SpreadsheetApp.getActiveSpreadsheet().toast('✅ Works（台湾グッズ）を更新しました', '完了', 3);
}

/* ============================================================
 * 重複チェック
 * ============================================================ */
function 台湾グッズ_重複チェック() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActive().getSheetByName('台湾グッズ');
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

function 台湾グッズ_プルダウン更新() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('台湾グッズ');
  if (!sh) { SpreadsheetApp.getUi().alert('台湾グッズシートが見つかりません'); return; }

  const 最終行 = Math.max(sh.getLastRow() + 100, 200);
  const ヘッダー = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const 列 = {};
  ヘッダー.forEach((h, i) => { if (h) 列[String(h).trim()] = i + 1; });

 // 言語プルダウン（言語マスターB列から取得）
  if (列['言語']) {
    const 言語マスターSh = ss.getSheetByName('言語マスター');
    let 言語リスト = 設定_台湾グッズ.言語リスト; // フォールバック
    if (言語マスターSh && 言語マスターSh.getLastRow() >= 2) {
      const codes = 言語マスターSh.getRange(2, 2, 言語マスターSh.getLastRow() - 1, 1)
        .getValues().flat().map(v => String(v || '').trim()).filter(v => v);
      if (codes.length > 0) 言語リスト = codes;
    }
    sh.getRange(2, 列['言語'], 最終行 - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(言語リスト, true).build()
    );
  }

  // 登録状況プルダウン
  if (列['登録状況']) {
    sh.getRange(2, 列['登録状況'], 最終行 - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(['未登録', '登録済み', '売り切れ'], true).build()
    );
  }

  // 配送パターンプルダウン（配送パターンマスターA列から取得）
  if (列['配送パターン']) {
    const 配送マスターSh = ss.getSheetByName('配送パターンマスター');
    let 配送リスト = ['通常', '大型', '冷蔵', '冷凍']; // フォールバック
    if (配送マスターSh && 配送マスターSh.getLastRow() >= 2) {
      const patterns = 配送マスターSh.getRange(2, 1, 配送マスターSh.getLastRow() - 1, 1)
        .getValues().flat().map(v => String(v || '').trim()).filter(v => v);
      if (patterns.length > 0) 配送リスト = patterns;
    }
    sh.getRange(2, 列['配送パターン'], 最終行 - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(配送リスト, true).build()
    );
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('✅ 台湾グッズ プルダウン更新完了', '完了', 3);
}