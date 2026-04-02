/**
 * 台湾グッズ.gs
 * 台湾グッズシート専用の設定・シート作成・onEdit・確定発行
 * ※ 共通関数は _kyoutuu ライブラリを使用
 *
 * 【確定発行の仕組み】
 * 発番発行列（チェックボックス）をONにした行だけが対象
 * 重複チェック：① 博客來商品コード → ② 商品名（原題）
 * 重複があればブロック（行番号をアラートで表示）
 * 問題なければ発行 → 登録状況=未登録、チェックOFF、Worksに登録
 */

const 設定_台湾グッズ = {
  マスターシート名:   '台湾グッズ',
  作品シート名:       'Works（台湾グッズ）',
  作品略称マスター名: '作品略称マスター（台湾）',
  マスターファイルID: '1ZAy5wtQCq1ixl47MMEIGnrVWeae-F_NfN34lSOUsN5M',

  作品ヘッダー: ['WorksKey', '作品ID', '作品名（原題）', '作品名（日本語）', '作品略称', '登録数', '更新日時'],
  作品列数: 7,

  列名: {
    登録状況:         '登録状況',
    発番発行:         '発番発行',
    言語:             '言語',
    親コード:         '親コード',
    商品名出品用:     '商品名（出品用）',
    売価:             '売価',
    配送パターン:     '配送パターン',
    登録者:           '登録者',
    原価:             '原価',
    商品説明:         '商品説明',
    商品名日本語:     '商品名（日本語）',
    作品名原題:       '作品名（原題）',
    作品名日本語:     '作品名（日本語）',
    特典メモ:         '特典メモ',
    補足項目:         '補足項目',
    商品名原題:       '商品名（原題）',
    博客來商品コード: '博客來商品コード',
    博客來URL:        '博客來URL',
    メイン画像URL:    'メイン画像URL',
    追加画像URL:      '追加画像URL',
    発売日:           '発売日',
    重複チェックキー: '重複チェックキー',
    登録日:           '登録日',
    粗利益率:         '粗利益率'
  }
};

/* ============================================================
 * ヘッダー色定義（統一カラーパレット）
 * 🟢 緑  #6aa84f … コード＋タイトル両方に直結
 * 🟡 黄  #f1c232 … タイトルのみに直結
 * ⬜ 灰  #999999 … 自動生成（GASが書き込む）
 * 🔵 青  #4a86e8 … 手動入力（生成に非直結）
 * 🟠 橙  #e69138 … API/外部から入力
 * ============================================================ */
const 台湾グッズ_ヘッダー色 = {
  '登録状況':         '#4a86e8', // 青
  '発番発行':         '#cc0000', // 赤（発行操作列）
  '言語':             '#6aa84f', // 緑
  '親コード':         '#999999', // 灰
  '商品名（出品用）': '#999999', // 灰
  '売価':             '#4a86e8', // 青
  '配送パターン':     '#4a86e8', // 青
  '登録者':           '#4a86e8', // 青
  '原価':             '#e69138', // 橙
  '商品説明':         '#e69138', // 橙
  '商品名（日本語）': '#f1c232', // 黄
  '作品名（原題）':   '#6aa84f', // 緑
  '作品名（日本語）': '#999999', // 灰（自動）
  '特典メモ':         '#f1c232', // 黄
  '補足項目':         '#f1c232', // 黄
  '商品名（原題）':   '#e69138', // 橙
  '博客來商品コード': '#e69138', // 橙
  '博客來URL':        '#e69138', // 橙
  'メイン画像URL':    '#e69138', // 橙
  '追加画像URL':      '#e69138', // 橙
  '発売日':           '#e69138', // 橙
  '重複チェックキー': '#999999', // 灰
  '登録日':           '#999999', // 灰
  '粗利益率':         '#999999', // 灰
};

/* ============================================================
 * マスター共通ファイル取得
 * ============================================================ */
function 台湾グッズ_マスターSSを取得_() {
  return SpreadsheetApp.openById(設定_台湾グッズ.マスターファイルID);
}

/* ============================================================
 * 作品略称マスター（台湾）作成
 * ============================================================ */
function 台湾グッズ作品マスターを作成() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const シート名 = 設定_台湾グッズ.作品略称マスター名;
  if (ss.getSheetByName(シート名)) { ui.alert(`「${シート名}」シートは既に存在します`); return; }

  const sh = ss.insertSheet(シート名);
  const ヘッダー = ['作品名（原題）', '作品名（日本語）', '作品略称', '備考'];
  sh.getRange(1, 1, 1, ヘッダー.length).setValues([ヘッダー]);

  const hr = sh.getRange(1, 1, 1, ヘッダー.length);
  hr.setBackground('#4a86e8').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  [200, 200, 100, 160].forEach((w, i) => sh.setColumnWidth(i + 1, w));

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

  const hr = sh.getRange(1, 1, 1, 設定_台湾グッズ.作品列数);
  hr.setBackground('#cc0000').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
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
    '発番発行', '登録状況', '言語', '親コード', '商品名（出品用）',
    '売価', '配送パターン', '登録者',
    '原価', '商品説明',
    '商品名（日本語）', '作品名（原題）', '作品名（日本語）',
    '特典メモ', '補足項目',
    '商品名（原題）', '博客來商品コード', '博客來URL', 'メイン画像URL', '追加画像URL', '発売日',
    '重複チェックキー', '登録日', '粗利益率'
  ];

  sh.getRange(1, 1, 1, ヘッダー.length).setValues([ヘッダー]);

  // 統一カラーパレットで色付け
  ヘッダー.forEach((h, i) => {
    const 色 = 台湾グッズ_ヘッダー色[h] || '#cccccc';
    sh.getRange(1, i + 1).setBackground(色).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  });

  sh.setRowHeight(1, 40);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(5);

  const 最終行 = 1000;

  // 発番発行チェックボックス（A列）
  sh.getRange(2, 1, 最終行 - 1, 1).insertCheckboxes();

  // 登録状況プルダウン（B列）
  sh.getRange(2, 2, 最終行 - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['未登録', '登録済み', '売り切れ'], true).build()
  );

  // 条件付き書式
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$B2="未登録"')
      .setBackground('#ffff00')
      .setRanges([sh.getRange(2, 1, 最終行 - 1, ヘッダー.length)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$B2="登録済み"')
      .setBackground('#d9ead3')
      .setRanges([sh.getRange(2, 1, 最終行 - 1, ヘッダー.length)])
      .build(),
  ]);

  // 列幅
  const 列幅マップ = {
    '登録状況': 80, '発番発行': 60, '言語': 60, '親コード': 160, '商品名（出品用）': 320,
    '売価': 80, '配送パターン': 100, '登録者': 80,
    '原価': 80, '商品説明': 200,
    '商品名（日本語）': 180, '作品名（原題）': 200, '作品名（日本語）': 180,
    '特典メモ': 150, '補足項目': 100,
    '商品名（原題）': 200, '博客來商品コード': 140, '博客來URL': 220,
    'メイン画像URL': 200, '追加画像URL': 200, '発売日': 100,
    '重複チェックキー': 150, '登録日': 120, '粗利益率': 80
  };
  ヘッダー.forEach((h, i) => { if (列幅マップ[h]) sh.setColumnWidth(i + 1, 列幅マップ[h]); });

  ui.alert('✅ 台湾グッズシートを作成しました\n\n【発行の手順】\n① 行を入力する\n② 発番発行列にチェックを入れる\n③ 確定発行ボタンを押す\n\n【色の意味（ヘッダー）】\n🟢 緑 = コード生成に直結\n🟡 黄 = タイトル生成に直結\n⬜ 灰 = 自動生成\n🔵 青 = 手動入力\n🟠 橙 = API/外部入力\n🔴 赤 = 発行操作列\n\n【行の色】\n🟡 黄色 = 未登録\n🟢 緑 = 登録済み');
}

/* ============================================================
 * 内部ヘルパー：マスター検索
 * ============================================================ */
function 台湾グッズ_作品マスターから取得_(作品名原題) {
  const ss = SpreadsheetApp.getActive();
  const マスター = ss.getSheetByName(設定_台湾グッズ.作品略称マスター名);
  if (!マスター || マスター.getLastRow() < 2) return null;
  const データ = マスター.getRange(2, 1, マスター.getLastRow() - 1, 4).getValues();
  for (const row of データ) {
    if (String(row[0]).trim() === String(作品名原題).trim()) {
      return { 作品名日本語: String(row[1]).trim(), 作品略称: String(row[2]).trim() };
    }
  }
  return null;
}

function 台湾グッズ_親コードを生成_(言語, 作品名原題, 現在行) {
  const 作品情報 = 台湾グッズ_作品マスターから取得_(作品名原題);
  if (!作品情報) return null;

  const langCode = String(言語 || 'TW').trim().toUpperCase();
  const prefix = `${langCode}-${作品情報.作品略称}-`;
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('台湾グッズ');
  if (!sh || sh.getLastRow() < 2) return `${prefix}0001`;

  const 既存コード = sh.getRange(2, 4, sh.getLastRow() - 1, 1).getValues().flat();
  let 最大番号 = 0;
  既存コード.forEach((code, i) => {
    if (i + 2 === 現在行) return;
    if (code && String(code).startsWith(prefix)) {
      const n = parseInt(String(code).replace(prefix, ''), 10);
      if (!isNaN(n) && n > 最大番号) 最大番号 = n;
    }
  });
  return `${prefix}${String(最大番号 + 1).padStart(4, '0')}`;
}

function 台湾グッズ_出品用商品名を生成_(言語, 作品名日本語, 商品名日本語, 補足項目, 商品名原題, 特典メモ) {
  if (!作品名日本語 || !商品名日本語) return '';
  const 言語表記 = String(言語 || '').trim();
  let 名前 = 言語表記 === 'CN' ? '中国 グッズ ' : '台湾 グッズ ';
  名前 += `『${作品名日本語} ${商品名日本語}`;
  if (補足項目 && String(補足項目).trim()) 名前 += ` ${String(補足項目).trim()}`;
  名前 += '』';
  if (特典メモ && String(特典メモ).trim()) 名前 += ` ${String(特典メモ).trim()}`;
  if (商品名原題 && String(商品名原題).trim()) 名前 += ` ${String(商品名原題).trim()}`;
  return 名前;
}

/* ============================================================
 * onEdit
 * ============================================================ */
function 台湾グッズ_onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== '台湾グッズ') return;
  const row = e.range.getRow();
  if (row < 2) return;
  if (_kyoutuu.自己更新中か_()) return;

  const lock = LockService.getDocumentLock();
  try { if (!lock.tryLock(5000)) return; } catch (_) { return; }

  try {
    _kyoutuu.自己更新を開始_();
    try {
      const ヘッダー = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      const 列 = {};
      ヘッダー.forEach((h, i) => { if (h) 列[String(h).trim()] = i + 1; });

      const col = e.range.getColumn();
      const get = (名前) => 列[名前] ? sh.getRange(row, 列[名前]).getValue() : '';
      const set = (名前, v) => { if (列[名前]) sh.getRange(row, 列[名前]).setValue(v); };

      const 言語       = get('言語');
      const 作品名原題 = get('作品名（原題）');
      let 再生成フラグ = false;

      // 言語 or 作品名（原題）変更 → コード生成＋作品名日本語補完
      if (col === 列['言語'] || col === 列['作品名（原題）']) {
        if (作品名原題) {
          const 作品情報 = 台湾グッズ_作品マスターから取得_(作品名原題);
          set('作品名（日本語）', 作品情報 ? 作品情報.作品名日本語 : '???');
        }
        if (言語 && 作品名原題) {
          const 親コード = 台湾グッズ_親コードを生成_(言語, 作品名原題, row);
          set('親コード', 親コード || 'ERROR:マスター未登録');
        }
        再生成フラグ = true;
      }

      if (col === 列['商品名（日本語）'] || col === 列['補足項目'] || col === 列['特典メモ']) {
        再生成フラグ = true;
      }

      // 博客來URLからコード自動抽出
      if (col === 列['博客來URL']) {
        const url = get('博客來URL');
        if (url) {
          const match = String(url).match(/\/product\/([A-Z0-9]+)/i);
          if (match) {
            set('博客來商品コード', match[1]);
            set('重複チェックキー', match[1]);
          }
        }
      }

      if (再生成フラグ) {
        const 作品名日本語最新 = get('作品名（日本語）');
        const 商品名日本語最新 = get('商品名（日本語）');
        if (作品名日本語最新 && 作品名日本語最新 !== '???' && 商品名日本語最新) {
          set('商品名（出品用）', 台湾グッズ_出品用商品名を生成_(
            get('言語'), 作品名日本語最新, 商品名日本語最新,
            get('補足項目'), get('商品名（原題）'), get('特典メモ')
          ));
        }
      }

      // 粗利益率自動計算（TWDレート）
      if (col === 列['売価'] || col === 列['原価']) {
        const 売価 = parseFloat(get('売価')) || 0;
        const 原価TWD = parseFloat(get('原価')) || 0;
        if (売価 > 0 && 原価TWD > 0) {
          try {
            const レート = _kyoutuu.為替レートを取得_('TWD');
            if (レート > 0) {
              const 粗利益率 = Math.round(((売価 - 原価TWD * レート) / 売価) * 1000) / 10;
              set('粗利益率', 粗利益率 / 100);
              if (列['粗利益率']) sh.getRange(row, 列['粗利益率']).setNumberFormat('0.0%');
            }
          } catch (_) {}
        }
      }

      if (作品名原題 && !get('登録日')) set('登録日', new Date());

    } finally {
      _kyoutuu.自己更新を終了_();
    }
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
 * 確定発行（台湾グッズ専用）
 *
 * 【仕様】
 * ① 発番発行列（チェックボックス）がONの行のみ対象
 * ② 重複チェック
 *    - 博客來商品コードが既存と一致 → ブロック
 *    - 商品名（原題）が既存と一致   → ブロック
 * ③ 問題なければ発行
 *    - 登録状況 = 未登録
 *    - 発番発行チェック = OFF
 *    - Worksシートに登録
 * ============================================================ */
function 台湾グッズ_確定発行() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('台湾グッズ');
  if (!sh || sh.getLastRow() < 2) { ui.alert('データがありません'); return; }

  const ヘッダー = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const 列 = {};
  ヘッダー.forEach((h, i) => { if (h) 列[String(h).trim()] = i + 1; });

  if (!列['発番発行']) { ui.alert('発番発行列が見つかりません'); return; }

  const 全データ = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const get行 = (row, 名前) => 列[名前] ? row[列[名前] - 1] : '';

  // チェックONの行を抽出
  const 対象行リスト = [];
  全データ.forEach((row, i) => {
    if (get行(row, '発番発行') === true) {
      対象行リスト.push({ row, rowNum: i + 2 });
    }
  });

  if (対象行リスト.length === 0) {
    ui.alert('発番発行列にチェックが入っている行がありません');
    return;
  }

  // 既存データ（発行済み行）からチェック用Mapを作成
  const 既存博客來Map = {};
  const 既存商品名Map = {};
  全データ.forEach((row, i) => {
    const rowNum = i + 2;
    const 博客來コード = String(get行(row, '博客來商品コード') || '').trim();
    const 商品名原題   = String(get行(row, '商品名（原題）') || '').trim();
    if (博客來コード) {
      if (!既存博客來Map[博客來コード]) 既存博客來Map[博客來コード] = [];
      既存博客來Map[博客來コード].push(rowNum);
    }
    if (商品名原題) {
      if (!既存商品名Map[商品名原題]) 既存商品名Map[商品名原題] = [];
      既存商品名Map[商品名原題].push(rowNum);
    }
  });

  // 重複チェック
  const ブロックリスト = [];
  対象行リスト.forEach(({ row, rowNum }) => {
    const 博客來コード = String(get行(row, '博客來商品コード') || '').trim();
    const 商品名原題   = String(get行(row, '商品名（原題）') || '').trim();

    // ① 博客來商品コード重複チェック
    if (博客來コード && 既存博客來Map[博客來コード]) {
      const 重複行 = 既存博客來Map[博客來コード].filter(r => r !== rowNum);
      if (重複行.length > 0) {
        ブロックリスト.push(`${rowNum}行目：博客來商品コード「${博客來コード}」が${重複行[0]}行目と重複`);
        return;
      }
    }

    // ② 商品名（原題）重複チェック
    if (商品名原題 && 既存商品名Map[商品名原題]) {
      const 重複行 = 既存商品名Map[商品名原題].filter(r => r !== rowNum);
      if (重複行.length > 0) {
        ブロックリスト.push(`${rowNum}行目：商品名（原題）「${商品名原題}」が${重複行[0]}行目と重複`);
      }
    }
  });

  if (ブロックリスト.length > 0) {
    ui.alert(`⚠️ 重複が見つかりました。以下を確認してください。\n\n${ブロックリスト.join('\n')}\n\n重複を解消してから再度実行してください。`);
    return;
  }

  // 発行処理
  const lock = LockService.getDocumentLock();
  lock.waitLock(60000);

  try {
    const Worksシート = 台湾グッズ_Worksシートを確保_(ss);
    const Works更新Map = {};
    let 発行数 = 0;

    対象行リスト.forEach(({ row, rowNum }) => {
      const 親コード = String(get行(row, '親コード') || '').trim();
      if (!親コード || 親コード.startsWith('ERROR')) {
        ui.alert(`${rowNum}行目：親コードが未生成のためスキップしました`);
        return;
      }

      // 発行
      sh.getRange(rowNum, 列['登録状況']).setValue('未登録');
      sh.getRange(rowNum, 列['発番発行']).setValue(false);
      発行数++;

      // Works更新Map積み上げ
      const 作品名 = String(get行(row, '作品名（原題）') || '').trim();
      if (作品名) {
        if (!Works更新Map[作品名]) Works更新Map[作品名] = {
          日本語: String(get行(row, '作品名（日本語）') || '').trim(),
          件数: 0
        };
        Works更新Map[作品名].件数++;
      }
    });

    台湾グッズ_Worksを更新_(Worksシート, Works更新Map);
    ui.alert(`✅ 確定発行完了: ${発行数}件\n\n発行された行の登録状況が「未登録」になりました。\nYahoo登録後に「登録済み」に変更してください。`);

  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
 * Works更新（内部）
 * ============================================================ */
function 台湾グッズ_Worksシートを確保_(ss) {
  let sh = ss.getSheetByName(設定_台湾グッズ.作品シート名);
  if (!sh) {
    sh = ss.insertSheet(設定_台湾グッズ.作品シート名);
    sh.getRange(1, 1, 1, 設定_台湾グッズ.作品列数).setValues([設定_台湾グッズ.作品ヘッダー]);
    sh.getRange(1, 1, 1, 設定_台湾グッズ.作品列数)
      .setBackground('#cc0000').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
    sh.setFrozenRows(1);
  }
  return sh;
}

function 台湾グッズ_Worksを更新_(Worksシート, Works更新Map) {
  const 既存Map = {};
  if (Worksシート.getLastRow() >= 2) {
    Worksシート.getRange(2, 1, Worksシート.getLastRow() - 1, 設定_台湾グッズ.作品列数).getValues()
      .forEach((row, i) => { if (row[2]) 既存Map[String(row[2]).trim()] = i + 2; });
  }

  Object.keys(Works更新Map).forEach(作品名 => {
    const info = Works更新Map[作品名];
    const WorksKey = 作品名.replace(/\s+/g, '').toLowerCase();
    if (既存Map[作品名]) {
      const 行 = 既存Map[作品名];
      Worksシート.getRange(行, 6).setValue(info.件数);
      Worksシート.getRange(行, 7).setValue(new Date());
    } else {
      const 新ID = 'TW-W-' + String(Worksシート.getLastRow()).padStart(4, '0');
      const マスター情報 = 台湾グッズ_作品マスターから取得_(作品名);
      Worksシート.appendRow([
        WorksKey, 新ID, 作品名, info.日本語,
        マスター情報 ? マスター情報.作品略称 : '',
        info.件数, new Date()
      ]);
      既存Map[作品名] = Worksシート.getLastRow();
    }
  });
}

/* ============================================================
 * プルダウン更新
 * ============================================================ */
function 台湾グッズ_プルダウン更新() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('台湾グッズ');
  if (!sh) { SpreadsheetApp.getUi().alert('台湾グッズシートが見つかりません'); return; }

  const 最終行 = Math.max(sh.getLastRow() + 100, 200);
  const ヘッダー = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const 列 = {};
  ヘッダー.forEach((h, i) => { if (h) 列[String(h).trim()] = i + 1; });

  let masterSS = null;
  try { masterSS = 台湾グッズ_マスターSSを取得_(); } catch(_) {}

  // 言語プルダウン（言語マスター参照）
  if (masterSS && 列['言語']) {
    const 言語シート = masterSS.getSheetByName('言語マスター');
    if (言語シート && 言語シート.getLastRow() >= 2) {
      const 言語値 = 言語シート.getRange(2, 1, 言語シート.getLastRow() - 1, 1)
        .getValues().flat().map(v => String(v).trim()).filter(v => v);
      if (言語値.length > 0) {
        sh.getRange(2, 列['言語'], 最終行 - 1, 1).setDataValidation(
          SpreadsheetApp.newDataValidation().requireValueInList(言語値, true).build()
        );
      }
    }
  }

  // 登録状況プルダウン
  if (列['登録状況']) {
    sh.getRange(2, 列['登録状況'], 最終行 - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['未登録', '登録済み', '売り切れ'], true).build()
    );
  }

  // 発番発行チェックボックス（念のため再設定）
  if (列['発番発行']) {
    sh.getRange(2, 列['発番発行'], 最終行 - 1, 1).insertCheckboxes();
  }

  // 配送パターンプルダウン
  if (masterSS && 列['配送パターン']) {
    const 配送マスター = masterSS.getSheetByName('配送パターンマスター');
    if (配送マスター && 配送マスター.getLastRow() >= 2) {
      const 配送値 = 配送マスター.getRange(2, 1, 配送マスター.getLastRow() - 1, 1)
        .getValues().flat().map(v => String(v).trim()).filter(v => v);
      if (配送値.length > 0) {
        sh.getRange(2, 列['配送パターン'], 最終行 - 1, 1).setDataValidation(
          SpreadsheetApp.newDataValidation().requireValueInList(配送値, true).setAllowInvalid(false).build()
        );
      }
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('✅ 台湾グッズ プルダウン更新完了', '完了', 3);
}