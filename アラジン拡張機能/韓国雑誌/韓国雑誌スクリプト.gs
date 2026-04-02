/**
 * 韓国雑誌.gs
 * 韓国雑誌シート専用の設定・シート作成・onEdit・確定発行
 *
 * ★ コード体系
 *   年月型: {略称}{年下2桁}{月2桁}      例: MAXI2603
 *   号数型: {略称}{号数}                 例: 1STLOOK127
 *
 * 【確定発行の仕組み】
 * 発番発行列（チェックボックス）をONにした行だけが対象
 * 重複チェック：雑誌名＋年＋月＋号数の組み合わせ
 * 問題なければ発行 → 登録状況=未登録、チェックOFF
 */

const 設定_韓国雑誌 = {
  マスターシート名: '韓国雑誌',
  雑誌マスター名:   '雑誌マスター（韓国）',
　候補シート名:     '雑誌マスター候補（共通）',  // ← 追加

  列名: {
    発番発行:           '発番発行',
    登録状況:           '登録状況',
    雑誌名:             '雑誌名',
    年:                 '年',
    月:                 '月',
    号数:               '号数',
    表紙情報:           '表紙情報',
    特典メモ:           '特典メモ',
    親コード:           '親コード',
    商品名出品用:       '商品名（出品用）',
    粗利益率:           '粗利益率',
    登録日:             '登録日',
    売価:               '売価',
    配送パターン:       '配送パターン',
    登録者:             '登録者',
    商品説明:           '商品説明',
    原価:               '原価',
    原題タイトル:       '原題タイトル',
    原題商品名:         '原題商品名',
    アラジン商品コード: 'アラジン商品コード',
    アラジンURL:        'アラジンURL',
    メイン画像URL:      'メイン画像URL',
    追加画像URL:        '追加画像URL'
  }
};

/* ============================================================
 * 雑誌マスター（韓国）作成
 * 列: 雑誌名（英字）| 雑誌名（カタカナ）| 略称コード | コード型 | 備考
 * ============================================================ */
function 韓国雑誌マスターを作成() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const シート名 = 設定_韓国雑誌.雑誌マスター名;
  if (ss.getSheetByName(シート名)) { ui.alert(`「${シート名}」シートは既に存在します`); return; }

  const sh = ss.insertSheet(シート名);
  const ヘッダー = ['雑誌名（英字）', '雑誌名（カタカナ）', '略称コード', 'コード型', '備考'];
  sh.getRange(1, 1, 1, ヘッダー.length).setValues([ヘッダー]);

  const 初期値 = [
    ['MAXIM KOREA',  'マキシム・コリア', 'MAXI',     '年月型', ''],
    ['GQ Korea',     'ジーキュー',       'GQ',       '年月型', ''],
    ['BAZAAR KOREA', 'バザー',           'BAZA',     '年月型', ''],
    ['1ST LOOK',     'ファーストルック', '1STLOOK',  '号数型', ''],
    ['CINE21',       'シネ21',           'CIN21',    '号数型', ''],
  ];
  sh.getRange(2, 1, 初期値.length, ヘッダー.length).setValues(初期値);

  sh.getRange(2, 4, 100, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['年月型', '号数型'], true).build()
  );

  sh.getRange(1, 1, 1, ヘッダー.length)
    .setBackground('#4a86e8').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  [200, 160, 120, 80, 200].forEach((w, i) => sh.setColumnWidth(i + 1, w));

  ui.alert('✅ 雑誌マスター（韓国）を作成しました\n\n' +
    '【コード型の意味】\n' +
    '年月型: {略称}{年下2桁}{月2桁}  例: MAXI2603\n' +
    '号数型: {略称}{号数}             例: 1STLOOK127\n\n' +
    '※ コードや商品名は自動生成後に手動で修正できます\n' +
    '追加後は「プルダウン更新」を実行してください');
}

/* ============================================================
 * 韓国雑誌シート作成
 * ============================================================ */
function 韓国雑誌シートを作成() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const シート名 = '韓国雑誌';
  if (ss.getSheetByName(シート名)) { ui.alert(`「${シート名}」シートは既に存在します`); return; }

  const sh = ss.insertSheet(シート名);

  const ヘッダー = [
    '発番発行', '登録状況', '雑誌名', '年', '月', '号数', '表紙情報', '特典メモ',
    '親コード', '商品名（出品用）', '粗利益率', '登録日',
    '売価', '配送パターン', '登録者', '商品説明',
    '原価', '原題タイトル', '原題商品名', 'アラジン商品コード', 'アラジンURL',
    'メイン画像URL', '追加画像URL'
  ];

  sh.getRange(1, 1, 1, ヘッダー.length).setValues([ヘッダー]);

  const 韓国雑誌_ヘッダー色 = {
    '発番発行':           '#cc0000', // 赤（発行操作列）
    '登録状況':           '#4a86e8', // 青
    '雑誌名':             '#6aa84f', // 緑
    '年':                 '#6aa84f', // 緑
    '月':                 '#6aa84f', // 緑
    '号数':               '#6aa84f', // 緑
    '表紙情報':           '#f1c232', // 黄
    '特典メモ':           '#f1c232', // 黄
    '親コード':           '#999999', // 灰
    '商品名（出品用）':   '#999999', // 灰
    '粗利益率':           '#999999', // 灰
    '登録日':             '#999999', // 灰
    '売価':               '#4a86e8', // 青
    '配送パターン':       '#4a86e8', // 青
    '登録者':             '#4a86e8', // 青
    '商品説明':           '#4a86e8', // 青
    '原価':               '#e69138', // 橙
    '原題タイトル':       '#e69138', // 橙
    '原題商品名':         '#e69138', // 橙
    'アラジン商品コード': '#e69138', // 橙
    'アラジンURL':        '#e69138', // 橙
    'メイン画像URL':      '#e69138', // 橙
    '追加画像URL':        '#e69138', // 橙
  };

  for (let i = 0; i < ヘッダー.length; i++) {
    const 色 = 韓国雑誌_ヘッダー色[ヘッダー[i]] || '#cccccc';
    sh.getRange(1, i + 1).setBackground(色).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  }

  sh.setRowHeight(1, 40);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);

  const 最終行 = 1000;

  // 発番発行チェックボックス（A列）
  sh.getRange(2, 1, 最終行 - 1, 1).insertCheckboxes();

  // 登録状況プルダウン（B列）
  sh.getRange(2, 2, 最終行 - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['未登録', '登録済み', '売り切れ'], true).build()
  );

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

  const 列幅マップ = {
    '発番発行': 60, '登録状況': 80, '雑誌名': 180, '年': 60, '月': 50, '号数': 70,
    '表紙情報': 200, '特典メモ': 240,
    '親コード': 140, '商品名（出品用）': 360, '粗利益率': 90, '登録日': 120,
    '売価': 80, '配送パターン': 100, '登録者': 80, '商品説明': 150,
    '原価': 80, '原題タイトル': 200, '原題商品名': 200,
    'アラジン商品コード': 140, 'アラジンURL': 220, 'メイン画像URL': 200, '追加画像URL': 200
  };
  ヘッダー.forEach((h, i) => { if (列幅マップ[h]) sh.setColumnWidth(i + 1, 列幅マップ[h]); });

  ui.alert('✅ 韓国雑誌シートを作成しました\n\n【発行の手順】\n① 行を入力する\n② 発番発行列にチェックを入れる\n③ 確定発行ボタンを押す\n\n【号数列について】\n・年月型: 年・月を入力（号数は空欄でOK）\n・号数型: 号数を入力（年・月は任意）\n\n※ 親コード・商品名（出品用）は自動生成後に手動修正できます\n\n続けて「プルダウン更新」を実行してください');
}

/* ============================================================
 * 内部ヘルパー
 * ============================================================ */
function 韓国雑誌_マスターを検索_(雑誌名英字) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(設定_韓国雑誌.雑誌マスター名);
  if (!sh || sh.getLastRow() < 2) return null;
  const データ = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  for (const row of データ) {
    if (String(row[0]).trim() === String(雑誌名英字).trim()) {
      return {
        英字名:     String(row[0]).trim(),
        カタカナ名: String(row[1]).trim(),
        略称:       String(row[2]).trim(),
        コード型:   String(row[3]).trim() || '年月型'
      };
    }
  }
  return null;
}

function 韓国雑誌_親コードを生成_(雑誌名, 年, 月, 号数) {
  const info = 韓国雑誌_マスターを検索_(雑誌名);
  if (!info || !info.略称) return 'ERROR:マスター未登録';
  if (info.コード型 === '号数型') {
    if (!号数) return 'ERROR:号数未入力';
    return `${info.略称}${String(号数).trim()}`;
  }
  if (!年 || !月) return 'ERROR:年月未入力';
  return `${info.略称}${String(年).slice(-2)}${String(月).padStart(2, '0')}`;
}

function 韓国雑誌_出品用商品名を生成_(雑誌名, 年, 月, 号数, 表紙情報, 特典メモ) {
  if (!雑誌名) return '';
  const info = 韓国雑誌_マスターを検索_(雑誌名);
  const 英字 = info ? info.英字名 : 雑誌名;
  const カタカナ = info ? info.カタカナ名 : '';
  const コード型 = info ? info.コード型 : '年月型';

  let 名前 = `韓国 雑誌 ${英字}`;
  if (カタカナ) 名前 += ` (${カタカナ})`;

  if (コード型 === '号数型') {
    if (!号数) return '';
    名前 += ` Vol.${String(号数).trim()}`;
    if (年 && 月) 名前 += ` (${年}年${月}月)`;
  } else {
    if (!年 || !月) return '';
    名前 += ` ${年}年 ${月}月号`;
  }

  if (表紙情報 && String(表紙情報).trim()) 名前 += ` (${String(表紙情報).trim()})`;
  if (特典メモ  && String(特典メモ).trim())  名前 += ` ${String(特典メモ).trim()}`;
  return 名前;
}

/* ============================================================
 * onEdit
 * ============================================================ */
function 韓国雑誌_onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== '韓国雑誌') return;
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

      const 監視列 = ['雑誌名', '年', '月', '号数', '表紙情報', '特典メモ'];
      const 編集列名 = Object.keys(列).find(h => 列[h] === col);
      if (!監視列.includes(編集列名)) return;

      const 雑誌名   = get('雑誌名');
      const 年       = get('年');
      const 月       = get('月');
      const 号数     = get('号数');
      const 表紙情報 = get('表紙情報');
      const 特典メモ = get('特典メモ');

      if (雑誌名) {
        const 親コード = 韓国雑誌_親コードを生成_(雑誌名, 年, 月, 号数);
        if (!親コード.startsWith('ERROR')) set('親コード', 親コード);
        const 商品名 = 韓国雑誌_出品用商品名を生成_(雑誌名, 年, 月, 号数, 表紙情報, 特典メモ);
        if (商品名) set('商品名（出品用）', 商品名);
      }

      if (雑誌名 && !get('登録日')) set('登録日', new Date());

      // 雑誌名を編集した時だけ候補チェック
      if (編集列名 === '雑誌名') {
        韓国雑誌_未解決候補を自動追加_(sh, row);
      }

    } finally {
      自己更新を終了_();
    }
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
 * 確定発行
 * 重複チェック: 雑誌名＋年＋月＋号数の組み合わせ
 * ============================================================ */
function 韓国雑誌_確定発行() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('韓国雑誌');
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

  // 重複チェック用Map（全データ）
  const 重複チェックMap = {};
  全データ.forEach((row, i) => {
    const rowNum = i + 2;
    const キー = 韓国雑誌_重複キーを作成_(
      String(get行(row, '雑誌名') || '').trim(),
      String(get行(row, '年') || '').trim(),
      String(get行(row, '月') || '').trim(),
      String(get行(row, '号数') || '').trim()
    );
    if (!キー) return;
    if (!重複チェックMap[キー]) 重複チェックMap[キー] = [];
    重複チェックMap[キー].push(rowNum);
  });

  // 重複チェック
  const ブロックリスト = [];
  対象行リスト.forEach(({ row, rowNum }) => {
    const キー = 韓国雑誌_重複キーを作成_(
      String(get行(row, '雑誌名') || '').trim(),
      String(get行(row, '年') || '').trim(),
      String(get行(row, '月') || '').trim(),
      String(get行(row, '号数') || '').trim()
    );
    if (!キー) return;
    const 重複行 = (重複チェックMap[キー] || []).filter(r => r !== rowNum);
    if (重複行.length > 0) {
      ブロックリスト.push(`${rowNum}行目：「${キー}」が${重複行[0]}行目と重複`);
    }
  });

  if (ブロックリスト.length > 0) {
    ui.alert(`⚠️ 重複が見つかりました。以下を確認してください。\n\n${ブロックリスト.join('\n')}\n\n重複を解消してから再度実行してください。`);
    return;
  }
// 親コード完全一致チェック
  const 親コード重複Map = {};
  全データ.forEach((row, i) => {
    const code = String(get行(row, '親コード') || '').trim();
    if (!code || code.startsWith('ERROR')) return;
    if (!親コード重複Map[code]) 親コード重複Map[code] = [];
    親コード重複Map[code].push(i + 2);
  });

  const 親コードブロックリスト = [];
  対象行リスト.forEach(({ row, rowNum }) => {
    const code = String(get行(row, '親コード') || '').trim();
    if (!code || code.startsWith('ERROR')) return;
    const 重複行 = (親コード重複Map[code] || []).filter(r => r !== rowNum);
    if (重複行.length > 0) {
      親コードブロックリスト.push(`${rowNum}行目：親コード「${code}」が${重複行[0]}行目と重複`);
    }
  });

  if (親コードブロックリスト.length > 0) {
    ui.alert(`⚠️ 親コードの重複が見つかりました。\n\n${親コードブロックリスト.join('\n')}\n\n重複を解消してから再度実行してください。`);
    return;
  }
  // 発行処理
  const lock = LockService.getDocumentLock();
  lock.waitLock(60000);
  try {
    let 発行数 = 0;
    対象行リスト.forEach(({ row, rowNum }) => {
      const 親コード = String(get行(row, '親コード') || '').trim();
      if (!親コード || 親コード.startsWith('ERROR')) {
        ui.alert(`${rowNum}行目：親コードが未生成のためスキップしました\n手動で親コードを入力してください`);
        return;
      }
      sh.getRange(rowNum, 列['登録状況']).setValue('未登録');
      sh.getRange(rowNum, 列['発番発行']).setValue(false);
      発行数++;
    });
    ui.alert(`✅ 確定発行完了: ${発行数}件\n\nYahoo登録後に登録状況を「登録済み」に変更してください。`);
  } finally {
    lock.releaseLock();
  }
}

function 韓国雑誌_重複キーを作成_(雑誌名, 年, 月, 号数) {
  if (!雑誌名) return '';
  if (号数) return `${雑誌名}-${号数}`;
  if (年 && 月) return `${雑誌名}-${年}-${月}`;
  return '';
}

/* ============================================================
 * プルダウン更新
 * ============================================================ */
function 韓国雑誌_プルダウン更新() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('韓国雑誌');
  if (!sh) { SpreadsheetApp.getUi().alert('韓国雑誌シートが見つかりません'); return; }

  const 最終行 = Math.max(sh.getLastRow() + 100, 200);
  const ヘッダー = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const 列 = {};
  ヘッダー.forEach((h, i) => { if (h) 列[String(h).trim()] = i + 1; });

  // 雑誌名プルダウン
  const 雑誌マスター = ss.getSheetByName(設定_韓国雑誌.雑誌マスター名);
  if (雑誌マスター && 雑誌マスター.getLastRow() >= 2 && 列['雑誌名']) {
    const 雑誌値 = 雑誌マスター.getRange(2, 1, 雑誌マスター.getLastRow() - 1, 1)
      .getValues().flat().map(v => String(v).trim()).filter(v => v);
    if (雑誌値.length > 0) {
      sh.getRange(2, 列['雑誌名'], 最終行 - 1, 1).setDataValidation(
        SpreadsheetApp.newDataValidation().requireValueInList(雑誌値, true).build()
      );
    }
  }

  // 登録状況
  if (列['登録状況']) {
    sh.getRange(2, 列['登録状況'], 最終行 - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['未登録', '登録済み', '売り切れ'], true).build()
    );
  }

  // 発番発行チェックボックス（念のため再設定）
  if (列['発番発行']) {
    sh.getRange(2, 列['発番発行'], 最終行 - 1, 1).insertCheckboxes();
  }

  // 配送パターン
  let masterSS = null;
  try { masterSS = SpreadsheetApp.openById('1ZAy5wtQCq1ixl47MMEIGnrVWeae-F_NfN34lSOUsN5M'); } catch(_) {}
  if (masterSS && 列['配送パターン']) {
    const 配送マスター = masterSS.getSheetByName('配送パターンマスター');
    if (配送マスター && 配送マスター.getLastRow() >= 2) {
      const 配送値 = 配送マスター.getRange(2, 1, 配送マスター.getLastRow() - 1, 1)
        .getValues().flat().map(v => String(v).trim()).filter(v => v);
      if (配送値.length > 0) {
        sh.getRange(2, 列['配送パターン'], 最終行 - 1, 1).setDataValidation(
          SpreadsheetApp.newDataValidation().requireValueInList(配送値, true).build()
        );
      }
    }
  }

  // 雑誌マスターのコード型プルダウンも更新
  if (雑誌マスター && 雑誌マスター.getLastRow() >= 2) {
    雑誌マスター.getRange(2, 4, 雑誌マスター.getLastRow() - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['年月型', '号数型'], true).build()
    );
  }

// 親コード重複を赤表示
  if (列['親コード']) {
    const 親コード列 = 列['親コード'];
    const colLetter = String.fromCharCode(64 + 親コード列);
    const rules = sh.getConditionalFormatRules();
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND($${colLetter}2<>"",COUNTIF($${colLetter}:$${colLetter},$${colLetter}2)>1)`)
        .setBackground('#f4cccc')
        .setFontColor('#cc0000')
        .setRanges([sh.getRange(2, 1, 最終行 - 1, sh.getLastColumn())])
        .build()
    );
    sh.setConditionalFormatRules(rules);
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('✅ 韓国雑誌 プルダウン更新完了', '完了', 3);
}

/* ============================================================
 * 韓国雑誌 未解決候補の自動収集・正式反映
 *
 * 【台湾雑誌との違い】
 *   - マスターはローカル（雑誌マスター（韓国）シート）
 *   - 言語接頭辞なし・原題解析なし
 *   - 候補シートは共通マスターファイルに集約（台湾と同じシート）
 * ============================================================ */

let _韓国雑誌_共通SSキャッシュ = null;

function 韓国雑誌_共通SS_() {
  if (_韓国雑誌_共通SSキャッシュ) return _韓国雑誌_共通SSキャッシュ;
  _韓国雑誌_共通SSキャッシュ = SpreadsheetApp.openById('1ZAy5wtQCq1ixl47MMEIGnrVWeae-F_NfN34lSOUsN5M');
  return _韓国雑誌_共通SSキャッシュ;
}

/**
 * 韓国雑誌マスターに完全一致するか（英字名のみ照合）
 */
function 韓国雑誌_マスターに完全一致するか_(候補名) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(設定_韓国雑誌.雑誌マスター名);
  if (!sh || sh.getLastRow() < 2 || !候補名) return false;
  const target = String(候補名).trim().toUpperCase();
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat();
  return values.some(v => String(v).trim().toUpperCase() === target);
}

/**
 * 候補シートに既に同じ正規化キーがあるか確認
 */
function 韓国雑誌_候補シートに既にあるか_(candidateSh, 正規化キー) {
  if (!candidateSh || candidateSh.getLastRow() < 2 || !正規化キー) return false;
  const values = candidateSh
    .getRange(2, 候補列.正規化キー, candidateSh.getLastRow() - 1, 1)
    .getDisplayValues().flat();
  return values.some(v => String(v).trim() === 正規化キー);
}

/**
 * 候補シートの既存行番号を返す（正規化キーで検索、なければ-1）
 */
function 韓国雑誌_候補シートの既存行番号_(candidateSh, 正規化キー) {
  if (!candidateSh || candidateSh.getLastRow() < 2 || !正規化キー) return -1;
  const values = candidateSh
    .getRange(2, 候補列.正規化キー, candidateSh.getLastRow() - 1, 1)
    .getDisplayValues().flat();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i]).trim() === 正規化キー) return i + 2;
  }
  return -1;
}

/**
 * 未解決行を候補シートへ自動追加する
 * - 親コードが「ERROR:マスター未登録」の行だけ対象
 * - 韓国マスター＋候補シートの両方で重複チェック
 */
function 韓国雑誌_未解決候補を自動追加_(sh, row) {
  if (!sh || sh.getName() !== '韓国雑誌' || row < 2) return false;

  const ヘッダー = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const 列 = {};
  ヘッダー.forEach((h, i) => { if (h) 列[String(h).trim()] = i + 1; });
  const get = (名前) => 列[名前]
    ? String(sh.getRange(row, 列[名前]).getValue() || '').trim() : '';

  const 親コード = get('親コード');
  if (!String(親コード).startsWith('ERROR:マスター未登録')) return false;

  const 雑誌名   = get('雑誌名');
  const 表紙情報 = get('表紙情報');
  const 特典メモ = get('特典メモ');

  const 候補英字名 = 雑誌名.trim();
  if (!候補英字名) return false;

  const 正規化キー = 候補英字名.toUpperCase();

  // 韓国マスターに一致 → 追加しない
  if (韓国雑誌_マスターに完全一致するか_(候補英字名)) return false;

  const masterSS    = 韓国雑誌_共通SS_();
  const candidateSh = 台湾雑誌_候補シートを確保_(masterSS); // 共通の候補シートを使用

  // 既に候補シートにある → 最終検出日時だけ更新
  const 既存行 = 韓国雑誌_候補シートの既存行番号_(candidateSh, 正規化キー);
  if (既存行 >= 2) {
    candidateSh.getRange(既存行, 候補列.最終検出).setValue(new Date());
    return false;
  }

  const 備考 = [
    表紙情報 ? `表紙:${表紙情報}` : '',
    特典メモ  ? `特典:${特典メモ}`  : ''
  ].filter(Boolean).join(' / ');

  const now = new Date();
  const writeRow = Math.max(candidateSh.getLastRow() + 1, 2);

  candidateSh.getRange(writeRow, 1, 1, 候補列数).setValues([[
    候補ステータス.未対応,      // A: ステータス
    false,                      // B: 反映
    '',                         // C: 略称コード候補
    候補英字名,                 // D: 雑誌名（英字）
    '',                         // E: カタカナ
    '年月型',                   // F: 基本キー型（デフォルト）
    '',                         // G: 通常タイプ結合記号
    '-',                        // H: 版種ありタイプ結合記号
    'KR',                       // I: 対応言語
    '', '', '',                 // J〜L: 別名1〜3
    備考,                       // M: 備考
    sh.getParent().getName(),   // N: 元ファイル
    sh.getName(),               // O: 元シート
    row,                        // P: 元行
    now,                        // Q: 初回検出日時
    now,                        // R: 最終検出日時
    正規化キー                  // S: 正規化キー（隠し列）
  ]]);

  candidateSh.getRange(writeRow, 候補列.反映).insertCheckboxes();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `未解決雑誌を候補シートへ追加: ${候補英字名}`,
    '📋 雑誌マスター候補',
    4
  );

  return true;
}

/**
 * 候補シートから韓国雑誌マスターへ反映する
 */
function 韓国雑誌_候補を正式マスターへ反映() {
  const ui = SpreadsheetApp.getUi();
  const masterSS    = 韓国雑誌_共通SS_();
  const candidateSh = masterSS.getSheetByName(設定_韓国雑誌.候補シート名);

  if (!candidateSh || candidateSh.getLastRow() < 2) {
    ui.alert('候補シートにデータがありません'); return;
  }

  const ss = SpreadsheetApp.getActive();
  const 韓国マスター = ss.getSheetByName(設定_韓国雑誌.雑誌マスター名);
  if (!韓国マスター) {
    ui.alert('雑誌マスター（韓国）シートが見つかりません'); return;
  }

  const lastRow = candidateSh.getLastRow();
  const 全データ = candidateSh.getRange(2, 1, lastRow - 1, 候補列数).getValues();

  const 対象行リスト = [];
  全データ.forEach((row, i) => {
    if (row[候補列.反映 - 1] === true) {
      対象行リスト.push({ data: row, rowNum: i + 2 });
    }
  });

  if (対象行リスト.length === 0) {
    ui.alert('「反映」にチェックが入っている行がありません\n\nB列にチェックを入れてから実行してください');
    return;
  }

  // 言語が KR の行だけ対象にする
  const KR対象 = 対象行リスト.filter(({ data }) => {
    const 言語 = String(data[候補列.言語 - 1] || '').trim().toUpperCase();
    return 言語 === 'KR' || 言語 === '韓国' || 言語 === '';
  });

  if (KR対象.length === 0) {
    ui.alert('対応言語が KR の行がありません\n\n韓国雑誌以外の行にはチェックを入れないでください');
    return;
  }

  const res = ui.alert(
    '確認',
    `${KR対象.length}件を韓国雑誌マスターへ反映します。続行しますか？`,
    ui.ButtonSet.OK_CANCEL
  );
  if (res !== ui.Button.OK) return;

  let 反映数 = 0;
  const スキップリスト = [];
  const 反映済み行番号 = [];

  KR対象.forEach(({ data, rowNum }) => {
    const 英字名   = String(data[候補列.英字名   - 1] || '').trim();
    const 略称     = String(data[候補列.略称コード - 1] || '').trim();
    const カタカナ = String(data[候補列.カタカナ名 - 1] || '').trim();
    const 基本キー型 = String(data[候補列.基本キー型 - 1] || '').trim() || '年月型';
    const コード型 = 基本キー型 === '号数型' ? '号数型' : '年月型';

    if (!英字名) { スキップリスト.push(`${rowNum}行目: 雑誌名（英字）が空`); return; }
    if (!略称)   { スキップリスト.push(`${rowNum}行目 [${英字名}]: 略称コードが空`); return; }
    if (韓国雑誌_マスターに完全一致するか_(英字名)) {
      スキップリスト.push(`${rowNum}行目 [${英字名}]: 既にマスターに登録済み`);
      return;
    }

    // 韓国マスターの列順: 雑誌名（英字）/ 雑誌名（カタカナ）/ 略称コード / コード型 / 備考
    const writeRow = Math.max(韓国マスター.getLastRow() + 1, 2);
    韓国マスター.getRange(writeRow, 1, 1, 5).setValues([[
      英字名, カタカナ, 略称, コード型, ''
    ]]);

    反映済み行番号.push(rowNum);
    反映数++;
  });

  // ステータス更新・チェックOFF
  反映済み行番号.forEach(rowNum => {
    candidateSh.getRange(rowNum, 候補列.ステータス).setValue(候補ステータス.登録済み);
    candidateSh.getRange(rowNum, 候補列.反映).setValue(false);
  });

  let msg = `✅ 反映完了: ${反映数}件`;
  if (スキップリスト.length > 0) {
    msg += `\n\n⚠️ スキップ: ${スキップリスト.length}件\n${スキップリスト.join('\n')}`;
  }
  if (反映数 > 0) {
    msg += '\n\n「プルダウン更新」を実行してください';
  }
  ui.alert(msg);
}

/**
 * 候補シートの韓国雑誌分の件数を確認
 */
function 韓国雑誌_候補件数を確認() {
  const masterSS = 韓国雑誌_共通SS_();
  const sh = masterSS.getSheetByName(設定_韓国雑誌.候補シート名);

  if (!sh || sh.getLastRow() < 2) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '候補シートにデータがありません', '📋 雑誌マスター候補', 3
    );
    return;
  }

  // 言語=KRの行だけカウント
  const 全行 = sh.getRange(2, 1, sh.getLastRow() - 1, 候補列数).getValues();
  const KR行 = 全行.filter(row => {
    const 言語 = String(row[候補列.言語 - 1] || '').trim().toUpperCase();
    return 言語 === 'KR' || 言語 === '';
  });

  const カウント = {};
  Object.values(候補ステータス).forEach(s => { カウント[s] = 0; });
  KR行.forEach(row => {
    const s = String(row[候補列.ステータス - 1] || '');
    if (カウント[s] !== undefined) カウント[s]++;
  });

  const msg = [
    `【韓国雑誌候補】`,
    `未対応: ${カウント[候補ステータス.未対応]}件`,
    `確認中: ${カウント[候補ステータス.確認中]}件`,
    `登録待ち: ${カウント[候補ステータス.登録待ち]}件`,
    `登録済み: ${カウント[候補ステータス.登録済み]}件`,
    `無視: ${カウント[候補ステータス.無視]}件`,
    `\n共通マスターの「${設定_韓国雑誌.候補シート名}」を確認してください`
  ].join('\n');

  SpreadsheetApp.getActiveSpreadsheet().toast(msg, '📋 雑誌マスター候補', 8);
}