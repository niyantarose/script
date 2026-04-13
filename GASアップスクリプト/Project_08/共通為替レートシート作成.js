/**
 * 為替レートマスターのD〜E列にエイリアス検索テーブルを自動生成する
 * 
 * A列の通貨コード（TWD, CN, KR, HK, TH, USD）を読み取り、
 * 各通貨に対応するエイリアス（日本語名、ISOコード等）をD-E列に展開する。
 * E列は =B{行番号} の数式で参照するため、GOOGLEFINANCEレート更新に自動追従。
 *
 * 使い方: メニュー or スクリプトエディタから 為替エイリアステーブルを生成() を実行
 */

function 為替エイリアステーブルを生成() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('為替レートマスター');
  if (!sh) { SpreadsheetApp.getUi().alert('「為替レートマスター」シートが見つかりません'); return; }

  // ── A列の通貨コード → 行番号マッピングを読み取り ──
  const lastRow = sh.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('通貨データがありません'); return; }
  const currencies = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  const codeToRow = {};
  for (let i = 0; i < currencies.length; i++) {
    const code = String(currencies[i][0] || '').trim().toUpperCase();
    if (code) codeToRow[code] = i + 2; // シート行番号
  }

  // ── エイリアス定義（通貨コード → 検索キー一覧） ──
  // A列のコードに対応するエイリアスを定義。A列のコード自体も含む。
  const ALIAS_MAP = {
    'TWD': ['TWD', '台湾', 'TW', '台灣'],
    'CN':  ['CN', '中国', 'CNY', '中國'],
    'KR':  ['KR', '韓国', 'KRW', '한국'],
    'HK':  ['HK', '香港', 'HKD'],
    'TH':  ['TH', 'タイ', 'THB'],
    'USD': ['USD', 'アメリカ', '英語', 'US', 'EN'],
  };

  // ── D-E列データを構築 ──
  const rows = [['検索キー', 'レート']]; // ヘッダー

  for (const [code, aliases] of Object.entries(ALIAS_MAP)) {
    const sheetRow = codeToRow[code];
    if (!sheetRow) {
      Logger.log('⚠ A列に「' + code + '」が見つかりません。スキップします。');
      continue;
    }
    for (const alias of aliases) {
      rows.push([alias, '=B' + sheetRow]);
    }
  }

  // 日本円（固定レート1）を追加
  const jpAliases = ['JP', '日本', '日本語', 'JPY'];
  for (const alias of jpAliases) {
    rows.push([alias, 1]);
  }

  // ── D-E列に書き込み（既存データをクリアしてから） ──
  const writeStartRow = 1;
  const existingLastRow = sh.getRange('D:D').getValues().filter(r => r[0] !== '').length;
  if (existingLastRow > 0) {
    sh.getRange(1, 4, Math.max(existingLastRow, rows.length) + 1, 2).clearContent();
  }

  sh.getRange(writeStartRow, 4, rows.length, 2).setValues(rows);

  // ── 書式整え ──
  // ヘッダー太字
  sh.getRange(1, 4, 1, 2).setFontWeight('bold');
  // D列幅調整
  sh.setColumnWidth(4, 100);
  sh.setColumnWidth(5, 100);

  SpreadsheetApp.getUi().alert(
    `✅ エイリアステーブル生成完了\n` +
    `D-E列に ${rows.length - 1} 件のエイリアスを書き込みました。\n\n` +
    `粗利計算数式例:\n` +
    `=VLOOKUP(P2,'為替レートマスター'!$D:$E,2,FALSE)`
  );
}

/**
 * メニューに追加（onOpen から呼ぶか、単独で使う）
 */
function 為替メニューを追加_() {
  SpreadsheetApp.getUi().createMenu('為替ツール')
    .addItem('エイリアステーブル生成/更新', '為替エイリアステーブルを生成')
    .addToUi();
}