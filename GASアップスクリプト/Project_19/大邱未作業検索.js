// ============================================================
//  発注リスト大邱データ → 大邱未作業データ（色なし行の抽出＋多角度検索）
//
//  ・発注リスト大邱データで行に背景色（オレンジ/黄色=EMS積載済み）が
//    付いていない行だけを「大邱未作業データ」シートに色なしで書き出す
//  ・B1: 全列キーワード検索（スペース区切りAND・部分一致）
//  ・2行目（水色）: 列ごとの絞り込み（スペース区切りAND・部分一致）
//  ・検索セルを編集すると onEdit 経由で自動絞り込み
// ============================================================

const MISAGYO_CFG = {
  SRC_SHEET: '発注リスト大邱データ',
  DST_SHEET: '大邱未作業データ',

  SRC_HEADER_JP: 4,   // 日本語ヘッダー行
  SRC_HEADER_EN: 5,   // 英語ヘッダー行
  SRC_DATA_START: 6,  // データ開始行
  SRC_COL_START: 2,   // B列から
  SRC_COL_END: 19,    // S列（支払金額）まで
  COLOR_CHECK_COL: 2, // 色判定はB列（発注日）の背景色

  DST_KEYWORD_ROW: 1, // B1 = 全列キーワード
  DST_FILTER_ROW: 2,  // 列別絞り込み行
  DST_HEADER_JP: 3,
  DST_HEADER_EN: 4,
  DST_DATA_START: 5
};

// ---------- メニュー ----------
function 大邱未作業メニューを追加_() {
  SpreadsheetApp.getUi()
    .createMenu('大邱未作業')
    .addItem('🔄 未作業データを更新', '大邱未作業_更新')
    .addItem('🔍 絞り込みを実行', '大邱未作業_絞り込み')
    .addItem('🧹 検索条件をクリア', '大邱未作業_検索条件をクリア')
    .addToUi();
}

// ---------- メイン：未作業データを再構築 ----------
function 大邱未作業_更新() {
  const result = 大邱未作業_再構築_(true);
  if (result) {
    SpreadsheetApp.getActive().toast(
      '未作業 ' + result.total + '件を書き出しました。',
      '大邱未作業データ', 5
    );
  }
}

function 大邱未作業_絞り込み() {
  const result = 大邱未作業_再構築_(false);
  if (result) {
    SpreadsheetApp.getActive().toast(
      '該当 ' + result.shown + '件 / 未作業 ' + result.total + '件',
      '大邱未作業データ', 5
    );
  }
}

function 大邱未作業_検索条件をクリア() {
  const cfg = MISAGYO_CFG;
  const sh = SpreadsheetApp.getActive().getSheetByName(cfg.DST_SHEET);
  if (sh) {
    const width = cfg.SRC_COL_END - cfg.SRC_COL_START + 1;
    sh.getRange(cfg.DST_KEYWORD_ROW, 2).clearContent();          // B1
    sh.getRange(cfg.DST_FILTER_ROW, 1, 1, width).clearContent(); // 2行目
  }
  大邱未作業_更新();
}

// ---------- onEdit から呼ばれる（検索セル編集で自動絞り込み） ----------
function 大邱未作業_onEdit_(e) {
  const cfg = MISAGYO_CFG;
  const range = e.range;
  if (range.getRow() > cfg.DST_FILTER_ROW) return; // 1〜2行目だけ反応

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(1000)) return;
  try {
    大邱未作業_再構築_(false);
  } finally {
    lock.releaseLock();
  }
}

// ---------- 本体 ----------
// clearFilters=true: 検索条件を無視して全件表示
// 返り値: { total: 未作業件数, shown: 表示件数 }
function 大邱未作業_再構築_(clearFilters) {
  const cfg = MISAGYO_CFG;
  const ss = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(cfg.SRC_SHEET);
  if (!src) {
    SpreadsheetApp.getUi().alert('「' + cfg.SRC_SHEET + '」が見つかりません。');
    return null;
  }

  const width = cfg.SRC_COL_END - cfg.SRC_COL_START + 1;

  // --- ソース読み込み ---
  const lastRow = src.getLastRow();
  let values = [], display = [], backgrounds = [];
  if (lastRow >= cfg.SRC_DATA_START) {
    const numRows = lastRow - cfg.SRC_DATA_START + 1;
    const dataRange = src.getRange(cfg.SRC_DATA_START, cfg.SRC_COL_START, numRows, width);
    values = dataRange.getValues();
    display = dataRange.getDisplayValues();
    backgrounds = src
      .getRange(cfg.SRC_DATA_START, cfg.COLOR_CHECK_COL, numRows, 1)
      .getBackgrounds();
  }

  // --- 色なし行（未作業）だけ抽出 ---
  // F発注NO / I商品名 / K商品コード（幅内index: F=4, I=7, K=9）のどれかに値がある行だけ
  const requiredIdx = [4, 7, 9];
  const unworked = [];
  for (let i = 0; i < values.length; i++) {
    const bg = String(backgrounds[i][0] || '#ffffff').toLowerCase();
    if (bg !== '#ffffff' && bg !== 'white' && bg !== '') continue; // 色付き=EMS積載済み
    if (!requiredIdx.some(idx => String(display[i][idx] || '').trim() !== '')) continue;
    unworked.push({ values: values[i], display: display[i] });
  }

  // --- 出力先シート準備 ---
  let dst = ss.getSheetByName(cfg.DST_SHEET);
  if (!dst) {
    dst = ss.insertSheet(cfg.DST_SHEET);
  }

  // --- 検索条件の取得（clearFilters時は空扱い） ---
  let keyword = '';
  let colFilters = new Array(width).fill('');
  if (!clearFilters) {
    keyword = String(dst.getRange(cfg.DST_KEYWORD_ROW, 2).getDisplayValue() || '').trim();
    colFilters = dst
      .getRange(cfg.DST_FILTER_ROW, 1, 1, width)
      .getDisplayValues()[0]
      .map(v => String(v || '').trim());
  }

  const keywordTerms = keyword ? keyword.toLowerCase().split(/[\s　]+/).filter(Boolean) : [];
  const filterTerms = colFilters.map(f =>
    f ? f.toLowerCase().split(/[\s　]+/).filter(Boolean) : []
  );

  // --- 絞り込み ---
  const shownRows = unworked.filter(row => {
    const lowerCells = row.display.map(v => String(v || '').toLowerCase());

    // 全列キーワード: 各語がどこかの列に含まれる（AND）
    for (const term of keywordTerms) {
      if (!lowerCells.some(cell => cell.indexOf(term) >= 0)) return false;
    }
    // 列別: その列に全語が含まれる（AND）
    for (let c = 0; c < width; c++) {
      for (const term of filterTerms[c]) {
        if (lowerCells[c].indexOf(term) < 0) return false;
      }
    }
    return true;
  });

  // --- ヘッダー・検索エリアを整備（毎回上書きで壊れない） ---
  const jpHeader = src.getRange(cfg.SRC_HEADER_JP, cfg.SRC_COL_START, 1, width).getDisplayValues();
  const enHeader = src.getRange(cfg.SRC_HEADER_EN, cfg.SRC_COL_START, 1, width).getDisplayValues();

  dst.getRange(1, 1)
    .setValue('🔍 全列検索 →')
    .setFontWeight('bold');
  dst.getRange(cfg.DST_KEYWORD_ROW, 2).setBackground('#fff2cc'); // B1 黄色
  dst.getRange(1, 3).setValue(
    '←スペース区切りAND・部分一致／下の水色行は列ごとの条件。セルを編集すると自動で絞り込み。'
  ).setFontColor('#888888').setFontSize(9);

  dst.getRange(cfg.DST_FILTER_ROW, 1, 1, width).setBackground('#e8f0fe'); // 水色
  dst.getRange(cfg.DST_HEADER_JP, 1, 1, width)
    .setValues(jpHeader)
    .setFontWeight('bold')
    .setBackground('#efefef');
  dst.getRange(cfg.DST_HEADER_EN, 1, 1, width)
    .setValues(enHeader)
    .setFontWeight('bold')
    .setBackground('#efefef');
  dst.setFrozenRows(cfg.DST_HEADER_EN);

  // 列幅をソースに合わせる
  for (let c = 0; c < width; c++) {
    dst.setColumnWidth(c + 1, src.getColumnWidth(cfg.SRC_COL_START + c));
  }

  // --- データ書き込み（値のみ＝色なし） ---
  const maxRows = dst.getMaxRows();
  if (maxRows >= cfg.DST_DATA_START) {
    dst.getRange(cfg.DST_DATA_START, 1, maxRows - cfg.DST_DATA_START + 1, dst.getMaxColumns())
      .clearContent();
  }
  if (shownRows.length) {
    dst.getRange(cfg.DST_DATA_START, 1, shownRows.length, width)
      .setValues(shownRows.map(r => r.values));
  }

  // --- 件数表示 ---
  dst.getRange(1, width + 1)
    .setValue('該当 ' + shownRows.length + ' / 未作業 ' + unworked.length + '件')
    .setFontColor('#666666')
    .setFontWeight('bold');

  SpreadsheetApp.flush();
  return { total: unworked.length, shown: shownRows.length };
}
