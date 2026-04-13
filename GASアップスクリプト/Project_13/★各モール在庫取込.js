/******************************************************
 * 在庫取込（同期実行・詳細ログ版）
 *
 * - Yahoo：足りないときだけ appendDimension で行・列を「増やすだけ」
 * - Amazon / Qoo10：列数はそのまま、行だけ足りなければ増やす
 *
 * ✅ トリガーを使わず、すべて同期実行
 * ✅ 詳細ログを出力
 *
 * ✅ 前提
 * - サービス > Sheets API を ON
 * - サービス > Drive API を ON（Advanced Google Services）
 ******************************************************/

/**********************
 * 共通設定
 **********************/
const フォルダID_Yahoo在庫  = '1_QtHRIvPb-ZYcUNnFveGu9jd9TTOxGpY';
const フォルダID_Amazon在庫 = '1kzWkaIpPy-Npojgpdnl36cwiDfViquju';
const フォルダID_Qoo10在庫  = '1Sqtjh9zYqgpve-SHsc2IUKMyWaBcFV1a';

// 書き込みを分割する行数（Values API）
const WRITE_CHUNK_ROWS = 30000;

// ★ このスクリプトが貼ってあるスプレッドシートID
function getSsId_() {
  return SpreadsheetApp.getActiveSpreadsheet().getId();
}

/**********************
 * メニュー
 **********************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('在庫取込(同期版)')
    .addItem('最新在庫をすべて取込', '最新在庫をすべて取込')
    .addSeparator()
    .addItem('Yahoo在庫だけ取込', '最新Yahoo在庫を取込')
    .addItem('Amazon在庫だけ取込', '最新Amazon在庫を取込')
    .addItem('Qoo10在庫だけ取込', '最新Qoo10在庫を取込')
    .addToUi();
}

/**********************
 * 最新一括（同期実行：Amazon → Qoo10 → Yahoo）
 **********************/
function 最新在庫をすべて取込() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const t0 = Date.now();

  log_('═══════════════════════════════════════════════════════════════');
  log_('【全モール在庫取込】開始');
  log_('═══════════════════════════════════════════════════════════════');

  // ========== Amazon ==========
  log_('');
  log_('┌─────────────────────────────────────────────────────────────┐');
  log_('│ STEP 1/3: Amazon在庫取込                                    │');
  log_('└─────────────────────────────────────────────────────────────┘');
  ss.toast('Amazon在庫の取込中…', '在庫取込 1/3', -1);

  try {
    最新Amazon在庫を取込();
    log_('✅ Amazon在庫取込: 完了');
  } catch (e) {
    log_('❌ Amazon在庫取込: エラー発生');
    logErrorDetail_('Amazon在庫取込', e);
  }

  // ========== Qoo10 ==========
  log_('');
  log_('┌─────────────────────────────────────────────────────────────┐');
  log_('│ STEP 2/3: Qoo10在庫取込                                     │');
  log_('└─────────────────────────────────────────────────────────────┘');
  ss.toast('Qoo10在庫の取込中…', '在庫取込 2/3', -1);

  try {
    最新Qoo10在庫を取込();
    log_('✅ Qoo10在庫取込: 完了');
  } catch (e) {
    log_('❌ Qoo10在庫取込: エラー発生');
    logErrorDetail_('Qoo10在庫取込', e);
  }

  // ========== Yahoo ==========
  log_('');
  log_('┌─────────────────────────────────────────────────────────────┐');
  log_('│ STEP 3/3: Yahoo在庫取込                                     │');
  log_('└─────────────────────────────────────────────────────────────┘');
  ss.toast('Yahoo在庫の取込中…', '在庫取込 3/3', -1);

  try {
    最新Yahoo在庫を取込();
    log_('✅ Yahoo在庫取込: 完了');
  } catch (e) {
    log_('❌ Yahoo在庫取込: エラー発生');
    logErrorDetail_('Yahoo在庫取込', e);
  }

  // ========== 完了 ==========
  const elapsed = Date.now() - t0;
  log_('');
  log_('═══════════════════════════════════════════════════════════════');
  log_(`【全モール在庫取込】完了 (合計: ${formatMs_(elapsed)})`);
  log_('═══════════════════════════════════════════════════════════════');

  ss.toast(`全モールの在庫取込が完了しました🎉 (${formatMs_(elapsed)})`, '在庫取込 完了', 10);
}

/**********************
 * ランチャー
 **********************/
function 最新Yahoo在庫を取込() {
  実行処理_高速_('Yahoo在庫取込', フォルダID_Yahoo在庫, {
    encoding: 'Shift_JIS',
    keepFirstCols: 4,
    rowBuffer: 50,
    colBuffer: 2
  });
}

function 最新Amazon在庫を取込() {
  実行処理_高速_('Amazon在庫取込', フォルダID_Amazon在庫, {
    keepFirstCols: 4,
    rowBuffer: 50,
    colBuffer: 2
  });
}

function 最新Qoo10在庫を取込() {
  実行処理_高速_('Qoo10在庫取込', フォルダID_Qoo10在庫, {
    pickHeaders: ['item_number', 'seller_unique_item_id', 'edit_type', 'seller_unique_option_id', 'quantity'],
    rowBuffer: 50,
    colBuffer: 2
  });
}

/**********************
 * メイン処理
 **********************/
function 実行処理_高速_(sheetName, folderId, opt) {
  const ssId = getSsId_();
  const t0 = Date.now();
  let t1, t2, t3;

  log_(`[${sheetName}] ▶ 処理開始`);
  log_(`[${sheetName}]   スプレッドシートID: ${ssId}`);
  log_(`[${sheetName}]   フォルダID: ${folderId}`);
  log_(`[${sheetName}]   オプション: ${JSON.stringify(opt)}`);

  // --- 0) シート確保 ---
  log_(`[${sheetName}] [STEP 0] シート確保開始...`);
  let sheetId;
  try {
    sheetId = ensureSheetIdByTitle_(ssId, sheetName);
    log_(`[${sheetName}] [STEP 0] シート確保完了 sheetId=${sheetId}`);
  } catch (e) {
    logErrorDetail_(`[${sheetName}] シート確保失敗`, e);
    return;
  }

  // --- 1) 最新ファイル取得 ---
  log_(`[${sheetName}] [STEP 1] 最新ファイル検索中...`);
  const fileMeta = 最新ファイルのメタデータを取得_(folderId, sheetName);
  if (!fileMeta) {
    const ss = SpreadsheetApp.getActive();
    ss.toast(`${sheetName}: フォルダ内に在庫ファイルが見つかりません`, '在庫取込', 8);
    log_(`[${sheetName}] [STEP 1] ❌ ファイルなし`);
    return;
  }
  log_(`[${sheetName}] [STEP 1] ファイル発見:`);
  log_(`[${sheetName}]   - タイトル: ${fileMeta.title}`);
  log_(`[${sheetName}]   - ID: ${fileMeta.id}`);
  log_(`[${sheetName}]   - MimeType: ${fileMeta.mimeType}`);
  log_(`[${sheetName}]   - 更新日時: ${fileMeta.modifiedDate}`);

  // --- 2) データ読み込み ---
  log_(`[${sheetName}] [STEP 2] ファイル読込開始...`);
  t1 = Date.now();
  const data = ファイルデータを読み込む_高速_(fileMeta, opt, sheetName);
  t2 = Date.now();

  if (!data || data.length === 0) {
    const ss = SpreadsheetApp.getActive();
    ss.toast(`${sheetName}: ファイルは見つかったけど中身が空です`, '在庫取込', 8);
    log_(`[${sheetName}] [STEP 2] ❌ データ空`);
    return;
  }

  log_(`[${sheetName}] [STEP 2] 読込完了:`);
  log_(`[${sheetName}]   - 行数: ${data.length.toLocaleString()}`);
  log_(`[${sheetName}]   - 列数: ${data[0].length}`);
  log_(`[${sheetName}]   - 読込時間: ${formatMs_(t2 - t1)}`);
  log_(`[${sheetName}]   - ヘッダー: [${data[0].slice(0, 6).join(', ')}${data[0].length > 6 ? ', ...' : ''}]`);

  const needRows = data.length + (opt.rowBuffer ?? 50);
  const needCols = data[0].length + (opt.colBuffer ?? 2);
  log_(`[${sheetName}]   - 必要グリッド: ${needRows}行 × ${needCols}列 (バッファ含む)`);

  // --- 3) グリッド調整 ---
  log_(`[${sheetName}] [STEP 3] グリッド調整...`);
  let gridInfo;
  try {
    if (sheetName === 'Yahoo在庫取込') {
      log_(`[${sheetName}] [STEP 3] Yahoo用: 行・列両方を拡張対象`);
      ensureGridAtLeast_(ssId, sheetId, sheetName, needRows, needCols);
      gridInfo = getGridSizeBySheetId_(ssId, sheetId);
    } else {
      log_(`[${sheetName}] [STEP 3] Amazon/Qoo10用: 行のみ拡張対象`);
      gridInfo = getGridSizeBySheetId_(ssId, sheetId);
      log_(`[${sheetName}]   - 現在のグリッド: ${gridInfo.rowCount}行 × ${gridInfo.columnCount}列`);

      if (gridInfo.rowCount < needRows) {
        log_(`[${sheetName}]   - 行が不足 → 拡張実行`);
        ensureRowsAtLeast_(ssId, sheetId, sheetName, needRows);
        gridInfo = getGridSizeBySheetId_(ssId, sheetId);
      } else {
        log_(`[${sheetName}]   - 行は十分 → 拡張不要`);
      }
    }
    log_(`[${sheetName}] [STEP 3] グリッド調整後: ${gridInfo.rowCount}行 × ${gridInfo.columnCount}列`);
  } catch (e) {
    logErrorDetail_(`[${sheetName}] グリッド調整失敗`, e);
    return;
  }

  // --- 4) 書き込み ---
  log_(`[${sheetName}] [STEP 4] シートへ書込開始...`);
  t3 = Date.now();

  try {
    // A1起点で data を書き込む
    log_(`[${sheetName}]   - データ書込 (${data.length}行 × ${data[0].length}列)...`);
    writeValuesInChunks_(ssId, sheetName, data, data[0].length);

    // 右側の余計列をクリア
    if (gridInfo.columnCount > data[0].length) {
      log_(`[${sheetName}]   - 右側クリア (列${data[0].length + 1}～${gridInfo.columnCount})...`);
      clearRightCols_(ssId, sheetName, data[0].length + 1, gridInfo.columnCount);
    } else {
      log_(`[${sheetName}]   - 右側クリア: 不要`);
    }

    // 下側の余計行をクリア
    const startRow = data.length + 1;
    const endRow = gridInfo.rowCount;
    const endCol = gridInfo.columnCount;
    if (endRow >= startRow) {
      log_(`[${sheetName}]   - 下側クリア (行${startRow}～${endRow})...`);
      clearBottomRows_(ssId, sheetName, startRow, endRow, endCol);
    } else {
      log_(`[${sheetName}]   - 下側クリア: 不要`);
    }

  } catch (e) {
    logErrorDetail_(`[${sheetName}] 書込失敗`, e);
    return;
  }

  const t4 = Date.now();
  log_(`[${sheetName}] [STEP 4] 書込完了 (${formatMs_(t4 - t3)})`);

  // --- 完了サマリー ---
  log_(`[${sheetName}] ▶ 処理完了`);
  log_(`[${sheetName}]   読込: ${formatMs_(t2 - t1)}`);
  log_(`[${sheetName}]   書込: ${formatMs_(t4 - t3)}`);
  log_(`[${sheetName}]   合計: ${formatMs_(t4 - t0)}`);
}

/**********************
 * ログユーティリティ
 **********************/
function log_(msg) {
  const now = new Date();
  const ts = Utilities.formatDate(now, 'Asia/Tokyo', 'HH:mm:ss.SSS');
  Logger.log(`[${ts}] ${msg}`);
}

function formatMs_(ms) {
  if (ms < 1000) return `${ms}ms`;
  if (ms < 60000) return `${(ms / 1000).toFixed(1)}秒`;
  const min = Math.floor(ms / 60000);
  const sec = ((ms % 60000) / 1000).toFixed(1);
  return `${min}分${sec}秒`;
}

/**********************
 * Sheets API：シート確保（既存確認→作成）
 **********************/
function ensureSheetIdByTitle_(ssId, title) {
  log_(`  [API] Sheets.get でシート一覧取得...`);
  const res = Sheets.Spreadsheets.get(ssId, { fields: 'sheets(properties(sheetId,title))' });
  const sheets = res.sheets || [];
  log_(`  [API] シート数: ${sheets.length}`);

  const hit = sheets.find(s => s.properties && s.properties.title === title);
  if (hit) {
    log_(`  [API] 既存シート発見: "${title}" (id=${hit.properties.sheetId})`);
    return hit.properties.sheetId;
  }

  log_(`  [API] シート "${title}" が存在しない → 新規作成`);
  const addRes = Sheets.Spreadsheets.batchUpdate({
    requests: [{ addSheet: { properties: { title } } }]
  }, ssId);

  const newId = addRes.replies[0].addSheet.properties.sheetId;
  log_(`  [API] シート作成完了: "${title}" (id=${newId})`);
  return newId;
}

/**********************
 * Sheets API：現在グリッド取得
 **********************/
function getGridSizeBySheetId_(ssId, sheetId) {
  const res = Sheets.Spreadsheets.get(ssId, {
    fields: 'sheets(properties(sheetId,gridProperties(rowCount,columnCount)))'
  });
  const sh = (res.sheets || []).find(s => s.properties && s.properties.sheetId === sheetId);
  if (!sh) throw new Error('sheetId not found: ' + sheetId);
  const gp = sh.properties.gridProperties || {};
  return { rowCount: gp.rowCount || 0, columnCount: gp.columnCount || 0 };
}

/**********************
 * Sheets API：足りない分だけ grid を増やす（Yahoo専用）
 **********************/
function ensureGridAtLeast_(ssId, sheetId, sheetName, needRows, needCols) {
  const cur = getGridSizeBySheetId_(ssId, sheetId);
  const curRows = cur.rowCount;
  const curCols = cur.columnCount;

  log_(`  [Grid] 現在: ${curRows}行 × ${curCols}列`);
  log_(`  [Grid] 必要: ${needRows}行 × ${needCols}列`);

  const requests = [];
  if (curCols < needCols) {
    const addCols = needCols - curCols;
    log_(`  [Grid] 列を ${addCols} 追加`);
    requests.push({ appendDimension: { sheetId, dimension: 'COLUMNS', length: addCols } });
  }
  if (curRows < needRows) {
    const addRows = needRows - curRows;
    log_(`  [Grid] 行を ${addRows} 追加`);
    requests.push({ appendDimension: { sheetId, dimension: 'ROWS', length: addRows } });
  }

  if (!requests.length) {
    log_(`  [Grid] 拡張不要`);
    return;
  }

  sheetsBatchUpdateWithRetry_(ssId, requests, sheetName + '_AppendGrid');
}

/**********************
 * Sheets API：足りない分だけ ROW だけ増やす（Amazon/Qoo10用）
 **********************/
function ensureRowsAtLeast_(ssId, sheetId, sheetName, needRows) {
  const cur = getGridSizeBySheetId_(ssId, sheetId);
  const curRows = cur.rowCount;

  if (curRows >= needRows) {
    log_(`  [Grid] 行拡張不要 (現在=${curRows} >= 必要=${needRows})`);
    return;
  }

  const add = needRows - curRows;
  log_(`  [Grid] 行を ${add} 追加 (${curRows} → ${needRows})`);

  const requests = [{
    appendDimension: {
      sheetId,
      dimension: 'ROWS',
      length: add
    }
  }];

  sheetsBatchUpdateWithRetry_(ssId, requests, sheetName + '_AppendRows');
}

/**********************
 * Values：分割書き込み
 **********************/
function writeValuesInChunks_(ssId, sheetName, data, cols) {
  const rows = data.length;
  const chunks = Math.ceil(rows / WRITE_CHUNK_ROWS);

  if (chunks > 1) {
    log_(`  [Write] ${rows.toLocaleString()}行を ${chunks} チャンクに分割`);
  }

  for (let i = 0; i < rows; i += WRITE_CHUNK_ROWS) {
    const chunk = data.slice(i, Math.min(i + WRITE_CHUNK_ROWS, rows));
    const startRow = i + 1;
    const endRow = startRow + chunk.length - 1;
    const range = `'${sheetName}'!A${startRow}:${colLetter_(cols)}${endRow}`;

    const chunkNum = Math.floor(i / WRITE_CHUNK_ROWS) + 1;
    log_(`  [Write] チャンク ${chunkNum}/${chunks}: 行${startRow}～${endRow} (${chunk.length}行)`);

    const t0 = Date.now();
    Sheets.Spreadsheets.Values.update(
      { values: chunk, majorDimension: 'ROWS' },
      ssId,
      range,
      { valueInputOption: 'RAW' }
    );
    log_(`  [Write] チャンク ${chunkNum}/${chunks}: 完了 (${formatMs_(Date.now() - t0)})`);
  }
}

/**********************
 * Values：右側クリア（colStart..endCol まで）
 **********************/
function clearRightCols_(ssId, sheetName, colStart, endCol) {
  if (!endCol || endCol < colStart) {
    log_(`  [Clear] 右側クリア: スキップ (範囲なし)`);
    return;
  }

  const startLetter = colLetter_(colStart);
  const endLetter   = colLetter_(endCol);
  const range = `'${sheetName}'!${startLetter}1:${endLetter}`;

  log_(`  [Clear] 右側クリア: ${startLetter}～${endLetter}列`);
  const t0 = Date.now();
  Sheets.Spreadsheets.Values.clear({}, ssId, range);
  log_(`  [Clear] 右側クリア: 完了 (${formatMs_(Date.now() - t0)})`);
}

/**********************
 * Values：下側クリア（startRow..endRow, A..endCol）
 **********************/
function clearBottomRows_(ssId, sheetName, startRow, endRow, endCol) {
  if (endRow < startRow) {
    log_(`  [Clear] 下側クリア: スキップ (余白なし)`);
    return;
  }

  const range = `'${sheetName}'!A${startRow}:${colLetter_(endCol)}${endRow}`;
  const rowCount = endRow - startRow + 1;

  log_(`  [Clear] 下側クリア: 行${startRow}～${endRow} (${rowCount.toLocaleString()}行)`);
  const t0 = Date.now();
  Sheets.Spreadsheets.Values.clear({}, ssId, range);
  log_(`  [Clear] 下側クリア: 完了 (${formatMs_(Date.now() - t0)})`);
}

/**********************
 * APIリトライ（grid関連）
 **********************/
function sheetsBatchUpdateWithRetry_(ssId, requests, label) {
  const maxTry = 6;
  for (let i = 0; i < maxTry; i++) {
    try {
      log_(`  [API] batchUpdate "${label}" (試行 ${i + 1}/${maxTry})...`);
      const t0 = Date.now();
      const res = Sheets.Spreadsheets.batchUpdate({ requests }, ssId);
      log_(`  [API] batchUpdate "${label}": 成功 (${formatMs_(Date.now() - t0)})`);
      return res;
    } catch (e) {
      log_(`  [API] batchUpdate "${label}": エラー発生`);
      if (!isRetryableError_(e)) {
        logErrorDetail_(`batchUpdate ${label}`, e);
        throw e;
      }
      sleepWithBackoff_(i, label);
    }
  }
  throw new Error(`Retry limit exceeded: ${label}`);
}

function isRetryableError_(e) {
  const msg = String(e && e.message ? e.message : e);
  return (
    msg.includes('unavailable') ||
    msg.includes('Rate Limit') ||
    msg.includes('quota') ||
    msg.includes('429') ||
    msg.includes('500') ||
    msg.includes('503')
  );
}

function sleepWithBackoff_(i, label) {
  const ms = (800 * Math.pow(2, i)) + Math.floor(Math.random() * 400);
  log_(`  [Retry] "${label}" 試行 ${i + 1} 失敗 → ${ms}ms 待機後リトライ`);
  Utilities.sleep(ms);
}

/**********************
 * エラー詳細ログ
 **********************/
function logErrorDetail_(label, e) {
  log_(`[ERROR] ${label}`);
  if (e && e.message) {
    log_(`  message: ${e.message}`);
  } else {
    log_(`  raw: ${e}`);
  }

  if (e && e.details) {
    try {
      log_(`  details: ${JSON.stringify(e.details)}`);
    } catch (err) {
      log_(`  details(JSON stringify失敗): ${e.details}`);
    }
  }

  if (e && e.stack) {
    log_(`  stack: ${e.stack}`);
  }
}

/**********************
 * ユーティリティ：列番号→A1文字
 **********************/
function colLetter_(colNum) {
  let letter = '';
  while (colNum > 0) {
    const mod = (colNum - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    colNum = Math.floor((colNum - 1) / 26);
  }
  return letter;
}

/**********************
 * Drive：最新ファイル取得（Advanced Drive）
 **********************/
function 最新ファイルのメタデータを取得_(folderId, sheetName) {
  const q = `'${folderId}' in parents and trashed = false and mimeType != 'application/vnd.google-apps.folder'`;

  log_(`  [Drive] フォルダ内ファイル検索...`);
  log_(`  [Drive] クエリ: ${q}`);

  try {
    const t0 = Date.now();
    const res = Drive.Files.list({
      q,
      orderBy: 'modifiedDate desc',
      maxResults: 5,  // 確認用に複数取得
      supportsTeamDrives: true,
      includeTeamDriveItems: true,
      fields: 'items(id,title,modifiedDate,mimeType,fileExtension,fileSize)'
    });

    const items = res.items || [];
    log_(`  [Drive] 検索完了 (${formatMs_(Date.now() - t0)})`);
    log_(`  [Drive] ファイル数: ${items.length}`);

    if (items.length === 0) {
      return null;
    }

    // 最新5件をログ出力
    items.slice(0, 5).forEach((f, i) => {
      const size = f.fileSize ? `${Math.round(f.fileSize / 1024)}KB` : '不明';
      log_(`  [Drive] ${i + 1}. ${f.title} (${size}, ${f.modifiedDate})`);
    });

    return items[0];
  } catch (e) {
    logErrorDetail_('最新ファイル取得', e);
    return null;
  }
}

/**********************
 * ファイル→配列（CSV/TSV/GoogleSheet/XLSX変換）
 **********************/
function ファイルデータを読み込む_高速_(fileMeta, opt, sheetName) {
  const fileId = fileMeta.id;
  const mime = fileMeta.mimeType;
  const title = fileMeta.title;
  const ext = (fileMeta.fileExtension || '').toLowerCase();

  log_(`  [Read] ファイル種別判定...`);
  log_(`  [Read]   MimeType: ${mime}`);
  log_(`  [Read]   拡張子: ${ext || '(なし)'}`);

  if (mime === 'application/vnd.google-apps.spreadsheet') {
    log_(`  [Read] → Googleスプレッドシートとして読込`);
    return Googleシート読込_高速_(fileId, opt, sheetName);
  }

  if (mime.includes('spreadsheetml') || title.endsWith('.xlsx') || ext === 'xlsx') {
    log_(`  [Read] → Excelファイルとして読込 (変換経由)`);
    return エクセルを配列に変換_高速_(fileId, opt, sheetName);
  }

  log_(`  [Read] → CSV/TSVとして読込`);
  try {
    const t0 = Date.now();
    log_(`  [Read] DriveApp.getFileById でBlob取得中...`);
    const blob = DriveApp.getFileById(fileId).getBlob().setName(title);
    log_(`  [Read] Blob取得完了 (${formatMs_(Date.now() - t0)})`);

    const data = cutColumns_(Blobをパース_(blob, opt, sheetName), opt);
    return data;
  } catch (e) {
    logErrorDetail_('ファイル読み込み', e);
    return [];
  }
}

function Googleシート読込_高速_(fileId, opt, sheetName) {
  try {
    log_(`  [Read] Sheets API でシート情報取得...`);
    const ssInfo = Sheets.Spreadsheets.get(fileId, { fields: 'sheets.properties' });
    if (!ssInfo.sheets || ssInfo.sheets.length === 0) {
      log_(`  [Read] シートが空`);
      return [];
    }

    const srcSheetName = ssInfo.sheets[0].properties.title;
    log_(`  [Read] 読込対象シート: "${srcSheetName}"`);

    let range = `'${srcSheetName}'`;
    if (opt.keepFirstCols && Number(opt.keepFirstCols) > 0) {
      range += `!A:${colLetter_(Number(opt.keepFirstCols))}`;
      log_(`  [Read] 列制限: 先頭${opt.keepFirstCols}列のみ`);
    }

    log_(`  [Read] Values.get 実行中... (range=${range})`);
    const t0 = Date.now();
    const result = Sheets.Spreadsheets.Values.get(fileId, range, {
      valueRenderOption: 'UNFORMATTED_VALUE',
      dateTimeRenderOption: 'FORMATTED_STRING'
    });
    log_(`  [Read] Values.get 完了 (${formatMs_(Date.now() - t0)})`);

    let data = result.values || [];
    log_(`  [Read] 取得行数: ${data.length}`);

    if (opt.pickHeaders) {
      log_(`  [Read] pickHeaders 適用: [${opt.pickHeaders.join(', ')}]`);
      data = pickHeaderColumns_(data, opt.pickHeaders);
      log_(`  [Read] pickHeaders 後の列数: ${data[0]?.length || 0}`);
    }

    return data;
  } catch (e) {
    log_(`  [Read] Sheets API 失敗 → 標準APIにフォールバック`);
    return Googleシート読込_標準_(fileId, opt, sheetName);
  }
}

function Googleシート読込_標準_(fileId, opt, sheetName) {
  log_(`  [Read] SpreadsheetApp.openById で読込...`);
  const src = SpreadsheetApp.openById(fileId);
  const sh = src.getSheets()[0];
  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();

  log_(`  [Read] シートサイズ: ${lr}行 × ${lc}列`);

  if (!lr || !lc) return [];

  const getCols = (opt.keepFirstCols) ? Math.min(Number(opt.keepFirstCols), lc) : lc;
  log_(`  [Read] 取得列数: ${getCols}`);

  let data = sh.getRange(1, 1, lr, getCols).getValues();

  if (opt.pickHeaders) {
    if (getCols !== lc) data = sh.getRange(1, 1, lr, lc).getValues();
    data = pickHeaderColumns_(data, opt.pickHeaders);
  }

  return data;
}

function pickHeaderColumns_(data, pickHeaders) {
  if (!data || !data.length) return data;
  const headers = data[0].map(v => String(v ?? '').trim());
  const idxs = pickHeaders.map(h => headers.indexOf(h)).filter(i => i >= 0);

  log_(`  [Read] ヘッダーマッチ: ${idxs.length}/${pickHeaders.length}`);

  if (idxs.length === 0) {
    log_(`  [Read] ⚠ マッチするヘッダーなし！元データをそのまま返却`);
    return data;
  }

  return data.map(r => idxs.map(i => r[i]));
}

function エクセルを配列に変換_高速_(fileId, opt, sheetName) {
  let tempFileId = null;
  try {
    log_(`  [Read] Drive.Files.copy でGoogleシートに変換中...`);
    const t0 = Date.now();
    const resource = { title: 'tmp_' + Date.now(), mimeType: MimeType.GOOGLE_SHEETS };
    const convertFile = Drive.Files.copy(resource, fileId, { convert: true, supportsTeamDrives: true });
    tempFileId = convertFile.id;
    log_(`  [Read] 変換完了 (${formatMs_(Date.now() - t0)}) tempId=${tempFileId}`);

    const data = Googleシート読込_高速_(tempFileId, opt, sheetName);
    return data;
  } catch (e) {
    logErrorDetail_('Excel変換', e);
    return [];
  } finally {
    if (tempFileId) {
      try {
        log_(`  [Read] 一時ファイル削除: ${tempFileId}`);
        Drive.Files.trash(tempFileId);
      } catch (e) {
        log_(`  [Read] 一時ファイル削除失敗（無視）`);
      }
    }
  }
}

function Blobをパース_(blob, opt, sheetName) {
  log_(`  [Parse] Blob解析開始...`);

  let text = '';
  const encoding = opt.encoding || null;
  log_(`  [Parse] エンコーディング: ${encoding || '自動検出'}`);

  try {
    text = blob.getDataAsString(encoding);
  } catch (e) {
    log_(`  [Parse] 指定エンコーディングで失敗 → デフォルトで再試行`);
    text = blob.getDataAsString();
  }

  if (!text) {
    log_(`  [Parse] テキスト内容が空`);
    return [];
  }

  log_(`  [Parse] テキスト長: ${text.length.toLocaleString()}文字`);

  let delim = opt.delimiter;
  if (!delim) {
    const sample = text.substring(0, 1000);
    const tabCount = (sample.match(/\t/g) || []).length;
    const commaCount = (sample.match(/,/g) || []).length;
    delim = (tabCount > commaCount) ? '\t' : ',';
    log_(`  [Parse] 区切り文字検出: タブ=${tabCount}, カンマ=${commaCount} → "${delim === '\t' ? 'TAB' : 'COMMA'}"`);
  }

  log_(`  [Parse] CSV解析中...`);
  const t0 = Date.now();
  const data = Utilities.parseCsv(text, delim);
  log_(`  [Parse] CSV解析完了 (${formatMs_(Date.now() - t0)}): ${data.length}行`);

  return data;
}

function cutColumns_(data, opt) {
  if (!data || !data.length) return data;

  if (opt.keepFirstCols) {
    const n = Number(opt.keepFirstCols);
    log_(`  [Cut] 先頭${n}列に制限`);
    return data.map(r => r.slice(0, n));
  }

  if (opt.pickHeaders) {
    return pickHeaderColumns_(data, opt.pickHeaders);
  }

  return data;
}