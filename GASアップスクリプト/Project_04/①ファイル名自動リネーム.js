/**
 * Google Driveフォルダ内の画像ファイルを商品コードに基づいてリネーム
 * 高速版 v3: Advanced Drive Service (Drive API v3) 使用
 *
 * ★ 事前準備:
 *   GASエディタ → サービス(+) → Drive API → 追加
 *   ※ v3 が選択されていることを確認
 */

// ====================================================================
// 定数
// ====================================================================
const CONFIG = {
  PARENT_FOLDER_ID: '1Vk1cVgrXM5CNFqfBvAaBUJpheDn4f9Lz',
  SHEET_NAME: '①商品入力シート',
  CODE_COLUMN: 3,        // C列
  DATA_START_ROW: 3,     // ヘッダー2行をスキップ
  PAGE_SIZE: 1000,       // Drive API 1回あたりの取得件数
};

// ====================================================================
// UI ヘルパー（UIが無い環境でも安全）
// ====================================================================
function _alert(title, msg) {
  try {
    SpreadsheetApp.getUi().alert(title, msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (_) {
    try { SpreadsheetApp.getActiveSpreadsheet().toast(msg, title, 10); } catch (_2) {}
    Logger.log(`[${title}] ${msg}`);
  }
}

function _confirm(title, msg) {
  try {
    var ui = SpreadsheetApp.getUi();
    return ui.alert(title, msg, ui.ButtonSet.YES_NO) === ui.Button.YES;
  } catch (_) {
    Logger.log(`[確認スキップ] ${title}`);
    return true;
  }
}

// ====================================================================
// Drive API v3 ヘルパー
// ====================================================================

/**
 * 親フォルダ直下のサブフォルダを全て取得し { フォルダ名 → フォルダID } の Map を返す
 */
function _buildFolderMap(parentFolderId) {
  const map = {};
  let pageToken = null;
  const q = `'${parentFolderId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`;

  do {
    const res = Drive.Files.list({
      q: q,
      fields: 'nextPageToken, files(id, name)',
      pageSize: CONFIG.PAGE_SIZE,
      pageToken: pageToken,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
    });
    (res.files || []).forEach(f => { map[f.name] = f.id; });
    pageToken = res.nextPageToken;
  } while (pageToken);

  return map;
}

/**
 * 指定フォルダ内の全ファイルを取得
 * @return {Array<{id, name}>}
 */
function _listFiles(folderId) {
  const files = [];
  let pageToken = null;
  const q = `'${folderId}' in parents and mimeType!='application/vnd.google-apps.folder' and trashed=false`;

  do {
    const res = Drive.Files.list({
      q: q,
      fields: 'nextPageToken, files(id, name)',
      pageSize: CONFIG.PAGE_SIZE,
      pageToken: pageToken,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
    });
    (res.files || []).forEach(f => files.push(f));
    pageToken = res.nextPageToken;
  } while (pageToken);

  return files;
}

/**
 * ファイルをリネーム（Drive API v3）
 */
function _renameFile(fileId, newName) {
  Drive.Files.update({ name: newName }, fileId, null, { supportsAllDrives: true });
}

// ====================================================================
// ファイル名ユーティリティ
// ====================================================================

function _baseName(filename) {
  return filename.replace(/\.[^.]+$/, '');
}

function _extension(filename) {
  const m = filename.match(/\.[^.]+$/);
  return m ? m[0] : '';
}

/**
 * ソートキー抽出（数値系パターンを認識）
 */
function _sortKey(filename) {
  const base = _baseName(filename);

  // imgi_数字 / imgi-数字
  let m = base.match(/imgi[_-]?(\d+)/i);
  if (m) return { n: true, v: parseInt(m[1], 10) };

  // getImage (数字)
  m = base.match(/getImage\s*\((\d+)\)/i);
  if (m) return { n: true, v: parseInt(m[1], 10) };

  // getImage 単体 → 0
  if (/^getImage$/i.test(base)) return { n: true, v: 0 };

  // 数字のみ
  m = base.match(/^(\d+)$/);
  if (m) return { n: true, v: parseInt(m[1], 10) };

  return { n: false, v: base };
}

function _compareFiles(a, b) {
  const ka = _sortKey(a.name);
  const kb = _sortKey(b.name);

  if (ka.n && kb.n) {
    if (ka.v !== kb.v) return ka.v - kb.v;
    return a.name.localeCompare(b.name, 'ja');
  }
  if (ka.n) return -1;
  if (kb.n) return 1;
  return a.name.localeCompare(b.name, 'ja');
}

/**
 * フォルダ内のファイルをリネーム（Batchリクエスト版・免責除外修正）
 * ★ notice_ から始まるファイルはリネーム対象から除外します
 */
function _renameFilesInFolder(files, code) {
  if (files.length === 0) {
    return { success: false, message: 'ファイルなし' };
  }

  // 全ファイルが既にフォルダ名と一致 → スキップ
  const allMatch = files.every(f => _baseName(f.name) === code);
  if (allMatch) {
    return { success: true, allMatched: true, renamedCount: 0, skippedCount: files.length };
  }

  // ソート
  files.sort(_compareFiles);

  const renameQueue = [];
  let skippedCount = 0;
  let seq = 0;

  for (const file of files) {
    const name = file.name;
    const base = _baseName(name);
    const ext = _extension(name);

    // ▼▼▼ 修正: 免責画像（notice_ や 免責）は絶対に触らない ▼▼▼
    if (name.indexOf('notice_') === 0 || name.indexOf('免責') !== -1) {
      skippedCount++;
      continue;
    }
    // ▲▲▲ 修正ここまで ▲▲▲

    // メイン画像（フォルダ名と一致）はスキップ
    if (base === code) {
      skippedCount++;
      continue;
    }

    seq++;
    const newName = `${code}_${seq}${ext}`;

    // 名前が変わらないならスキップ
    if (name === newName) {
      skippedCount++;
      continue;
    }

    renameQueue.push({ id: file.id, newName: newName, oldName: name });
  }

  // リネーム対象がなければ終了
  if (renameQueue.length === 0) {
    return { success: true, allMatched: false, renamedCount: 0, skippedCount: skippedCount };
  }

  // Batch実行
 for (const item of renameQueue) {
    Drive.Files.update({ name: item.newName }, item.id, null, {
      supportsAllDrives: true
    });
  }

  return { success: true, allMatched: false, renamedCount: renameQueue.length, skippedCount };
}

// ====================================================================
// メイン処理
// ====================================================================

/**
 * スプレッドシートのC列にある全商品コードを一括処理（高速版）
 */
function 全商品一括リネーム() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    _alert('エラー', `${CONFIG.SHEET_NAME} が見つかりません。`);
    return;
  }

  if (!_confirm('確認', 'C列の全商品コードを対象に画像リネームを実行します。\nよろしいですか？')) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) {
    _alert('エラー', 'データが見つかりません');
    return;
  }

  // 商品コード取得 → 重複除去・空白除外
  const rawCodes = sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.CODE_COLUMN, lastRow - CONFIG.DATA_START_ROW + 1, 1).getValues();
  const codes = [...new Set(
    rawCodes.map(r => String(r[0] ?? '').trim()).filter(c => c && c !== '商品コード')
  )];

  Logger.log(`=== 一括処理開始（高速版） ===`);
  Logger.log(`対象商品コード数: ${codes.length}件（重複除去済み）`);

  // ★ 高速化ポイント1: サブフォルダ一覧を1回で取得
  const startMap = Date.now();
  const folderMap = _buildFolderMap(CONFIG.PARENT_FOLDER_ID);
  Logger.log(`フォルダマップ構築: ${Object.keys(folderMap).length}件 (${Date.now() - startMap}ms)`);

  let processedCount = 0;
  let skippedCount = 0;
  let errorCount = 0;
  let notFoundCount = 0;
  const logs = [];

  const startProcess = Date.now();

  for (const code of codes) {
    const folderId = folderMap[code];

    if (!folderId) {
      notFoundCount++;
      logs.push(`- ${code}: フォルダなし`);
      continue;
    }

    try {
      // ★ 高速化ポイント2: Drive API v3 で一括取得
      const files = _listFiles(folderId);
      const result = _renameFilesInFolder(files, code);

      if (result.success) {
        if (result.allMatched) {
          skippedCount++;
        } else {
          processedCount++;
          logs.push(`✓ ${code}: ${result.renamedCount}個リネーム, ${result.skippedCount}個スキップ`);
        }
      } else {
        errorCount++;
        logs.push(`✗ ${code}: ${result.message}`);
      }
    } catch (e) {
      errorCount++;
      logs.push(`✗ ${code}: ${e.message}`);
    }
  }

  const elapsed = ((Date.now() - startProcess) / 1000).toFixed(1);

  Logger.log(`\n=== 処理完了 (${elapsed}秒) ===`);
  Logger.log(`リネーム実行: ${processedCount}件`);
  Logger.log(`変更不要スキップ: ${skippedCount}件`);
  Logger.log(`フォルダなし: ${notFoundCount}件`);
  Logger.log(`エラー: ${errorCount}件`);
  if (logs.length > 0) {
    Logger.log('\n詳細:');
    logs.forEach(m => Logger.log(m));
  }

  _alert('処理完了',
    `処理時間: ${elapsed}秒\n` +
    `リネーム実行: ${processedCount}件\n` +
    `変更不要: ${skippedCount}件\n` +
    `フォルダなし: ${notFoundCount}件\n` +
    `エラー: ${errorCount}件\n\n` +
    `詳細は 表示 → ログ`
  );
}

/**
 * 単一の商品コードでテスト
 */
function テスト実行() {
  const testCode = 'AOHAKO19-TW';

  Logger.log(`=== テスト: ${testCode} ===`);

  const folderMap = _buildFolderMap(CONFIG.PARENT_FOLDER_ID);
  const folderId = folderMap[testCode];

  if (!folderId) {
    Logger.log('フォルダが見つかりません');
    _alert('テスト結果', `${testCode}: フォルダが見つかりません`);
    return;
  }

  const files = _listFiles(folderId);
  const result = _renameFilesInFolder(files, testCode);

  if (result.success) {
    if (result.allMatched) {
      Logger.log('すべてフォルダ名一致 → スキップ');
    } else {
      Logger.log(`リネーム: ${result.renamedCount}件, スキップ: ${result.skippedCount}件`);
    }
  } else {
    Logger.log(`失敗: ${result.message}`);
  }

  _alert('テスト完了', `商品コード: ${testCode}\n詳細はログを確認`);
}

/**
 * ソート順序のテスト
 */
function ソート順序テスト() {
  const testFiles = [
    { name: '15.jpg' }, { name: 'imgi_8_getImage.jpg' },
    { name: '3.jpg' }, { name: 'getImage.jpg' },
    { name: 'imgi_10_getImage.jpg' }, { name: 'getImage (1).jpg' },
    { name: '1.jpg' }, { name: 'getImage (2).jpg' },
    { name: 'imgi_9_getImage.jpg' }, { name: '20.jpg' },
    { name: 'TWS-KRBL-NZS01.jpeg' }, { name: 'AOHAKO19-TW_1.jpg' },
  ];

  testFiles.sort(_compareFiles);

  Logger.log('=== ソート順序テスト ===');
  testFiles.forEach((f, i) => {
    Logger.log(`${i + 1}. ${f.name} (key: ${JSON.stringify(_sortKey(f.name))})`);
  });
}

/**
 * 現在の設定を表示
 */
function 設定確認() {
  Logger.log('=== 設定 ===');
  Logger.log(`親フォルダID: ${CONFIG.PARENT_FOLDER_ID}`);
  Logger.log(`シート名: ${CONFIG.SHEET_NAME}`);
  Logger.log(`列: C列, 開始行: ${CONFIG.DATA_START_ROW}`);
  Logger.log(`Drive API: v3 (Advanced Drive Service)`);
}

// ====================================================================
// 追加: 複数フォルダの中身をまとめて取得して高速化
// ====================================================================

function _chunkArray(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

/**
 * 複数フォルダ内のファイルをまとめて取得して
 * folderId => files[] に仕分けして返す
 *
 * @param {string[]} folderIds
 * @return {Object<string, Array<{id,name,parents}>>}
 */
function _listFilesInFoldersBulk(folderIds) {
  const byFolder = {};
  folderIds.forEach(id => (byFolder[id] = []));

  // クエリ長制限回避のため分割（50前後が安全）
  const CHUNK = 50;
  const chunks = _chunkArray(folderIds, CHUNK);

  for (const ids of chunks) {
    let pageToken = null;

    // ORで束ねる
    const parentOr = ids.map(id => `'${id}' in parents`).join(' or ');
    const q = `(${parentOr}) and mimeType!='application/vnd.google-apps.folder' and trashed=false`;

    do {
      const res = Drive.Files.list({
        q,
        fields: 'nextPageToken, files(id, name, parents)',
        pageSize: 1000,
        pageToken,
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
      });

      for (const f of (res.files || [])) {
        const ps = f.parents || [];
        // 通常1親やけど念のため全部見る
        for (const p of ps) {
          if (byFolder[p]) byFolder[p].push(f);
        }
      }

      pageToken = res.nextPageToken;
    } while (pageToken);
  }

  return byFolder;
}

/**
 * 親フォルダ配下の「全サブフォルダ」を対象に一括リネーム（超一括版）
 * ※シートは見ない。Drive構造だけで回す。
 */
function 親フォルダ配下_全フォルダ一括リネーム() {
  if (!_confirm('確認', '親フォルダ配下の全サブフォルダを対象に画像リネームします。\nよろしいですか？')) return;

  const start = Date.now();

  // 1) サブフォルダ一覧（フォルダ名=商品コード想定）
  const folderMap = _buildFolderMap(CONFIG.PARENT_FOLDER_ID); // {name:id}
  const codes = Object.keys(folderMap);
  const folderIds = codes.map(c => folderMap[c]);

  Logger.log(`サブフォルダ数: ${codes.length}`);

  // 2) まとめてファイル取得（ここが高速化の肝）
  const t1 = Date.now();
  const filesByFolder = _listFilesInFoldersBulk(folderIds);
  Logger.log(`一括ファイル取得: ${(Date.now() - t1)}ms`);

  // 3) リネーム
  let processed = 0, skipped = 0, error = 0, empty = 0;

  for (const code of codes) {
    const folderId = folderMap[code];
    const files = filesByFolder[folderId] || [];

    if (files.length === 0) {
      empty++;
      continue;
    }

    try {
      const result = _renameFilesInFolder(files, code);
      if (result.success) {
        if (result.allMatched) skipped++;
        else processed++;
      } else {
        error++;
      }
    } catch (e) {
      error++;
      Logger.log(`✗ ${code}: ${e.message}`);
    }
  }

  const elapsed = ((Date.now() - start) / 1000).toFixed(1);

  _alert('完了',
    `処理時間: ${elapsed}秒\n` +
    `リネーム実行: ${processed}件\n` +
    `変更不要: ${skipped}件\n` +
    `空フォルダ: ${empty}件\n` +
    `エラー: ${error}件\n\n` +
    `詳細は 表示 → ログ`
  );
}
/**
 * Drive API Batch Endpoint を叩いて一括リネームを実行する（安全・高速版）
 * @param {Array<{id:string, newName:string, oldName?:string}>} items
 */

// =====================================================
// DBG: フォルダ内の「全ファイル」を暴露するスクリプト
// =====================================================
function DBG_ファイル数がおかしいフォルダを調査() {
  // ★ここに「12個もないはずだ！」と思う商品コードを入力
  var TARGET_CODE = 'TWS0003-CM-01'; 

  Logger.log('=== 徹底調査: ' + TARGET_CODE + ' ===');
  
  var folderMap = IMG_drive_buildFolderMap_(CONFIG_IMAGES.商品画像ルートID);
  var folderId = folderMap[TARGET_CODE];
  
  if (!folderId) {
    Logger.log('❌ フォルダ自体が見つかりません');
    return;
  }

  // フィルタなしで全取得
  var files = IMG_drive_listFilesInFolders_parallel_([folderId])[folderId];
  
  Logger.log('📂 フォルダID: ' + folderId);
  Logger.log('📊 スクリプトが見つけたファイル総数: ' + files.length + ' 個');
  Logger.log('--------------------------------------------------');
  
  // 内訳を表示
  files.sort(function(a, b) { return a.name.localeCompare(b.name); });
  
  files.forEach(function(f, i) {
    Logger.log((i + 1) + '. [' + f.name + '] (ID: ' + f.id + ')');
    // もし「画像じゃない」怪しいやつなら警告
    if (String(f.mimeType).indexOf('image/') === -1) {
      Logger.log('   ⚠️ ↑ これは画像ではありません！ (' + f.mimeType + ')');
    }
  });
}

// ====================================================================
// 権限取得用のおまじない（削除しないでください）
// ====================================================================
function _forceDriveScope() {
  DriveApp.getRootFolder();
}