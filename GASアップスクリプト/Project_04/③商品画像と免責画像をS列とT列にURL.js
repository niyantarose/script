/******************************************************
 * 30_images_drive_links_ultra.gs
 * 画像URL挿入（Drive直リンク）+ 免責画像Batchコピー + Drive API fetchAll並列 “統合・一括修正版”
 *
 * ✅ Drive API(v3) でフォルダ/ファイル一覧を高速取得
 * ✅ UrlFetchApp.fetchAll で商品フォルダのファイル一覧を並列取得（1ページ目を並列、続きは必要分のみ追撃）
 * ✅ 免責画像はカテゴリ別に「番号順で抽出」→ 商品フォルダへ batch copy で一括コピー
 * ✅ 商品画像は数字優先で順番どおり
 * ✅ 免責画像は必ず最後（固定名順）
 * ✅ S列: 1つURLをリンク化（メイン画像）
 * ✅ T列: 保存は「;」区切りで投入 → RichText化時に表示は改行＆各行リンク
 *   ※Sheets仕様上、RichTextのテキストがセルの表示/値になるため、
 *     最終的にセルは改行表示のテキストになります（ただし復元可能：splitは ;/改行両対応）
 *
 * ★準備（必須）
 * 1) Apps Script エディタ → サービス(+) →「Drive API」追加（v3）
 * 2) （求められたら）GCP側で Drive API を有効化
 ******************************************************/

// ▼▼▼ 設定エリア（ここを自分の環境に合わせて書き換えてください） ▼▼▼
var CONFIG_IMAGES = {
  // フォルダID（ブラウザのURL末尾のID）
  商品画像ルートID: '1Vk1cVgrXM5CNFqfBvAaBUJpheDn4f9Lz',

  // 免責画像ルート（親フォルダ or カテゴリフォルダ自身 どちらでもOK）
  免責画像ルートID: '1qfwT1NC3wwKgOLx2rvnSH836Zc0hROIa',

  // シート設定
  シート名_商品入力: '①商品入力シート',
  シート名_Yahoo: 'Yahoo商品登録',

  // ヘッダー行数
  ヘッダー行数_商品入力: 2,
  ヘッダー行数_Yahoo: 2,

  // 列番号（A=1, B=2...）
  COL_CODE: 3,  // 商品コード (C列)
  COL_NAME: 4,  // 商品名 (D列) ※免責判定用
  COL_S: 19,    // メイン画像 (S列)
  COL_T: 20,    // 詳細画像 (T列)

  // 画像枚数制限
  MAX_TOTAL: 21,   // 合計最大枚数
  MAX_DETAIL: 20,  // 詳細画像最大枚数
  MAX_DISC: 3,     // 免責画像最大枚数（実際はカテゴリで上限決める）

  // fetchAll の並列数（多すぎると429出る）
  FETCHALL_BATCH: 80,

  // Batch copy の一回の件数（多すぎると失敗しやすい）
  BATCH_COPY_SIZE: 80,

  // コピー後の反映待ち（ms）
  POST_COPY_SLEEP_MS: 1500,

  // URLの形式（VPSが直接取得するなら export=download の方が安定しやすいことがある）
  URL_EXPORT_MODE: 'view' // 'view' or 'download'
};
// ▲▲▲ 設定エリアここまで ▲▲▲


// ====================================================================
// 免責画像の固定ファイル名マッピング（“商品フォルダに入る最終名”）
// ====================================================================
var 免責画像固定名 = {
  '韓国マンガ初版限定免責画像': ['notice_kr_manga_first_01.jpg', 'notice_kr_manga_first_02.jpg'],
  '韓国マンガ一般版免責画像':   ['notice_kr_manga_standard_01.jpg', 'notice_kr_manga_standard_02.jpg'],
  '韓国マンガ特装版免責画像':   ['notice_kr_manga_special_01.jpg', 'notice_kr_manga_special_02.jpg'],
  '韓国マンガ限定版免責画像':   ['notice_kr_manga_limited_01.jpg', 'notice_kr_manga_limited_02.jpg'],
  'グッズ免責画像':             ['notice_goods_01.jpg'],
  '台湾免責画像':               ['notice_tw_goods_01.jpg', 'notice_tw_goods_02.jpg'],
  // ↓↓↓ 追加 ↓↓↓
  '中国書籍免責画像':           ['notice_cn_book_01.jpg', 'notice_cn_book_02.jpg'],
  '中国グッズ免責画像':         ['notice_cn_goods_01.jpg'],
  '雑誌免責画像':               ['notice_magazine_01.jpg']
};


// ====================================================================
// ★ 安全なUI通知（UIが無くても落ちない）
// ====================================================================
function IMGCV_uiSafeAlert_(msg) {
  try {
    SpreadsheetApp.getUi().alert(String(msg));
    return;
  } catch (e) {}
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(String(msg), '画像URL挿入', 10);
  } catch (e2) {}
  Logger.log('[IMGCV_UI_FALLBACK] ' + msg);
}

function IMGCV_uiSafeToast_(msg, title, sec) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(String(msg), title || 'IMGCV', sec || 5);
  } catch (e) {
    Logger.log('[IMGCV_TOAST_FALLBACK] ' + (title ? title + ': ' : '') + msg);
  }
}


// ====================================================================
// 実行用関数
// ====================================================================

/** 実行：①商品入力シート */
function 商品画像URLを挿入_超高速_商品入力シート() {
  商品画像URLを挿入_超高速_汎用_(CONFIG_IMAGES.シート名_商品入力, CONFIG_IMAGES.ヘッダー行数_商品入力);
}

/** 実行：Yahoo商品登録シート */
function 商品画像URLを挿入_超高速_Yahoo商品登録() {
  商品画像URLを挿入_超高速_汎用_(CONFIG_IMAGES.シート名_Yahoo, CONFIG_IMAGES.ヘッダー行数_Yahoo);
}


// ====================================================================
// メイン（超高速）
// ====================================================================

/**
 * 超高速メイン
 * - フォルダMap構築（Drive API）
 * - 商品フォルダファイル一覧（fetchAll並列）
 * - 免責ソース（カテゴリごとに一回だけ取得＆番号順ソート）
 * - 免責コピー計画 → batch copy
 * - コピー後、影響フォルダのみ再スキャン
 * - URL生成 → 一括書き込み → RichTextリンク化
 */
function 商品画像URLを挿入_超高速_汎用_(sheetName, headerRows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) { IMGCV_uiSafeAlert_('シートが見つかりません: ' + sheetName); return; }

  var lastRow = sheet.getLastRow();
  if (lastRow <= headerRows) { IMGCV_uiSafeAlert_('データ行がありません: ' + sheetName); return; }

  // 列インデックス（0始まり）を推定（ヘッダー名探索 → 見つからなければ固定列）
  var codeCol = IMG_検索列位置_ヘッダー名_複数行_(sheet, headerRows, ['code','item-code','商品コード']);
　var nameCol = IMG_検索列位置_ヘッダー名_複数行_(sheet, headerRows, ['name','商品名','商品名称','商品名 ', '商品名　']);
  var idxCode = (codeCol === -1) ? (CONFIG_IMAGES.COL_CODE - 1) : (codeCol - 1);
  var idxName = (nameCol === -1) ? (CONFIG_IMAGES.COL_NAME - 1) : (nameCol - 1);

  var startRow = headerRows + 1;
  var numRows = lastRow - headerRows;

  // 必要列数まで一括取得（落ちないように）
  var needCols = Math.max(idxCode + 1, idxName + 1, CONFIG_IMAGES.COL_CODE, CONFIG_IMAGES.COL_NAME);
  var values = sheet.getRange(startRow, 1, numRows, needCols).getValues();

  // 行ごとのコード・商品名を作る
  var codes = new Array(numRows);
  var names = new Array(numRows);

  var uniqueCodes = [];
  var codeSeen = {};

  for (var i = 0; i < numRows; i++) {
    var c = String(values[i][idxCode] || '').trim();
    var n = String(values[i][idxName] || '').trim();
    codes[i] = c;
    names[i] = n;

    if (c && !codeSeen[c]) {
      codeSeen[c] = true;
      uniqueCodes.push(c);
    }
  }

  if (uniqueCodes.length === 0) {
    IMGCV_uiSafeAlert_('商品コードが見つからないため処理を終了しました。');
    return;
  }

  IMGCV_uiSafeToast_('フォルダMap構築中…', 'IMG', 5);

  // 1) 商品ルート直下：フォルダMap（name->id）
  var productFolderMap = IMG_drive_buildFolderMap_(CONFIG_IMAGES.商品画像ルートID);

  // 2) 免責ルートの扱い（親 or 自分がカテゴリ）
  var disclaimerRootInfo = IMG_drive_getFileMeta_(CONFIG_IMAGES.免責画像ルートID); // {id,name,mimeType}
  var disclaimerFolderMap = IMG_drive_buildFolderMap_(CONFIG_IMAGES.免責画像ルートID); // 親直下のカテゴリ探し用

  // 3) 行→商品フォルダIDを解決（存在しないコードもある）
  var rowFolderIds = new Array(numRows);   // row -> folderId or ''
  var neededFolderIds = [];
  var neededFolderSeen = {};

  var noFolder = 0;

  for (var r = 0; r < numRows; r++) {
    var code = codes[r];
    if (!code) { rowFolderIds[r] = ''; continue; }

    var fid = productFolderMap[code];
    if (!fid) {
      rowFolderIds[r] = '';
      noFolder++;
      continue;
    }
    rowFolderIds[r] = fid;

    if (!neededFolderSeen[fid]) {
      neededFolderSeen[fid] = true;
      neededFolderIds.push(fid);
    }
  }

  IMGCV_uiSafeToast_('商品フォルダの中身を並列取得…', 'IMG', 5);

  // 4) 商品フォルダ内の画像一覧を “並列” で取得（基本1ページ目は fetchAll）
  //    1000枚を超えるフォルダは稀なので、nextPageTokenがある分だけ追撃
  var productFilesCache = IMG_drive_listFilesInFolders_parallel_(neededFolderIds);

  // 5) 免責ソース（カテゴリごとに一回）を取得（番号順でソートして先頭を使う）
  IMGCV_uiSafeToast_('免責ソースを準備…', 'IMG', 5);
  var disclaimerSource = IMG_disclaimer_prepareSources_(disclaimerRootInfo, disclaimerFolderMap);

  // 6) コピー計画
  var copyQueue = []; // {sourceId, targetFolderId, newName}
  var affectedFolder = {}; // targetFolderId -> true
  var noDiscHit = 0;

  for (var rr = 0; rr < numRows; rr++) {
    var folderId = rowFolderIds[rr];
    if (!folderId) continue;

    var discType = IMG_判定免責フォルダ名_(names[rr]);
    if (!discType) continue;

    var required = IMG_免責カテゴリ_必要ファイル名_(discType);
    if (!required.length) continue;

    // 既存名一覧（小規模なので配列でOK、規模が大きいならset化）
    var files = productFilesCache[folderId] || [];
    var nameSet = {};
    for (var k = 0; k < files.length; k++) {
      nameSet[String(files[k].name)] = true;
    }

    // ソースID配列（番号順）
    var srcIds = disclaimerSource[discType] || [];
    var needAny = false;

    for (var j = 0; j < required.length; j++) {
      var targetName = required[j];

      if (nameSet[targetName]) continue; // 既にある

      var srcId = srcIds[j]; // 01→0, 02→1…
      if (!srcId) continue;

      copyQueue.push({ sourceId: srcId, targetFolderId: folderId, newName: targetName });
      affectedFolder[folderId] = true;
      needAny = true;
    }

    if (needAny && srcIds.length === 0) noDiscHit++;
  }

  // 7) Batch copy 実行
  if (copyQueue.length > 0) {
    IMGCV_uiSafeToast_('免責を一括コピー中… (' + copyQueue.length + '件)', 'IMG', 8);
    IMG_drive_executeBatchCopy_(copyQueue);

    // 反映待ち
    if (CONFIG_IMAGES.POST_COPY_SLEEP_MS) Utilities.sleep(CONFIG_IMAGES.POST_COPY_SLEEP_MS);

    // 影響フォルダのみ再スキャン（ID確定のため）
    var affectedIds = Object.keys(affectedFolder);
    IMGCV_uiSafeToast_('コピー反映フォルダを再取得… (' + affectedIds.length + ')', 'IMG', 6);
    var updated = IMG_drive_listFilesInFolders_parallel_(affectedIds);
    for (var a = 0; a < affectedIds.length; a++) {
      var afid = affectedIds[a];
      productFilesCache[afid] = updated[afid] || [];
    }
  } else {
    IMGCV_uiSafeToast_('新規コピーは不要でした', 'IMG', 3);
  }

  // 8) URL生成 & 出力
  IMGCV_uiSafeToast_('URL生成…', 'IMG', 5);

  var outS = new Array(numRows);
  var outT = new Array(numRows);

  var hitRows = 0;

  for (var x = 0; x < numRows; x++) {
    var code2 = codes[x];
    if (!code2) { outS[x] = ['']; outT[x] = ['']; continue; }

    var fid2 = rowFolderIds[x];
    if (!fid2) { outS[x] = ['']; outT[x] = ['']; continue; }

    var files2 = productFilesCache[fid2] || [];
    var discType2 = IMG_判定免責フォルダ名_(names[x]);

    var urls = IMG_generateSortedUrls_(files2, discType2);
    if (urls.length === 0) {
      outS[x] = [''];
      outT[x] = [''];
      continue;
    }

    outS[x] = [urls[0]];

    // 保存は ; 区切り（その後RichText化で表示は改行になる）
    var detail = urls.slice(1, 1 + CONFIG_IMAGES.MAX_DETAIL).join(';');
    outT[x] = [detail];

    hitRows++;
  }

  // 9) 一括書き込み
  sheet.getRange(startRow, CONFIG_IMAGES.COL_S, numRows, 1).setValues(outS);
  sheet.getRange(startRow, CONFIG_IMAGES.COL_T, numRows, 1).setValues(outT);

  // 10) RichTextリンク化（URLがある行だけを速く処理）
  IMG_ST列をリンク化_範囲_(sheet, headerRows, startRow, numRows);

  IMGCV_uiSafeAlert_(
    '完了: ' + sheetName + '\n\n' +
    '更新行数: ' + hitRows + '\n' +
    '商品フォルダなし: ' + noFolder + '\n' +
    '免責ソース無し(カテゴリ不備など): ' + noDiscHit + '\n' +
    '免責コピー: ' + copyQueue.length + '\n\n' +
    '並び順: 商品画像（数字優先）→ 免責画像（最後・固定名順）'
  );
}


// ====================================================================
// Drive API（Advanced + fetchAll）
// ====================================================================

/** Drive: ファイルメタ取得 */
function IMG_drive_getFileMeta_(fileId) {
  try {
    var f = Drive.Files.get(fileId, { fields: 'id,name,mimeType' });
    return { id: f.id, name: f.name, mimeType: f.mimeType };
  } catch (e) {
    Logger.log('[DriveMetaErr] ' + e.message);
    return { id: String(fileId), name: '', mimeType: '' };
  }
}

/** Drive: 直下フォルダ map（name -> id） */
function IMG_drive_buildFolderMap_(parentId) {
  var map = {};
  var pageToken = null;

  // folder mime
  var q = "'" + parentId + "' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false";

  do {
    var res = Drive.Files.list({
      q: q,
      fields: 'nextPageToken, files(id,name)',
      pageSize: 1000,
      pageToken: pageToken,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true
    });

    if (res && res.files) {
      res.files.forEach(function(f) { map[f.name] = f.id; });
    }
    pageToken = res.nextPageToken;
  } while (pageToken);

  return map;
}

/**
 * Drive: 単一フォルダ内の画像ファイル一覧（全ページ）
 * - 並列が効かない“追撃用”として使う
 */
function IMG_drive_listImagesInFolder_allPages_(folderId) {
  var files = [];
  var pageToken = null;

  var q = "'" + folderId + "' in parents and trashed=false and mimeType contains 'image/'";

  do {
    var res = Drive.Files.list({
      q: q,
      fields: 'nextPageToken, files(id,name,mimeType)',
      pageSize: 1000,
      pageToken: pageToken,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true
    });

    if (res && res.files) files = files.concat(res.files);
    pageToken = res.nextPageToken;
  } while (pageToken);

  return files;
}

/**
 * Drive: 複数フォルダ内画像を “並列” で取得（基本1ページ目fetchAll）
 * - nextPageTokenが出たフォルダのみ全ページ追撃
 */
function IMG_drive_listFilesInFolders_parallel_(folderIds) {
  var out = {}; // folderId -> files[]

  if (!folderIds || folderIds.length === 0) return out;

  var token = ScriptApp.getOAuthToken();
  var batchSize = Math.max(1, Number(CONFIG_IMAGES.FETCHALL_BATCH || 80));

  // まず空配列で初期化
  for (var i = 0; i < folderIds.length; i++) out[folderIds[i]] = [];

  // fetchAll で1ページ目
  for (var s = 0; s < folderIds.length; s += batchSize) {
    var chunk = folderIds.slice(s, s + batchSize);

    var reqs = chunk.map(function(fid) {
      var q = "'" + fid + "' in parents and trashed=false and mimeType contains 'image/'";
      var url = 'https://www.googleapis.com/drive/v3/files'
        + '?q=' + encodeURIComponent(q)
        + '&fields=' + encodeURIComponent('nextPageToken,files(id,name,mimeType)')
        + '&pageSize=1000'
        + '&supportsAllDrives=true'
        + '&includeItemsFromAllDrives=true';

      return {
        url: url,
        method: 'get',
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      };
    });

    var resps = UrlFetchApp.fetchAll(reqs);

    for (var k = 0; k < resps.length; k++) {
      var fid2 = chunk[k];
      var resp = resps[k];
      var code = resp.getResponseCode();

      if (code >= 200 && code < 300) {
        var obj = {};
        try { obj = JSON.parse(resp.getContentText() || '{}'); } catch (e) { obj = {}; }
        var files = obj.files || [];
        out[fid2] = files;

        // nextPageToken があるフォルダだけ追撃（稀）
        if (obj.nextPageToken) {
          var more = IMG_drive_listImagesInFolder_allPages_(fid2);
          out[fid2] = more;
        }
      } else {
        // 429/5xxなどは追撃で救う
        Logger.log('[ListErr] folder=' + fid2 + ' code=' + code + ' body=' + (resp.getContentText() || '').slice(0, 200));
        out[fid2] = IMG_drive_listImagesInFolder_allPages_(fid2);
      }
    }
  }

  return out;
}

/**
 * Drive: Batch copy（Drive API batch）
 * queue item: {sourceId, targetFolderId, newName}
 *
 * - 成否を厳密に追わず、後で再スキャンして確定させる戦略
 */
function IMG_drive_executeBatchCopy_(queue) {
  if (!queue || queue.length === 0) return;

  var token = ScriptApp.getOAuthToken();
  var size = Math.max(1, Number(CONFIG_IMAGES.BATCH_COPY_SIZE || 80));

  for (var i = 0; i < queue.length; i += size) {
    var chunk = queue.slice(i, i + size);
    var boundary = 'batch_copy_' + Utilities.getUuid();

    var parts = [];

    for (var j = 0; j < chunk.length; j++) {
      var item = chunk[j];

      // fields=id,name を付けておく（レスポンス解析はしないが、デバッグ時に嬉しい）
      var path = '/drive/v3/files/' + encodeURIComponent(item.sourceId) + '/copy?fields=id,name';

      parts.push(
        '--' + boundary,
        'Content-Type: application/http',
        'Content-ID: <item-' + j + '>',
        '',
        'POST ' + path + ' HTTP/1.1',
        'Content-Type: application/json; charset=UTF-8',
        '',
        JSON.stringify({
          parents: [String(item.targetFolderId)],
          name: String(item.newName)
        }),
        ''
      );
    }

    parts.push('--' + boundary + '--', '');

    var options = {
      method: 'post',
      contentType: 'multipart/mixed; boundary=' + boundary,
      payload: parts.join('\r\n'),
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    };

    var resp = UrlFetchApp.fetch('https://www.googleapis.com/batch/drive/v3', options);
    var rc = resp.getResponseCode();
    if (!(rc >= 200 && rc < 300)) {
      Logger.log('[BatchCopyErr] code=' + rc + ' body=' + (resp.getContentText() || '').slice(0, 400));
    }
  }
}


// ====================================================================
// 免責ソース準備（カテゴリごとに一回）
// ====================================================================

/**
 * 免責ソースを準備して返す
 * return: { discType: [sourceFileId_for_01, sourceFileId_for_02, ...] }
 *
 * - 免責画像ルートIDが「親」でも「カテゴリ自身」でもOK
 * - カテゴリ内は “番号順” でソートして先頭を採用
 */
function IMG_disclaimer_prepareSources_(rootInfo, folderMapUnderRoot) {
  var out = {};

  var types = Object.keys(免責画像固定名);

  for (var i = 0; i < types.length; i++) {
    var discType = types[i];

    // カテゴリフォルダIDを解決
    var catFolderId = null;

    // ルート自身がカテゴリフォルダならそれを使う
    if (rootInfo && rootInfo.name && String(rootInfo.name) === String(discType)) {
      catFolderId = rootInfo.id;
    } else {
      // ルート直下にカテゴリがある想定
      catFolderId = folderMapUnderRoot[discType] || null;
    }

    if (!catFolderId) {
      out[discType] = [];
      continue;
    }

    // カテゴリ内の画像を全部取得 → 番号順に並べる
    var files = IMG_drive_listImagesInFolder_allPages_(catFolderId);

    // 番号ソート（①②/01/末尾_02 などを拾う）
    files.sort(function(a, b) {
      var na = IMG_先頭番号抽出_強化_(a.name);
      var nb = IMG_先頭番号抽出_強化_(b.name);
      if (na != null && nb != null) return na - nb;
      if (na != null) return -1;
      if (nb != null) return 1;
      return String(a.name || '').localeCompare(String(b.name || ''), 'ja');
    });

    // 必要枚数ぶんだけIDを取る（グッズは1、他は2…だが MAX_DISC も尊重）
    var need = IMG_免責カテゴリ_必要ファイル名_(discType).length;
    var limit = Math.min(need, Number(CONFIG_IMAGES.MAX_DISC || 99), files.length);

    var ids = [];
    for (var k = 0; k < limit; k++) ids.push(files[k].id);

    out[discType] = ids;
  }

  return out;
}

/** カテゴリごとの最終ファイル名（商品フォルダに入る固定名） */
function IMG_免責カテゴリ_必要ファイル名_(discType) {
  var arr = 免責画像固定名[discType];
  if (!arr || !arr.length) return [];
  // MAX_DISC が 1 のカテゴリもあるので念のため
  return arr.slice(0, Math.min(arr.length, Number(CONFIG_IMAGES.MAX_DISC || arr.length)));
}


// ====================================================================
// URL生成（ソート規約）
// ====================================================================

/**
 * files: Drive API files[] {id,name,mimeType}
 * discType: 判定された免責カテゴリ名（固定順ソートに使う）
 */
function IMG_generateSortedUrls_(files, discType) {
  var product = [];
  var disclaimer = [];

  for (var i = 0; i < files.length; i++) {
    var f = files[i];
    var mime = String(f.mimeType || '');
    if (mime.indexOf('image/') !== 0) continue;

    var name = String(f.name || '');
    var url = IMG_作成画像URL_(f.id);

    if (IMG_isDisclaimerName_(name)) {
      disclaimer.push({ name: name, url: url });
    } else {
      var num = IMG_画像順番抽出_(name);
      product.push({ name: name, num: num, url: url });
    }
  }

  // 商品画像：数字優先
  product.sort(function(a, b) {
    if (a.num != null && b.num != null) return a.num - b.num;
    if (a.num != null) return -1;
    if (b.num != null) return 1;
    return a.name.localeCompare(b.name, 'ja');
  });

  // 免責画像：固定名順（discTypeが分かれば）
  var fixed = (discType && 免責画像固定名[discType]) ? 免責画像固定名[discType] : [];
  if (fixed && fixed.length) {
    var order = {};
    fixed.forEach(function(fname, idx) { order[String(fname).toLowerCase()] = idx; });
    disclaimer.sort(function(a, b) {
      var ai = order[String(a.name).toLowerCase()];
      var bi = order[String(b.name).toLowerCase()];
      if (ai !== undefined && bi !== undefined) return ai - bi;
      if (ai !== undefined) return -1;
      if (bi !== undefined) return 1;
      return a.name.localeCompare(b.name, 'ja');
    });
  } else {
    disclaimer.sort(function(a, b) { return a.name.localeCompare(b.name, 'ja'); });
  }

  var urls = product.concat(disclaimer).map(function(x) { return x.url; });
  return urls.slice(0, Number(CONFIG_IMAGES.MAX_TOTAL || 21));
}


// ====================================================================
// S列/T列リンク化（RichText）
// ====================================================================

/**
 * S列/T列をリンク化（RichText）
 * @param {Sheet} sheet
 * @param {number} headerRows
 * @param {number} startRow
 * @param {number} numRows
 */
function IMG_ST列をリンク化_範囲_(sheet, headerRows, startRow, numRows) {
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow <= headerRows) return;

  var r0 = Math.max(headerRows + 1, startRow || (headerRows + 1));
  var n = Math.min(numRows || (lastRow - headerRows), lastRow - r0 + 1);
  if (n <= 0) return;

  var colS = CONFIG_IMAGES.COL_S;
  var colT = CONFIG_IMAGES.COL_T;

  var sRange = sheet.getRange(r0, colS, n, 1);
  var tRange = sheet.getRange(r0, colT, n, 1);

  // 書式設定（テキスト）
  sRange.setNumberFormat('@');
  tRange.setNumberFormat('@');

  var sVals = sRange.getValues();
  var tVals = tRange.getValues();

  var sRich = new Array(n);
  var tRich = new Array(n);

  for (var i = 0; i < n; i++) {
    // ---- S列（単一URL）----
    var sUrl = IMG_URL掃除_(sVals[i][0]);
    if (sUrl && /^https?:\/\//i.test(sUrl)) {
      sRich[i] = [SpreadsheetApp.newRichTextValue().setText(sUrl).setLinkUrl(sUrl).build()];
    } else {
      sRich[i] = [SpreadsheetApp.newRichTextValue().setText(String(sVals[i][0] || '')).build()];
    }

    // ---- T列（複数URL）----
    var urls = IMG_TセルをURL配列へ_(tVals[i][0]);

    if (!urls.length) {
      tRich[i] = [SpreadsheetApp.newRichTextValue().setText(String(tVals[i][0] || '')).build()];
      continue;
    }

    // 表示は改行で見やすく
    var display = urls.join('\n');
    var builder = SpreadsheetApp.newRichTextValue().setText(display);

    var pos = 0;
    for (var j = 0; j < urls.length; j++) {
      var u = urls[j];
      var start = pos;
      var end = pos + u.length;
      if (/^https?:\/\//i.test(u)) {
        builder = builder.setLinkUrl(start, end, u);
      }
      pos = end + 1; // 改行分
    }

    tRich[i] = [builder.build()];
  }

  sRange.setRichTextValues(sRich);
  tRange.setRichTextValues(tRich);
}


// ====================================================================
// ヘルパー（既存互換・強化）
// ====================================================================

/** ヘッダー名から列番号を探す */
function IMG_検索列位置_ヘッダー名_複数行_(sheet, headerRows, headerNames) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return -1;

  var rows = Math.max(1, Number(headerRows || 1));
  var vals = sheet.getRange(1, 1, rows, lastCol).getValues();

  var targets = headerNames.map(function(h) {
    return String(h).trim().toLowerCase();
  });

  for (var r = 0; r < vals.length; r++) {
    for (var c = 0; c < vals[r].length; c++) {
      var v = String(vals[r][c] || '').trim().toLowerCase();
      if (!v) continue;
      if (targets.indexOf(v) !== -1) return c + 1;
    }
  }
  return -1;
}

/** URL掃除（不可視文字除去） */
function IMG_URL掃除_(s) {
  s = String(s || '');
  s = s.replace(/[\u0000-\u001F\u007F]/g, '');
  s = s.replace(/[\u200B-\u200F\u202A-\u202E\u2060-\u206F\uFEFF]/g, '');
  return s.trim();
}

/** Tセル文字列 → URL配列（セミコロン/改行両対応） */
function IMG_TセルをURL配列へ_(tRaw) {
  var raw = IMG_URL掃除_(tRaw);
  if (!raw) return [];
  return raw
    .split(/[;\n]+/g)
    .map(function(x) { return IMG_URL掃除_(x); })
    .filter(Boolean);
}

/** DriveファイルID → 直リンク */
function IMG_作成画像URL_(fileId) {
  var mode = String(CONFIG_IMAGES.URL_EXPORT_MODE || 'view');
  if (mode !== 'view' && mode !== 'download') mode = 'view';
  return 'https://drive.google.com/uc?export=' + mode + '&id=' + String(fileId).trim();
}

/** 免責判定（商品名から免責カテゴリ名を決定） */
function IMG_判定免責フォルダ名_(name) {
  if (!name) return '';
  var t = String(name);

  // --- 言語判定 ---
  var lang = '';
  if (t.indexOf('台湾') !== -1)                          lang = 'tw';
  else if (t.indexOf('中国') !== -1 || t.indexOf('中国語') !== -1 || t.indexOf('中国版') !== -1) lang = 'cn';
  else if (t.indexOf('韓国') !== -1 || t.indexOf('韓国語') !== -1) lang = 'kr';
  else if (t.indexOf('英語') !== -1 || t.indexOf('English') !== -1) lang = 'en';

  // --- カテゴリ判定 ---
  var cat = '';
  if (t.indexOf('グッズ') !== -1)                                              cat = 'goods';
  else if (t.indexOf('雑誌') !== -1)                                           cat = 'magazine';
  else if (t.indexOf('まんが') !== -1 || t.indexOf('マンガ') !== -1 || t.indexOf('漫画') !== -1) cat = 'manga';
  else if (t.indexOf('小説') !== -1 || t.indexOf('書籍') !== -1 || t.indexOf('本') !== -1)      cat = 'book';

  // --- 言語×カテゴリ → 免責カテゴリ ---
  var map = {
    'tw:goods':    'グッズ免責画像',       // 台湾グッズ（分けるなら '台湾グッズ免責画像' に）
    'tw:magazine': '雑誌免責画像',
    'tw:manga':    '台湾免責画像',
    'tw:book':     '台湾免責画像',
    'tw:':         '台湾免責画像',         // カテゴリ不明でも台湾なら台湾免責

    'cn:goods':    '中国グッズ免責画像',
    'cn:magazine': '雑誌免責画像',
    'cn:manga':    '中国書籍免責画像',
    'cn:book':     '中国書籍免責画像',
    'cn:':         '中国書籍免責画像',

    'kr:goods':    'グッズ免責画像',
    'kr:magazine': '雑誌免責画像',
    'kr:manga':    '',                     // 後段の版型判定に委譲
    'kr:book':     '韓国マンガ一般版免責画像',
    'kr:':         '',                     // 後段に委譲

    'en:goods':    'グッズ免責画像',
    'en:magazine': '雑誌免責画像',
    'en:':         '',

    ':magazine':   '雑誌免責画像',         // 言語不明でも雑誌なら
    ':goods':      'グッズ免責画像',
  };

  var key = lang + ':' + cat;
  if (map[key] !== undefined && map[key] !== '') return map[key];

  // --- 韓国マンガの版型判定（既存ロジック） ---
  if (t.indexOf('初版限定') !== -1 || t.indexOf('初版') !== -1) return '韓国マンガ初版限定免責画像';
  if (t.indexOf('特装版') !== -1)  return '韓国マンガ特装版免責画像';
  if (t.indexOf('限定版') !== -1)  return '韓国マンガ限定版免責画像';
  if (t.indexOf('一般') !== -1 || t.indexOf('通常版') !== -1) return '韓国マンガ一般版免責画像';
  if (lang === 'kr') return '韓国マンガ一般版免責画像';

  return '';
}

/** 画像順番抽出（数字優先ソート用） */
function IMG_画像順番抽出_(name) {
  var s = String(name || '');

  var m = s.match(/^(\d+)/);
  if (m) return Number(m[1]);

  m = s.match(/[_-](\d+)(?=\.[^.]+$)/);
  if (m) return Number(m[1]);

  m = s.match(/(\d+)(?=\.[^.]+$)/);
  if (m) return Number(m[1]);

  return null;
}

/** 免責ファイル判定（ファイル名） */
function IMG_isDisclaimerName_(fileName) {
  var n = String(fileName || '').toLowerCase();
  return n.indexOf('notice_') === 0 || n.indexOf('免責') !== -1;
}

/**
 * 先頭番号抽出 強化
 * 対応例:
 *  - "①台湾注意.jpg" -> 1
 *  - "② ...png" -> 2
 *  - "1_....jpg" -> 1
 *  - "01-....jpg" -> 1
 *  - "notice_xxx_02.jpg" -> 2（末尾も見る）
 */
function IMG_先頭番号抽出_強化_(name) {
  var s = String(name || '');

  // ①②…（丸数字）
  var m = s.match(/^([①②③④⑤⑥⑦⑧⑨⑩])/);
  if (m) {
    var map = {'①':1,'②':2,'③':3,'④':4,'⑤':5,'⑥':6,'⑦':7,'⑧':8,'⑨':9,'⑩':10};
    return map[m[1]] || null;
  }

  // 先頭の数字
  m = s.match(/^(\d{1,2})/);
  if (m) return Number(m[1]);

  // _01 / -02 など（拡張子直前）
  m = s.match(/[_-](\d{1,2})(?=\.[^.]+$)/);
  if (m) return Number(m[1]);

  // 末尾の数字（拡張子直前）
  m = s.match(/(\d{1,2})(?=\.[^.]+$)/);
  if (m) return Number(m[1]);

  return null;
}


// ====================================================================
// デバッグ用（必要なら）
// ====================================================================
function DBG_設定確認() {
  Logger.log(JSON.stringify(CONFIG_IMAGES, null, 2));
}

// =====================================================
// DBG: フォルダ＆中身が本当に取れてるか（先頭10件）
// =====================================================
function DBG_商品フォルダ認識チェック() { // ←末尾の _ を削除
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_IMAGES.シート名_商品入力);
  if (!sheet) { Logger.log('no sheet'); return; }

  var headerRows = CONFIG_IMAGES.ヘッダー行数_商品入力;
  var lastRow = sheet.getLastRow();
  var startRow = headerRows + 1;
  var numRows = Math.min(50, lastRow - headerRows);
  if (numRows <= 0) { Logger.log('no rows'); return; }

  var needCols = Math.max(CONFIG_IMAGES.COL_CODE, CONFIG_IMAGES.COL_NAME);
  var v = sheet.getRange(startRow, 1, numRows, needCols).getValues();

  var codes = v.map(r => String(r[CONFIG_IMAGES.COL_CODE - 1] || '').trim());
  var names = v.map(r => String(r[CONFIG_IMAGES.COL_NAME - 1] || '').trim());

  // 商品フォルダmap
  var map = IMG_drive_buildFolderMap_(CONFIG_IMAGES.商品画像ルートID);

  for (var i = 0; i < Math.min(10, codes.length); i++) {
    var code = codes[i];
    if (!code) continue;

    var fid = map[code] || '';
    Logger.log('--- row=' + (startRow + i) + ' code=' + code + ' folderId=' + fid);

    if (!fid) continue;

    // 中身（画像）を1回だけ
    var files = IMG_drive_listImagesInFolder_allPages_(fid);
    Logger.log('    imageCount=' + files.length);
    Logger.log('    sample=' + files.slice(0, 5).map(x => x.name).join(' | '));
  }

  // 免責判定も一緒に見せる
  for (var j = 0; j < Math.min(10, codes.length); j++) {
    var t = IMG_判定免責フォルダ名_(names[j]);
    Logger.log('discJudge row=' + (startRow + j) + ' name=' + names[j] + ' => ' + t);
  }
}

// =====================================================
// DBG: 免責ソース（カテゴリ）側が本当に取れてるか
// =====================================================
function DBG_免責ソース確認() { // ←末尾の _ を削除
  var rootInfo = IMG_drive_getFileMeta_(CONFIG_IMAGES.免責画像ルートID);
  var under = IMG_drive_buildFolderMap_(CONFIG_IMAGES.免責画像ルートID);
  Logger.log('rootInfo=' + JSON.stringify(rootInfo));

  var src = IMG_disclaimer_prepareSources_(rootInfo, under);

  Object.keys(免責画像固定名).forEach(function(type) {
    Logger.log('type=' + type + ' sourceIds=' + JSON.stringify(src[type] || []));
    Logger.log('requiredNames=' + JSON.stringify(IMG_免責カテゴリ_必要ファイル名_(type)));
  });
}

// =====================================================
// DBG: S/T に入ったURLが “見れるURL” か（権限/形式の簡易チェック）
//  - 先頭3件だけ HEAD で確認（重いので最小）
// =====================================================
function DBG_URL到達チェック() { // ←末尾の _ を削除
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_IMAGES.シート名_商品入力);
  if (!sheet) return;

  var headerRows = CONFIG_IMAGES.ヘッダー行数_商品入力;
  var startRow = headerRows + 1;
  var numRows = Math.min(3, sheet.getLastRow() - headerRows);
  if (numRows <= 0) return;

  var sVals = sheet.getRange(startRow, CONFIG_IMAGES.COL_S, numRows, 1).getValues().flat();
  sVals.forEach(function(u, idx) {
    u = String(u || '').trim();
    if (!u) return;
    try {
      var resp = UrlFetchApp.fetch(u, { method: 'get', followRedirects: true, muteHttpExceptions: true });
      Logger.log('row=' + (startRow + idx) + ' code=' + resp.getResponseCode() + ' url=' + u);
      Logger.log('contentType=' + resp.getHeaders()['Content-Type']);
      Logger.log('bodyHead=' + (resp.getContentText() || '').slice(0, 120));
    } catch (e) {
      Logger.log('row=' + (startRow + idx) + ' fetchErr=' + e.message);
    }
  });
}
function DBG_ヘッダー確認() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_IMAGES.シート名_商品入力);
  var lastCol = sh.getLastColumn();
  var rows = CONFIG_IMAGES.ヘッダー行数_商品入力; // 2 なら 1-2行を見る
  var hv = sh.getRange(1, 1, rows, lastCol).getValues();

  for (var r = 0; r < hv.length; r++) {
    var line = [];
    for (var c = 0; c < hv[r].length; c++) {
      var v = String(hv[r][c] || '').trim();
      if (v) line.push((c+1) + ':' + v);
    }
    Logger.log('HEADER row=' + (r+1) + ' => ' + line.join(' | '));
  }
}
