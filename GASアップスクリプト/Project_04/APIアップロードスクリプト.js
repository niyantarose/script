/******************************************************
 * 10_shipping_master.gs - 配送マスタ読込 & 配送テキスト反映
 * v3: 重量(K列)自動設定 + 管理番号(R列)変換 追加
 * 
 * ✅ v3の変更点:
 * - 配送グループ設定シートのD列（管理番号）を読み込み
 * - R列(postage-set)を旧番号→管理番号(D列)に変換
 * - K列(ship-weight)を配送種別に応じて自動設定
 *   佐川 → 1000、ゆうパケ系 → 100
 ******************************************************/

// ====================================================================
// ★ 安全なUI通知（UIが無くても落ちない）
// ====================================================================

function SHIP_uiSafeAlert_(msg) {
  try {
    SpreadsheetApp.getUi().alert(String(msg));
    return;
  } catch (e) {}

  // フォールバック: toast → log
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(String(msg), '配送マスタ', 10);
  } catch (e2) {}

  Logger.log('[SHIP_UI_FALLBACK] ' + msg);
}


/**
 * 配送マスタ読み込み
 * - 配送グループ設定: postage-set(A列) → type(B列), postage-set(A列) → 管理番号(D列)
 * - 配送種別テキスト: type → txt（Driveテキスト）
 */
function 配送マスタを読み込み_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const groupSheet = ss.getSheetByName(SHEET_配送グループ設定);
  const textSheet  = ss.getSheetByName(SHEET_配送種別テキスト);

  if (!groupSheet || !textSheet) {
    SHIP_uiSafeAlert_('配送グループ設定 / 配送種別テキスト シートが見つかりません。');
    throw new Error('配送マスタシートが見つからない');
  }

  Logger.log('=== 配送マスタ読み込み 開始 ===');

  // ① postage-set(A) -> type(B) AND postage-set(A) -> 管理番号(D)
  const groupLastRow = groupSheet.getLastRow();
  /* ✅ v3: 4列読み込み（A:postage-set, B:配送種別, C:備考, D:管理番号） */
  const groupValues = groupLastRow >= 2
    ? groupSheet.getRange(2, 1, groupLastRow - 1, 4).getValues()
    : [];

  const postageToType = {};   // A列 → B列（配送種別）
  const postageToNum  = {};   // A列 → D列（管理番号）
  groupValues.forEach(row => {
    const postageSet = String(row[0] || '').trim();   // A列
    const type       = String(row[1] || '').trim();   // B列
    const num        = row[3];                         // D列（管理番号）
    if (postageSet && type) postageToType[postageSet] = type;
    if (postageSet && num !== '' && num !== null && num !== undefined) {
      postageToNum[postageSet] = num;
    }
  });
  Logger.log('postageToType 件数: ' + Object.keys(postageToType).length);
  Logger.log('postageToNum 件数: '  + Object.keys(postageToNum).length);

  // ② type -> text(Drive)
  const textLastRow = textSheet.getLastRow();
  const textRange = textLastRow >= 2
    ? textSheet.getRange(2, 1, textLastRow - 1, 2)
    : null;

  const typeToText = {};
  if (textRange) {
    const values = textRange.getValues();
    const rich   = textRange.getRichTextValues();

    for (let i = 0; i < values.length; i++) {
      const type = String(values[i][0] || '').trim();
      if (!type) continue;

      let url = rich[i][1]?.getLinkUrl();
      if (!url) url = String(values[i][1] || '').trim();
      if (!url) continue;

      let fileId = url.match(/[-\w]{25,}/)?.[0];

      if (!fileId) {
        const files = DriveApp.getFilesByName(url);
        if (files.hasNext()) fileId = files.next().getId();
      }
      if (!fileId) continue;

      try {
        const txt = DriveApp.getFileById(fileId).getBlob().getDataAsString('UTF-8');
        typeToText[type] = txt;
      } catch (e) {
        Logger.log('テキスト取得失敗: type=' + type + ' error=' + e);
      }
    }
  }

  Logger.log('typeToText 件数: ' + Object.keys(typeToText).length);
  Logger.log('=== 配送マスタ読み込み 終了 ===');

  return { postageToType, typeToText, postageToNum };
}

// ====== エントリポイント ======
function 配送テキストを反映_商品入力シート対象() {
  配送テキストを反映_汎用_(SHEET_商品入力, HEADER_商品入力);
}
function 配送テキストを反映_Yahoo商品登録シート対象() {
  配送テキストを反映_汎用_(SHEET_Yahoo商品登録, HEADER_Yahoo商品登録);
}

/**
 * 配送テキスト反映 共通
 * ✅ v3: 重量(K列) + 管理番号(R列) の自動設定を追加
 * 
 * 処理フロー:
 *   R列の旧postage-set番号 → 配送グループ設定で照合
 *   → K列: 佐川=1000, ゆうパケ系=100
 *   → R列: D列の管理番号に変換
 *   → I列/Q列: 配送テキスト反映
 */
function 配送テキストを反映_汎用_(sheetName, headerRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    SHIP_uiSafeAlert_('シートが見つかりません: ' + sheetName);
    return;
  }

  const { postageToType, typeToText, postageToNum } = 配送マスタを読み込み_();

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= headerRows) {
    SHIP_uiSafeAlert_('データ行がありません: ' + sheetName);
    return;
  }

  const range = sheet.getRange(headerRows + 1, 1, lastRow - headerRows, lastCol);
  const data = range.getValues();

  const idxI = COL_ABSTRACT - 1;
  const idxQ = COL_SP_ADDITIONAL - 1;
  const idxR = COL_POSTAGE_SET - 1;
  const idxK = COL_SHIP_WEIGHT - 1;

  let updateCount = 0;
  let weightUpdate = 0;
  let numUpdate = 0;

  for (let r = 0; r < data.length; r++) {
    const postageSet = String(data[r][idxR] || '').trim();
    if (!postageSet) continue;

    const type = postageToType[postageSet];
    if (!type) continue;

    // --- 配送テキスト反映（従来通り） ---
    const txt = typeToText[type];
    if (txt) {
      data[r][idxI] = txt;
      data[r][idxQ] = txt;
      updateCount++;
    }

    // --- ✅ v3: 重量(K列)の自動設定 ---
    if (type === '佐川') {
      data[r][idxK] = 1000;
      weightUpdate++;
    } else if (type.startsWith('ゆうパケ')) {
      // ゆうパケ, ゆうパケ_1, ゆうパケ_2, ゆうパケ_3 全て100
      data[r][idxK] = 100;
      weightUpdate++;
    }

    // --- ✅ v3: R列を管理番号(D列の値)に変換 ---
    const num = postageToNum[postageSet];
    if (num !== undefined && num !== '') {
      data[r][idxR] = Number(num);
      numUpdate++;
    }
  }

  range.setValues(data);

  SHIP_uiSafeAlert_(
    '配送テキスト反映完了\n' +
    'テキスト更新行数: ' + updateCount + '\n' +
    '重量変更: ' + weightUpdate + '\n' +
    '管理番号変換: ' + numUpdate
  );
}/******************************************************
 * Yahoo画像アップロード（VPS高速版）完全版 v2.0
 *
 * ✅ VPSで並列処理（GASは軽い）
 * ✅ 6分制限対策：自動継続（1分後トリガー）
 * ✅ 既存URL行でカーソルが進まない致命バグを修正
 * ✅ 自動継続でも「再アップロード（削除あり）」モードを維持
 * ✅ 削除→反映→待機→アップロード→反映 の順で安定化
 ******************************************************/

/******************************************************
 * Yahoo画像アップロード（VPS高速版）完全版 v2
 *
 * ✅ 改善点
 * - itemImageList の query を安全化（im-01004対策）
 * - URLからimage_id抽出時に拡張子/クエリ除去＋形式チェック（im-04005対策）
 * - VPSが部分成功(ok<総数)のときはデフォルトでエラー扱い（画像欠け防止）
 ******************************************************/

var YahooVPS = YahooVPS || {};

YahooVPS.設定 = {
  // ========== シート設定 ==========
  シート名: '①商品入力シート',
  ヘッダー行数: 2,
  列_コード: 3,       // C
  列_メイン画像: 19,  // S
  列_詳細画像: 20,    // T

  // ========== VPS設定 ==========
  VPS_UPLOAD_URL: 'https://img.niyantarose.com/yahoo/upload_images.php',

  // ========== Drive設定 ==========
  画像ルートフォルダID: '1Vk1cVgrXM5CNFqfBvAaBUJpheDn4f9Lz',

  // ========== 処理設定 ==========
  最大画像枚数: 21,
  最大詳細枚数: 20,
  '1回の最大処理行数': 30,

  // VPSが「一部だけ成功」の時にどうするか
  許容_部分成功: false, // trueにすると ok分だけ反映して進める

  // ========== 免責キーワード ==========
  免責キーワード: ['notice_', '免責', 'menseki', 'disclaimer'],
};

/** =====================================================
 * メニュー
 * ===================================================== */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 Yahoo画像(VPS高速)')
    .addItem('① 初期設定（client/secret/refresh/seller）', 'YahooVPS_初期設定')
    .addItem('② refresh_token 生存確認', 'YahooVPS_トークン確認')
    .addSeparator()
    .addItem('③ 全行アップロード', 'YahooVPS_全行アップロード')
    .addItem('④ 再アップロード（既存削除→再登録）', 'YahooVPS_再アップロード')
    .addSeparator()
    .addItem('⏹ 処理停止', 'YahooVPS_停止')
    .addItem('📊 進捗確認', 'YahooVPS_進捗確認')
    .addItem('🔄 最初からやり直し', 'YahooVPS_リセット')
    .addSeparator()
    .addItem('🧪 デバッグ（設定確認）', 'YahooVPS_デバッグ')
    .addToUi();
}

/** =====================================================
 * 初期設定
 * ===================================================== */
function YahooVPS_初期設定() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();

  var clientId = ui.prompt('Yahoo client_id を入力').getResponseText().trim();
  var clientSecret = ui.prompt('Yahoo client_secret を入力').getResponseText().trim();
  var refreshToken = ui.prompt('Yahoo refresh_token を入力').getResponseText().trim();
  var sellerId = ui.prompt('Yahoo seller_id（例: niyantarose）を入力').getResponseText().trim();

  if (!clientId || !clientSecret || !refreshToken || !sellerId) {
    ui.alert('未入力があるため中断');
    return;
  }

  props.setProperties({
    Y_CLIENT_ID: clientId,
    Y_CLIENT_SECRET: clientSecret,
    Y_REFRESH_TOKEN: refreshToken,
    Y_SELLER_ID: sellerId
  });

  ui.alert('✅ 保存完了');
}

function YahooVPS_トークン確認() {
  try {
    var token = YahooVPS__getAccessToken_(true);
    SpreadsheetApp.getUi().alert('✅ トークン有効\n先頭: ' + token.slice(0, 12) + '...');
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ トークンエラー\n' + (e.message || e));
  }
}

/** =====================================================
 * 入口関数
 * ===================================================== */
function YahooVPS_全行アップロード() {
  YahooVPS__メイン処理_({
    既存画像削除: false,
    確認メッセージ: '既にYahoo URLがある行はsubmitのみ実行'
  });
}

function YahooVPS_再アップロード() {
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert(
    '再アップロード確認',
    '⚠️ 既存画像を全削除してから再登録します\n\n' +
    '1. Yahoo上の画像を削除\n' +
    '2. 削除を反映\n' +
    '3. 新しい画像をアップロード\n' +
    '4. 反映\n\n' +
    '実行しますか？',
    ui.ButtonSet.OK_CANCEL
  );

  if (confirm !== ui.Button.OK) return;

  YahooVPS__メイン処理_({
    既存画像削除: true,
    確認メッセージ: '既存画像削除 → 新規登録'
  });
}

function YahooVPS_停止() {
  YahooVPS__deleteTrigger_();
  SpreadsheetApp.getUi().alert('⏹ 自動継続を停止しました');
}

function YahooVPS_進捗確認() {
  var cfg = YahooVPS.設定;
  var props = PropertiesService.getScriptProperties();
  var cursor = Number(props.getProperty('YAHOOVPS_CURSOR') || '0');

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.シート名);
  var total = sheet ? sheet.getLastRow() - cfg.ヘッダー行数 : 0;
  var pct = total > 0 ? ((cursor / total) * 100).toFixed(1) : 0;
  var hasTrigger = YahooVPS__hasTrigger_();

  SpreadsheetApp.getUi().alert(
    '【進捗状況】\n' +
    '総行数: ' + total + '\n' +
    '処理済み: ' + cursor + '\n' +
    '進捗: ' + pct + '%\n' +
    '自動継続: ' + (hasTrigger ? '実行中' : '停止中')
  );
}

function YahooVPS_リセット() {
  PropertiesService.getScriptProperties().deleteProperty('YAHOOVPS_CURSOR');
  YahooVPS__deleteTrigger_();
  SpreadsheetApp.getUi().alert('🔄 カーソルをリセットしました');
}

/** 自動継続用（トリガーから呼ばれる） */
function YahooVPS_自動継続() {
  YahooVPS__メイン処理_({ 既存画像削除: false, 自動実行: true });
}

/** =====================================================
 * メイン処理
 * ===================================================== */
function YahooVPS__メイン処理_(opt) {
  var cfg = YahooVPS.設定;
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var isAuto = opt.自動実行 || false;

  var MAX_PER_RUN = Number(cfg['1回の最大処理行数']) || 30;
  var MAX_TIME_MS = 5 * 60 * 1000;
  var SAFE_MARGIN = 45 * 1000;

  var t0 = Date.now();
  var safeEnd = t0 + MAX_TIME_MS - SAFE_MARGIN;

  // ロック
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    if (!isAuto) ui.alert('別の処理が実行中です');
    return;
  }

  try {
    cfg.VPS_UPLOAD_URL = YahooVPS__normalizeUrl_(cfg.VPS_UPLOAD_URL);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(cfg.シート名);
    if (!sheet) throw new Error('シートが見つかりません: ' + cfg.シート名);

    var lastRow = sheet.getLastRow();
    if (lastRow <= cfg.ヘッダー行数) {
      if (!isAuto) ui.alert('データ行がありません');
      return;
    }

    var numRows = lastRow - cfg.ヘッダー行数;
    var cursor = Number(props.getProperty('YAHOOVPS_CURSOR') || '0');
    if (cursor >= numRows) cursor = 0;

    // 確認ダイアログ（手動のみ）
    if (!isAuto) {
      var msg =
        '処理対象: ' + numRows + ' 行\n' +
        '開始位置: ' + (cfg.ヘッダー行数 + 1 + cursor) + ' 行目\n' +
        '1回の処理: 最大 ' + MAX_PER_RUN + ' 行\n\n' +
        (opt.確認メッセージ || '') + '\n\n実行しますか？';

      if (ui.alert('確認', msg, ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
    }

    // データ読み込み
    var codes = sheet.getRange(cfg.ヘッダー行数 + 1, cfg.列_コード, numRows, 1).getValues();
    var sVals = sheet.getRange(cfg.ヘッダー行数 + 1, cfg.列_メイン画像, numRows, 1).getValues();
    var tVals = sheet.getRange(cfg.ヘッダー行数 + 1, cfg.列_詳細画像, numRows, 1).getValues();

    var sellerId = YahooVPS__getProp_('Y_SELLER_ID');
    var rootFolder = DriveApp.getFolderById(cfg.画像ルートフォルダID);

    Logger.log('🚀 VPS高速版開始 cursor=' + cursor + ' MAX=' + MAX_PER_RUN + ' del=' + !!opt.既存画像削除);

    var processed = 0, okNew = 0, okExisting = 0, skipped = 0, errCount = 0;
    var errors = [];
    var endIdx = Math.min(cursor + MAX_PER_RUN, numRows);

    for (var r = cursor; r < endIdx; r++) {
      if (Date.now() > safeEnd) {
        Logger.log('⏰ 時間切れ r=' + r);
        break;
      }

      var code = String(codes[r][0] || '').trim();
      var rowNum = cfg.ヘッダー行数 + 1 + r;
      if (!code) continue;

      processed++;
      Logger.log('--- 行 ' + rowNum + ' [' + code + '] ---');

      var existingS = YahooVPS__cleanCellUrl_(sVals[r][0]);
      var existingT = YahooVPS__cleanCellUrl_(tVals[r][0]);
      var hasYahooS = existingS && YahooVPS__isYahooUrl_(existingS);

      try {
        // 既存URLあり & 削除なし → submitのみ
        if (hasYahooS && !opt.既存画像削除) {
          Logger.log('  👉 既存URL、submitのみ');
          YahooVPS__submitItem_(sellerId, code);
          okExisting++;
          props.setProperty('YAHOOVPS_CURSOR', String(r + 1));
          continue;
        }

        // 既存画像削除
        if (opt.既存画像削除) {
          Logger.log('  🗑️ 既存画像削除');
          YahooVPS__deleteExistingImages_(sellerId, code, existingS, existingT);
        }

        // 画像ソース収集
        var sources = YahooVPS__collectSources_(rootFolder, code, existingS, existingT);
        if (!sources.length) {
          skipped++;
          Logger.log('  ⚠️ 画像なし（Driveフォルダ未発見 or URL空）');
          props.setProperty('YAHOOVPS_CURSOR', String(r + 1));
          continue;
        }

        sources = sources.slice(0, cfg.最大画像枚数);
        Logger.log('  📦 画像: ' + sources.length + '枚');

        // VPSでアップロード
        var t1 = Date.now();
        var yahooUrls = YahooVPS__uploadViaVPS_(sellerId, code, sources);
        var uploadTime = ((Date.now() - t1) / 1000).toFixed(1);
        Logger.log('  ✅ VPSアップロード完了 (' + uploadTime + '秒)');

        // シートに書き込み
        sVals[r][0] = yahooUrls[0] || '';
        tVals[r][0] = yahooUrls.slice(1, 1 + cfg.最大詳細枚数).join(';');

        // 反映
        YahooVPS__submitItem_(sellerId, code);
        okNew++;
        Logger.log('  🎉 完了 メイン1 + 詳細' + (yahooUrls.length - 1));

      } catch (e) {
        errCount++;
        var errMsg = (e && e.message) ? e.message : String(e);
        errors.push(code + '(行' + rowNum + '): ' + errMsg);
        Logger.log('  ❌ ERROR: ' + errMsg);
      }

      props.setProperty('YAHOOVPS_CURSOR', String(r + 1));
    }

    // 書き戻し
    sheet.getRange(cfg.ヘッダー行数 + 1, cfg.列_メイン画像, numRows, 1).setValues(sVals);
    sheet.getRange(cfg.ヘッダー行数 + 1, cfg.列_詳細画像, numRows, 1).setValues(tVals);

    var newCursor = Number(props.getProperty('YAHOOVPS_CURSOR') || '0');
    var remaining = numRows - newCursor;
    var elapsed = ((Date.now() - t0) / 1000).toFixed(1);

    var result =
      '=== 完了 ===\n' +
      '処理: ' + processed + ' 行\n' +
      '成功(既存): ' + okExisting + '\n' +
      '成功(新規): ' + okNew + '\n' +
      'スキップ: ' + skipped + '\n' +
      'エラー: ' + errCount + '\n' +
      '時間: ' + elapsed + '秒\n' +
      '残り: ' + remaining + ' 行';

    if (errors.length > 0) result += '\n\n--- エラー（先頭5件）---\n' + errors.slice(0, 5).join('\n');

    Logger.log(result.replace(/\n/g, ' | '));

    if (remaining > 0) {
      YahooVPS__ensureTrigger_();
      if (!isAuto) ui.alert(result + '\n\n⏳ 1分後に自動継続します');
    } else {
      props.deleteProperty('YAHOOVPS_CURSOR');
      YahooVPS__deleteTrigger_();
      if (!isAuto) ui.alert(result + '\n\n🏁 全行完了！');
    }

  } finally {
    lock.releaseLock();
  }
}

/** =====================================================
 * VPS経由アップロード
 * ===================================================== */
function YahooVPS__uploadViaVPS_(sellerId, code, sources) {
  var cfg = YahooVPS.設定;
  cfg.VPS_UPLOAD_URL = YahooVPS__normalizeUrl_(cfg.VPS_UPLOAD_URL);

  var accessToken = YahooVPS__getAccessToken_();

  // 画像リスト作成
  var images = [];
  var detailIdx = 0;

  for (var i = 0; i < sources.length; i++) {
    var src = sources[i];
    var name = (i === 0) ? (code + '.jpg') : (code + '_' + (++detailIdx) + '.jpg');

    var imgData = YahooVPS__getImageData_(src);
    if (imgData) {
      imgData.name = name;
      images.push(imgData);
    }
  }

  if (images.length === 0) throw new Error('アップロード可能な画像がありません');

  Logger.log('    VPSに送信: ' + images.length + '枚');
  Logger.log('    VPS_URL: ' + cfg.VPS_UPLOAD_URL);

  var payload = { seller_id: sellerId, access_token: accessToken, images: images };

  var resp = UrlFetchApp.fetch(cfg.VPS_UPLOAD_URL, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var httpCode = resp.getResponseCode();
  var body = resp.getContentText() || '';

  if (httpCode !== 200) {
    // トークン期限切れっぽい場合だけ更新して再試行
    if (httpCode === 401 || body.toLowerCase().indexOf('expired') !== -1) {
      Logger.log('    トークン更新して再試行');
      accessToken = YahooVPS__getAccessToken_(true);
      payload.access_token = accessToken;

      resp = UrlFetchApp.fetch(cfg.VPS_UPLOAD_URL, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      httpCode = resp.getResponseCode();
      body = resp.getContentText() || '';
    }
    if (httpCode !== 200) throw new Error('VPSエラー HTTP=' + httpCode + ' ' + body.slice(0, 300));
  }

  var result;
  try {
    result = JSON.parse(body);
  } catch (e) {
    throw new Error('VPS応答JSON解析失敗: ' + body.slice(0, 300));
  }

  if (!result.success) throw new Error('VPSエラー: ' + (result.error || 'unknown'));

  var ok = (result.stats && result.stats.ok) ? Number(result.stats.ok) : 0;
  var ng = (result.stats && result.stats.ng) ? Number(result.stats.ng) : 0;

  Logger.log('    VPS結果: ok=' + ok + ' ng=' + ng + ' time=' + (result.stats ? result.stats.time_sec : '?') + 's');

  // 成功URL抽出（順序維持）
  var yahooUrls = [];
  (result.results || []).forEach(function(r) {
    if (r && r.ok && r.yahoo_url) yahooUrls.push(String(r.yahoo_url).trim());
  });

  if (!yahooUrls.length) throw new Error('アップロード成功が0件');

  // ★部分成功を失敗扱い（画像欠け防止）
  if (!cfg.許容_部分成功 && yahooUrls.length < images.length) {
    var ngList = (result.results || [])
      .filter(function(x){ return x && !x.ok; })
      .slice(0, 5)
      .map(function(x){ return (x.name || '?') + ':' + (x.error || 'NG'); })
      .join(' / ');
    throw new Error('VPS部分成功 ok=' + yahooUrls.length + '/' + images.length + ' NG例=' + ngList);
  }

  return yahooUrls;
}

/**
 * 画像ソースからデータを取得（URL形式で統一）
 */
function YahooVPS__getImageData_(src) {
  if (!src) return null;

  if (src.種類 === 'url') {
    var u = YahooVPS__normalizeUrl_(src.url);
    if (!/^https?:\/\//i.test(u)) return null;
    return { url: u };
  }

  if (src.種類 === 'drive') {
    try {
      var file = src.ファイル;
      var fileId = file.getId();

      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (e) {
        Logger.log('    共有設定スキップ: ' + e);
      }

      var url = 'https://drive.google.com/uc?export=download&id=' + encodeURIComponent(fileId);

      Logger.log('    [DRIVE] name=' + file.getName() + ' id=' + fileId);
      Logger.log('    [DRIVE] url=' + url);

      return { url: url };

    } catch (e) {
      Logger.log('    Drive URL取得エラー: ' + e);
      return null;
    }
  }

  return null;
}

/** =====================================================
 * 画像ソース収集
 * ===================================================== */
function YahooVPS__collectSources_(rootFolder, code, sUrl, tUrl) {
  var driveList = YahooVPS__getDriveImages_(rootFolder, code);
  if (driveList.length) return driveList;

  var urls = [];
  if (sUrl && !YahooVPS__isYahooUrl_(sUrl)) urls.push(sUrl);

  if (tUrl) {
    tUrl.split(';').forEach(function(u) {
      u = YahooVPS__cleanCellUrl_(u);
      if (u && !YahooVPS__isYahooUrl_(u)) urls.push(u);
    });
  }

  // 重複除去（正規化後）
  var seen = {};
  urls = urls
    .map(YahooVPS__normalizeUrl_)
    .filter(function(u) {
      if (!u) return false;
      if (seen[u]) return false;
      seen[u] = true;
      return true;
    });

  // 免責を後ろへ
  var cfg = YahooVPS.設定;
  var normal = [], disc = [];
  urls.forEach(function(u) {
    var lu = u.toLowerCase();
    var isDisc = cfg.免責キーワード.some(function(k) {
      return lu.indexOf(String(k).toLowerCase()) !== -1;
    });
    if (isDisc) disc.push(u); else normal.push(u);
  });

  return normal.concat(disc).map(function(u){ return { 種類:'url', url:u }; });
}

function YahooVPS__getDriveImages_(rootFolder, code) {
  var it = rootFolder.getFoldersByName(code);
  if (!it.hasNext()) return [];

  var folder = it.next();
  var fit = folder.getFiles();
  var cfg = YahooVPS.設定;

  var product = [], disc = [];
  var seen = {};

  while (fit.hasNext()) {
    var f = fit.next();
    var name = String(f.getName() || '').trim();
    if (!/\.(jpg|jpeg|png|webp|gif)$/i.test(name)) continue;
    if (seen[name]) continue;
    seen[name] = true;

    var isDisc = cfg.免責キーワード.some(function(k) {
      return name.toLowerCase().indexOf(String(k).toLowerCase()) !== -1;
    });

    var entry = {
      種類: 'drive',
      ファイル: f,
      表示名: name,
      番号: YahooVPS__getFileNumber_(name, code),
      免責フラグ: isDisc
    };

    if (isDisc) disc.push(entry); else product.push(entry);
  }

  product.sort(function(a, b) {
    var an = (a.番号 == null) ? 999999 : a.番号;
    var bn = (b.番号 == null) ? 999999 : b.番号;
    return an - bn || String(a.表示名).localeCompare(String(b.表示名), 'ja');
  });
  disc.sort(function(a, b) {
    return String(a.表示名).localeCompare(String(b.表示名), 'ja');
  });

  return product.concat(disc);
}

function YahooVPS__getFileNumber_(name, code) {
  var m = name.match(/^(\d+)\./);
  if (m) return Number(m[1]);

  var re = new RegExp('^' + code.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '[_-](\\d+)\\.', 'i');
  m = name.match(re);
  if (m) return Number(m[1]);

  m = name.match(/(\d+)\.[^.]+$/);
  if (m) return Number(m[1]);

  return null;
}

/** =====================================================
 * 既存画像削除
 * ===================================================== */
function YahooVPS__deleteExistingImages_(sellerId, code, existingS, existingT) {
  var ids = [];

  // 1) itemImageListで取得（query安全化）
  try {
    ids = YahooVPS__getImageIds_(sellerId, code);
  } catch (e) {
    Logger.log('    itemImageList失敗（例外）: ' + e);
    ids = [];
  }

  // 2) フォールバック：URLから抽出
  if (!ids.length) {
    ids = YahooVPS__extractIdsFromUrls_(existingS, existingT);
  }

  // 3) 形式チェック（怪しいIDは捨てる）
  ids = YahooVPS__filterValidImageIds_(ids);

  if (!ids.length) {
    Logger.log('    削除対象なし');
    return;
  }

  Logger.log('    削除対象: ' + ids.length + '件');

  // 削除
  YahooVPS__deleteImages_(sellerId, ids);

  // 反映（ロック避けで少し待つ）
  Utilities.sleep(1500);
  YahooVPS__submitItem_(sellerId, code);

  Logger.log('    削除反映完了');
}

function YahooVPS__getImageIds_(sellerId, code) {
  // queryを安全化（im-01004対策）
  var q = YahooVPS__safeQuery_(code);
  if (!q) return [];

  var url = 'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/itemImageList' +
    '?seller_id=' + encodeURIComponent(sellerId) +
    '&query=' + encodeURIComponent(q) +
    '&results=100';

  var resp = YahooVPS__apiCall_('get', url);
  var body = resp.getContentText() || '';

  var ids = [];
  var idMatches = body.match(/<Id>([^<]+)<\/Id>/g) || [];
  var nameMatches = body.match(/<Name>([^<]+)<\/Name>/g) || [];

  for (var i = 0; i < idMatches.length && i < nameMatches.length; i++) {
    var id = idMatches[i].replace(/<\/?Id>/g, '');
    var name = nameMatches[i].replace(/<\/?Name>/g, '').toLowerCase();

    // nameを安全化した上で、qを含むものだけ拾う（過取得を防ぐ）
    var nameSafe = YahooVPS__safeQuery_(name);
    if (nameSafe && nameSafe.indexOf(q) !== -1) ids.push(id);
  }

  return ids;
}

function YahooVPS__safeQuery_(s) {
  // 英数以外→スペース、連続スペース→1、trim、長すぎるならカット
  var q = String(s || '').toLowerCase()
    .replace(/[^0-9a-z]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!q) return '';
  if (q.length > 80) q = q.slice(0, 80).trim();
  return q;
}

function YahooVPS__extractIdsFromUrls_(sUrl, tUrl) {
  var urls = [];
  if (sUrl && YahooVPS__isYahooUrl_(sUrl)) urls.push(sUrl);
  if (tUrl) {
    tUrl.split(';').forEach(function(u) {
      u = YahooVPS__cleanCellUrl_(u);
      if (u && YahooVPS__isYahooUrl_(u)) urls.push(u);
    });
  }

  var ids = [];
  urls.forEach(function(u) {
    var id = YahooVPS__extractImageIdFromYahooUrl_(u);
    if (id) ids.push(id);
  });

  return ids;
}

function YahooVPS__extractImageIdFromYahooUrl_(url) {
  // 例）.../i/STORE/im12345.jpg みたいなのを想定して、
  // - 最後のパス要素を取る
  // - クエリ/フラグメント除去
  // - 拡張子除去
  var u = String(url || '');
  u = u.split('#')[0].split('?')[0];

  var parts = u.split('/');
  if (!parts.length) return '';

  var last = parts[parts.length - 1] || '';
  last = last.trim();
  if (!last) return '';

  // 拡張子除去
  last = last.replace(/\.(jpg|jpeg|png|webp|gif)$/i, '');

  // たまに末尾に変な記号がつくのを除去
  last = last.replace(/[^0-9a-zA-Z_-]/g, '');

  return last;
}

function YahooVPS__filterValidImageIds_(ids) {
  var seen = {};
  return (ids || [])
    .map(function(x){ return String(x || '').trim(); })
    .filter(function(x){
      if (!x) return false;
      // Yahooのimage_idは英数・記号限定想定（ここで弾く）
      if (!/^[0-9A-Za-z_-]{5,80}$/.test(x)) return false;
      if (seen[x]) return false;
      seen[x] = true;
      return true;
    });
}

function YahooVPS__deleteImages_(sellerId, imageIds) {
  if (!imageIds.length) return;

  var url = 'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/deleteItemImage';

  for (var i = 0; i < imageIds.length; i += 100) {
    var batch = imageIds.slice(i, i + 100);

    try {
      YahooVPS__apiCall_('post', url, {
        seller_id: sellerId,
        image_id: batch.join(',')
      });
    } catch (e) {
      var msg = (e && e.message) ? e.message : String(e);

      // ロック系は待って再試行
      if (msg.indexOf('ed-00006') !== -1 || msg.indexOf('反映またはアップロード中') !== -1) {
        Utilities.sleep(2500);
        YahooVPS__apiCall_('post', url, {
          seller_id: sellerId,
          image_id: batch.join(',')
        });
      } else {
        // im-04005が出たら「混ざってるIDが悪い」なので、ここでは落とさずログにして継続も可能
        // ただし今回は確実運用のため、エラーで止める（原因追跡しやすい）
        throw e;
      }
    }
  }
}

/** =====================================================
 * submitItem（反映）
 * ===================================================== */
function YahooVPS__submitItem_(sellerId, code) {
  var url = 'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/submitItem';

  try {
    YahooVPS__apiCall_('post', url, {
      seller_id: sellerId,
      item_code: code
    });
  } catch (e) {
    var msg = (e && e.message) ? e.message : String(e);
    if (msg.indexOf('it-07004') !== -1) {
      Logger.log('    submitスキップ（新規ページ）');
      return;
    }
    if (msg.indexOf('ed-00006') !== -1 || msg.indexOf('反映またはアップロード中') !== -1) {
      Utilities.sleep(2500);
      YahooVPS__apiCall_('post', url, {
        seller_id: sellerId,
        item_code: code
      });
      return;
    }
    throw e;
  }
}

/** =====================================================
 * API呼び出し
 * ===================================================== */
function YahooVPS__apiCall_(method, url, payload) {
  var token = YahooVPS__getAccessToken_();

  var options = {
    method: method,
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  };
  if (payload) options.payload = payload;

  var resp = UrlFetchApp.fetch(url, options);
  var code = resp.getResponseCode();
  var body = resp.getContentText() || '';

  if (code === 401 || body.indexOf('AccessToken has been expired') !== -1 || body.indexOf('px-04102') !== -1) {
    token = YahooVPS__getAccessToken_(true);
    options.headers['Authorization'] = 'Bearer ' + token;
    resp = UrlFetchApp.fetch(url, options);
    code = resp.getResponseCode();
    body = resp.getContentText() || '';
  }

  if (code !== 200) {
    throw new Error('API失敗 HTTP=' + code + ' ' + body.slice(0, 300));
  }

  return resp;
}

/** =====================================================
 * トークン管理
 * ===================================================== */
var YahooVPS__tokenCache = null;

function YahooVPS__getAccessToken_(forceRefresh) {
  if (!forceRefresh && YahooVPS__tokenCache && YahooVPS__tokenCache.expireAt > Date.now() + 60000) {
    return YahooVPS__tokenCache.token;
  }

  var clientId = YahooVPS__getProp_('Y_CLIENT_ID');
  var clientSecret = YahooVPS__getProp_('Y_CLIENT_SECRET');
  var refreshToken = YahooVPS__getProp_('Y_REFRESH_TOKEN');

  var resp = UrlFetchApp.fetch('https://auth.login.yahoo.co.jp/yconnect/v2/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      grant_type: 'refresh_token',
      refresh_token: refreshToken,
      client_id: clientId,
      client_secret: clientSecret
    },
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() !== 200) {
    throw new Error('トークン更新失敗: ' + resp.getContentText());
  }

  var obj = JSON.parse(resp.getContentText());

  YahooVPS__tokenCache = {
    token: obj.access_token,
    expireAt: Date.now() + (Number(obj.expires_in || 1800) * 1000)
  };

  return YahooVPS__tokenCache.token;
}

function YahooVPS__getProp_(key) {
  var v = PropertiesService.getScriptProperties().getProperty(key);
  if (!v) throw new Error('設定なし: ' + key);
  return String(v).trim();
}

/** =====================================================
 * ユーティリティ
 * ===================================================== */
function YahooVPS__cleanCellUrl_(v) {
  var s = String(v == null ? '' : v);
  s = s.replace(/\r|\n/g, '').trim();
  return s;
}

function YahooVPS__normalizeUrl_(u) {
  u = YahooVPS__cleanCellUrl_(u);
  if (!u) return '';
  u = u.replace(/\u3000/g, ' ').trim();
  u = u.replace(/\/+$/g, '');
  return u;
}

function YahooVPS__isYahooUrl_(url) {
  var u = YahooVPS__cleanCellUrl_(url).toLowerCase();
  return u.indexOf('https://shopping.c.yimg.jp/') === 0 ||
         u.indexOf('http://shopping.c.yimg.jp/') === 0 ||
         u.indexOf('https://item-shopping.c.yimg.jp/') === 0 ||
         u.indexOf('http://item-shopping.c.yimg.jp/') === 0;
}

/** =====================================================
 * トリガー管理
 * ===================================================== */
function YahooVPS__ensureTrigger_() {
  if (YahooVPS__hasTrigger_()) return;
  ScriptApp.newTrigger('YahooVPS_自動継続')
    .timeBased()
    .after(60 * 1000)
    .create();
}

function YahooVPS__deleteTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'YahooVPS_自動継続') ScriptApp.deleteTrigger(t);
  });
}

function YahooVPS__hasTrigger_() {
  return ScriptApp.getProjectTriggers().some(function(t) {
    return t.getHandlerFunction() === 'YahooVPS_自動継続';
  });
}

/** =====================================================
 * デバッグ
 * ===================================================== */
function YahooVPS_デバッグ() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var cfg = YahooVPS.設定;

  var msg = '=== 設定確認 ===\n';
  msg += 'Y_CLIENT_ID: ' + (props.getProperty('Y_CLIENT_ID') ? '✅' : '❌未設定') + '\n';
  msg += 'Y_CLIENT_SECRET: ' + (props.getProperty('Y_CLIENT_SECRET') ? '✅' : '❌未設定') + '\n';
  msg += 'Y_REFRESH_TOKEN: ' + (props.getProperty('Y_REFRESH_TOKEN') ? '✅' : '❌未設定') + '\n';
  msg += 'Y_SELLER_ID: ' + (props.getProperty('Y_SELLER_ID') || '❌未設定') + '\n\n';

  msg += 'VPS_UPLOAD_URL: ' + cfg.VPS_UPLOAD_URL + '\n';
  msg += '許容_部分成功: ' + cfg.許容_部分成功 + '\n';

  msg += '\n=== シート確認 ===\n';
  msg += 'シート名: ' + cfg.シート名 + '\n';
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.シート名);
    if (sheet) {
      var lastRow = sheet.getLastRow();
      var numRows = lastRow - cfg.ヘッダー行数;
      msg += 'シート: ✅ 存在\n';
      msg += '最終行: ' + lastRow + '\n';
      msg += 'データ行数: ' + numRows + '\n';
    } else {
      msg += 'シート: ❌ 見つからない\n';
    }
  } catch (e) {
    msg += 'シート: ❌ エラー (' + e.message + ')\n';
  }

  msg += '\n=== Driveフォルダ確認 ===\n';
  try {
    var folder = DriveApp.getFolderById(cfg.画像ルートフォルダID);
    msg += 'フォルダ: ✅ ' + folder.getName() + '\n';
  } catch (e) {
    msg += 'フォルダ: ❌ エラー (' + e.message + ')\n';
  }

  ui.alert(msg);
}
