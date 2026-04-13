/******************************************************
 * Yahoo画像アップロード（VPS API連携・追跡ログ強化・単一実行版）
 * v5.5 ST_ONLY + GETITEM_VERIFY
 *
 * v5.5 変更点:
 *   - U/V列を完全廃止。S/T列のみをソースとして使用。
 *   - 対象抽出・診断ログをS/T専用に整理。
 *   - 検証に getItem API を追加（商品との画像紐付け確認）。
 *
 * 注意:
 *   - getItem は参照API（read-only）。紐付け実処理は VPS submit 側。
 ******************************************************/

'use strict';

var YahooUpload = YahooUpload || {};

/** =============================
 * 設定（ここだけ触ればOK）
 * ============================= */
YahooUpload.設定 = {
  // VPS API Endpoint
  API_URL: 'https://img.niyantarose.com/_api/imgcv_api.php',

  // Yahoo 商品参照API（本番）
  GETITEM_URL: 'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/getItem',

  // スプレッドシートID（空欄ならアクティブ）
  スプレッドシートID: '1jN-3aLsLg_wRjLQU-mDqaImVVBgdpvwJWsHegxSJRF0',

  // 対象シート
  シート名: '①商品入力シート',
  ヘッダー行数: 2,
  列_コード: 3,   // C
  列_S: 19,       // S（メインURL）
  列_T: 20,       // T（詳細URL群）

  // 1回で送る商品数
  BATCH_SIZE: 200,

  // 受付後に VPSへ submit を叩く
  SUBMIT_AFTER: true,

  // 検証（軽量）
  VERIFY: {
    CDN_CHECK: true,      // main CDN URL の 200/404
    PAGE_CHECK: true,     // 商品ページHTML
    GETITEM_CHECK: true,  // getItem API で画像紐付け確認
    MAX_CODES: 30
  },

  // ストア商品ページURL
  STORE_BASE: 'https://store.shopping.yahoo.co.jp/niyantarose/',

  // ログ
  LOG: {
    トースト: false,     // 右下のポップアップ通知をオフ
    実行ログ: true,      // 実行ログ（エクスキューズログ）のみオン
    シートログ: false    // シートへの書き込みをオフ
  },
  LOG_SHEET_RUN: 'YahooVPS_LOG',
  LOG_SHEET_ITEM: 'YahooVPS_LOG_ITEMS',
  LOG_JSON_MAXLEN: 45000,

  // true: S/TがYahoo CDN URLでもソースとして使う
  // false: Yahoo CDN URLはソースから除外
  CDN_URLもソースにする: true,

  // 6分制限保護
  MAX_RUN_MS: 5 * 60 * 1000,
  SAFE_MARGIN_MS: 25 * 1000
};

/** =============================
 * ScriptProperties keys
 * ============================= */
var YU_KEY_ACCESS_TOKEN   = 'Y_ACCESS_TOKEN';
var YU_KEY_EXPIRES_AT     = 'Y_ACCESS_EXPIRES_AT';
var YU_KEY_REFRESH_TOKEN  = 'Y_REFRESH_TOKEN';
var YU_KEY_CLIENT_ID      = 'Y_CLIENT_ID';
var YU_KEY_CLIENT_SECRET  = 'Y_CLIENT_SECRET';
var YU_KEY_SELLER_ID      = 'Y_SELLER_ID';
var YU_KEY_VPS_SECRET     = 'IMGCV_VPS_SECRET';
var YU_TOKEN_URL          = 'https://auth.login.yahoo.co.jp/yconnect/v2/token';

/** =============================
 * 入口（ボタンはこれだけ）
 * ============================= */
function YahooUpload_実行() {
  var cfg = YahooUpload.設定;
  var ss = YahooUpload__ss_();
  var props = PropertiesService.getScriptProperties();

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    return YahooUpload__notify_(ss, '他の実行が動いてるみたいや。少し待ってから再実行してな。');
  }

  var runId = Utilities.getUuid();
  var tStart = Date.now();

  try {
    if (cfg.LOG && cfg.LOG.シートログ) YahooUpload__ensureLogSheets_(ss);

    var sh = ss.getSheetByName(cfg.シート名);
    if (!sh) throw new Error('シートが見つからん: ' + cfg.シート名);

    var secret = String(props.getProperty(YU_KEY_VPS_SECRET) || '').trim();
    if (!secret) throw new Error('VPS Secret が未設定やで（IMGCV_VPS_SECRET）');

    var sellerId = String(props.getProperty(YU_KEY_SELLER_ID) || 'niyantarose').trim();
    var accessToken = YahooUpload__getValidAccessToken_(false);

    // データ取得（C〜T）
    var lastRow = YahooUpload__getLastDataRowByCol_(sh, cfg.列_コード, cfg.ヘッダー行数);
    var startRow = cfg.ヘッダー行数 + 1;
    var numRows = Math.max(0, lastRow - cfg.ヘッダー行数);
    if (numRows <= 0) throw new Error('データが無い（コード列が空）');

    var width = Math.max(cfg.列_T, cfg.列_S, cfg.列_コード);
    var values = sh.getRange(startRow, 1, numRows, width).getValues();

    YahooUpload__log_(
      'データ読込完了: startRow=' + startRow + ' numRows=' + numRows + ' width=' + width + ' lastRow=' + lastRow,
      'DATA'
    );

    // 対象抽出（S/Tのみ）
    var targets = YahooUpload__collectTargets_(values, startRow, cfg);

    if (targets.length === 0) {
      var diag = YahooUpload._lastDiag || {};
      var diagMsg = 'アップロード対象が無い。'
        + ' 読込行数=' + (diag.totalRows || 0)
        + ' コード有り=' + (diag.codesFound || 0)
        + ' Yahoo既存skip=' + (diag.skippedYahooOnly || 0)
        + ' URL無しskip=' + (diag.skippedNoUrl || 0)
        + ' コード無しskip=' + (diag.skippedNoCode || 0);

      if (diag.samples && diag.samples.length > 0) {
        diagMsg += '\n--- 先頭行のセル値（診断用）---';
        for (var si = 0; si < diag.samples.length; si++) {
          var ds = diag.samples[si];
          diagMsg += '\n  row' + ds.row
            + ' code=' + ds.code
            + ' | S[型:' + ds.sType + ']=' + (ds.S || '(空)')
            + ' | T[型:' + ds.tType + ']=' + (ds.T || '(空)')
            + ' | skip理由=' + (ds.skipReason || '(対象)');
        }
      }

      if ((diag.skippedYahooOnly || 0) > 0 && (diag.skippedYahooOnly || 0) === (diag.codesFound || 0)) {
        diagMsg += '\n\n【ヒント】全行のS/T列がYahoo CDN URLです。'
          + '\nCDN_URLもソースにする=true で再実行するか、S/Tに元URLを入れてください。';
      }
      if ((diag.skippedNoUrl || 0) > 0) {
        diagMsg += '\n\n【ヒント】S/T列にURLが見つからない行があります。'
          + '\n列番号設定を確認してください（現在: S=' + cfg.列_S + ', T=' + cfg.列_T + '）'
          + '\nセル値がURL（http://〜）で始まるか確認してください。';
      }

      YahooUpload__logRun_(ss, runId, 'NO_TARGETS', 0, diag.codesFound || 0, 0, '', 0, false, diagMsg, '');
      YahooUpload__log_(diagMsg, 'DIAG');
      return YahooUpload__notify_(ss, diagMsg);
    }

    YahooUpload__notify_(ss, 'VPSへ指示開始：codes=' + targets.length + ' runId=' + runId);

    var BATCH = Math.max(1, Number(cfg.BATCH_SIZE) || 200);
    var sent = 0;
    var chunkIndex = 0;

    for (var p = 0; p < targets.length; ) {
      chunkIndex++;

      if (Date.now() - tStart > (cfg.MAX_RUN_MS - cfg.SAFE_MARGIN_MS)) {
        YahooUpload__notify_(ss, '時間が近いからここで止めるで（途中まで送信済み） sent=' + sent + '/' + targets.length);
        break;
      }

      var end = Math.min(p + BATCH, targets.length);
      var chunk = targets.slice(p, end);

      var items = YahooUpload__buildItemsFromSources_(chunk);

      if (items.length === 0) {
        YahooUpload__logRun_(ss, runId, 'UPLOAD', chunkIndex, chunk.length, 0, 200, 0, false, 'SKIP(items=0)', '');
        YahooUpload__logItemSkips_(ss, runId, chunk, 'NO_ITEMS_FROM_SOURCE');
        p = end;
        continue;
      }

      var payload = {
        mode: 'yahoo_upload',
        run_id: runId,
        seller_id: sellerId,
        secret: secret,
        yahoo_token: accessToken,
        items: items,
        debug: { verbose: 1, return_details: 1 }
      };

      YahooUpload__log_('[API] POST chunk=' + chunkIndex + ' codes=' + chunk.length + ' items=' + items.length, 'API');

      var post = YahooUpload__postJsonDebug_(cfg.API_URL, payload, secret);

      YahooUpload__logRun_(
        ss, runId, 'UPLOAD', chunkIndex, chunk.length, items.length, post.http, post.elapsed_ms,
        (post.json && post.json.ok) ? true : false,
        YahooUpload__pickMsg_(post.json, post.body),
        YahooUpload__stringifyForCell_(post.json || { raw: post.body }, cfg.LOG_JSON_MAXLEN)
      );

      if (post.http !== 200 || !(post.json && post.json.ok)) {
        YahooUpload__logItemSkips_(ss, runId, chunk, 'VPS_UPLOAD_FAILED');
        throw new Error('VPS error HTTP=' + post.http + ' msg=' + YahooUpload__pickMsg_(post.json, post.body));
      }

      YahooUpload__logItemDetailsFromResponse_(ss, runId, chunk, post.json);

      sent += chunk.length;
      YahooUpload__notify_(ss, '受付OK：' + sent + '/' + targets.length + ' codes');

      p = end;
    }

    if (cfg.SUBMIT_AFTER) {
      var sub = YahooUpload__trySubmit_(ss, runId, sellerId, secret, accessToken, targets);

      var subDetail = 'submit試行：HTTP=' + sub.http
        + ' | ok=' + (sub.ok ? 'true' : 'false')
        + ' | items送信=' + (sub.itemsCount || 0) + '件'
        + ' | msg=' + (sub.msg || '(空)')
        + ' | 処理数=' + (sub.processed || 0)
        + ' | 成功=' + (sub.succeeded || 0)
        + ' | 失敗=' + (sub.failed || 0);

      YahooUpload__notify_(ss, subDetail);
      YahooUpload__log_(subDetail + '\n応答JSON: ' + (sub.jsonStr || '(なし)'), 'SUBMIT_RESULT');
    }

    // 検証（CDN / 商品ページ / getItem）
    YahooUpload__verify_(ss, runId, sellerId, accessToken, targets);

    YahooUpload__notify_(ss, '完了（GAS側）：' + ((Date.now() - tStart) / 1000).toFixed(1) + '秒 runId=' + runId);

  } catch (e) {
    YahooUpload__notify_(ss, 'エラー: ' + (e.message || e));
    throw e;
  } finally {
    try { lock.releaseLock(); } catch (_) {}
    YahooUpload__log_('finally runId=' + runId, 'DONE');
  }
}

/** =============================
 * 対象抽出：コードがある行で、S/TにURLがある行
 * ============================= */
function YahooUpload__collectTargets_(values, startRow, cfg) {
  var out = [];
  var diag = {
    totalRows: values.length,
    codesFound: 0,
    skippedYahooOnly: 0,
    skippedNoUrl: 0,
    skippedNoCode: 0,
    samples: []
  };

  var cCode = cfg.列_コード - 1;
  var cS = cfg.列_S - 1;
  var cT = cfg.列_T - 1;

  for (var i = 0; i < values.length; i++) {
    var rawCode = values[i][cCode];
    var code = YahooUpload__safeStr_(rawCode).trim();

    if (!code) {
      diag.skippedNoCode++;
      if (diag.samples.length < 5) {
        var rawS_nc = values[i][cS];
        var rawT_nc = values[i][cT];
        diag.samples.push({
          row: startRow + i,
          code: '(空)',
          S: YahooUpload__safeStr_(rawS_nc).substring(0, 100),
          T: YahooUpload__safeStr_(rawT_nc).substring(0, 100),
          sType: YahooUpload__typeLabel_(rawS_nc),
          tType: YahooUpload__typeLabel_(rawT_nc),
          skipReason: 'コード空'
        });
      }
      continue;
    }

    diag.codesFound++;

    var rawS = values[i][cS];
    var rawT = values[i][cT];
    var sNow = YahooUpload__safeStr_(rawS).trim();
    var tNow = YahooUpload__safeStr_(rawT).trim();

    var allowCdn = cfg.CDN_URLもソースにする;
    var srcMain = (!allowCdn && YahooUpload__isYahooUrl_(sNow)) ? '' : sNow;
    var srcDetail = (!allowCdn && YahooUpload__isYahooUrl_(tNow)) ? '' : tNow;

    var mainUrl = YahooUpload__extractUrl_(srcMain);
    var detailUrls = YahooUpload__extractUrls_(srcDetail);

    if (diag.samples.length < 10) {
      var skipReason = '';
      if (!mainUrl && detailUrls.length === 0) {
        if (!allowCdn && (YahooUpload__isYahooUrl_(sNow) || YahooUpload__isYahooUrl_(tNow))) {
          skipReason = 'S/TがYahoo URL(CDN許可=off)';
        } else if (!sNow && !tNow) {
          skipReason = 'S/T共に空';
        } else {
          skipReason = 'URL抽出失敗(srcMain=' + srcMain.substring(0, 50) + ')';
        }
      }

      diag.samples.push({
        row: startRow + i,
        code: code,
        S: sNow.substring(0, 100),
        T: tNow.substring(0, 100),
        sType: YahooUpload__typeLabel_(rawS),
        tType: YahooUpload__typeLabel_(rawT),
        skipReason: skipReason || '(対象OK)'
      });
    }

    if (!mainUrl && detailUrls.length === 0) {
      var hasYahoo = !allowCdn && (YahooUpload__isYahooUrl_(sNow) || YahooUpload__isYahooUrl_(tNow));
      if (hasYahoo) diag.skippedYahooOnly++;
      else diag.skippedNoUrl++;
      continue;
    }

    out.push({
      row: startRow + i,
      code: code,
      srcMainRaw: srcMain,
      srcDetailRaw: srcDetail,
      detailCount: detailUrls.length
    });
  }

  YahooUpload._lastDiag = diag;

  YahooUpload__log_(
    'collectTargets 結果: totalRows=' + diag.totalRows
    + ' codesFound=' + diag.codesFound
    + ' targets=' + out.length
    + ' skippedYahoo=' + diag.skippedYahooOnly
    + ' skippedNoUrl=' + diag.skippedNoUrl
    + ' skippedNoCode=' + diag.skippedNoCode,
    'COLLECT'
  );

  return out;
}

/** =============================
 * items生成
 * - S: code.jpg
 * - T: code_1.jpg, code_2.jpg ...
 * ============================= */
function YahooUpload__buildItemsFromSources_(chunk) {
  var out = [];
  for (var i = 0; i < chunk.length; i++) {
    var code = chunk[i].code;

    var mainUrl = YahooUpload__extractUrl_(chunk[i].srcMainRaw);
    if (mainUrl) out.push(YahooUpload__makeItem_(mainUrl, code, code + '.jpg'));

    var detailUrls = YahooUpload__extractUrls_(chunk[i].srcDetailRaw);
    for (var k = 0; k < detailUrls.length; k++) {
      out.push(YahooUpload__makeItem_(detailUrls[k], code, code + '_' + (k + 1) + '.jpg'));
    }
  }
  return out;
}


/** =============================
 * 検証（CDN + 商品ページ + getItem）
 * ============================= */
function YahooUpload__verify_(ss, runId, sellerId, accessToken, targets) {
  var cfg = YahooUpload.設定;
  if (!cfg.VERIFY) return;

  var n = Math.min(Number(cfg.VERIFY.MAX_CODES || 20), targets.length);
  if (n <= 0) return;

  var sample = targets.slice(0, n);

  // CDN main check
  if (cfg.VERIFY.CDN_CHECK) {
    var cdnBase = 'https://item-shopping.c.yimg.jp/i/n/' + sellerId + '_';
    var reqs = [];
    var meta = [];

    for (var i = 0; i < sample.length; i++) {
      var code = sample[i].code;
      var u = cdnBase + code + '.jpg';
      reqs.push({ url: u, muteHttpExceptions: true, followRedirects: true });
      meta.push({ code: code, url: u });
    }

    var resps = UrlFetchApp.fetchAll(reqs);
    var ok = 0, ng = 0;

    for (var k = 0; k < resps.length; k++) {
      var http = resps[k].getResponseCode();
      if (http === 200) ok++; else ng++;
      YahooUpload__logItem_(ss, runId, meta[k].code, '', 'CDN_MAIN', (http === 200), 'http=' + http, meta[k].url);
    }

    YahooUpload__notify_(ss, 'CDN検証結果：OK=' + ok + ' / NG=' + ng + '（NG=404なら未アップの可能性高い）');
  }

  // Product page check
  if (cfg.VERIFY.PAGE_CHECK) {
    var base = String(cfg.STORE_BASE || '').trim();
    if (base) {
      var reqs2 = [];
      var meta2 = [];

      for (var j = 0; j < sample.length; j++) {
        var c = sample[j].code;
        var pageUrl = base + encodeURIComponent(c) + '.html';
        reqs2.push({ url: pageUrl, muteHttpExceptions: true, followRedirects: true });
        meta2.push({ code: c, page: pageUrl });
      }

      var resps2 = UrlFetchApp.fetchAll(reqs2);
      var hit = 0;

      for (var m = 0; m < resps2.length; m++) {
        var http2 = resps2[m].getResponseCode();
        var body = String(resps2[m].getContentText() || '');
        var found = (http2 === 200) && (
          body.indexOf(meta2[m].code + '.jpg') !== -1 ||
          body.indexOf('item-shopping.c.yimg.jp') !== -1
        );

        if (found) hit++;
        YahooUpload__logItem_(ss, runId, meta2[m].code, '', 'PAGE_HTML', found, 'http=' + http2, meta2[m].page);
      }

      YahooUpload__notify_(ss, '商品ページ検証：HIT=' + hit + '/' + sample.length + '（HIT少ない＝紐付け/submit未完の疑い）');
    }
  }

  // getItem check
  if (cfg.VERIFY.GETITEM_CHECK) {
    YahooUpload__verifyGetItem_(ss, runId, sellerId, accessToken, sample);
  }
}

function YahooUpload__verifyGetItem_(ss, runId, sellerId, accessToken, sample) {
  var cfg = YahooUpload.設定;
  var api = String(cfg.GETITEM_URL || '').trim();
  if (!api) return;

  if (!accessToken) {
    YahooUpload__notify_(ss, 'getItem検証スキップ：アクセストークンが空');
    return;
  }

  var reqs = [];
  var meta = [];

  for (var i = 0; i < sample.length; i++) {
    var code = sample[i].code;
    var url = api
      + '?seller_id=' + encodeURIComponent(sellerId)
      + '&item_code=' + encodeURIComponent(code);

    reqs.push({
      url: url,
      method: 'get',
      headers: { Authorization: 'Bearer ' + accessToken },
      muteHttpExceptions: true,
      followRedirects: true
    });
    meta.push({ code: code, url: url });
  }

  var resps = UrlFetchApp.fetchAll(reqs);
  var ok = 0, ng = 0;

  for (var k = 0; k < resps.length; k++) {
    var http = resps[k].getResponseCode();
    var body = String(resps[k].getContentText() || '');
    var sum = YahooUpload__parseGetItemSummary_(body);

    var hasImage = (http === 200) && (
      sum.mainImageCount > 0 || sum.libImageCount > 0 || sum.subCodeImageCount > 0
    );

    if (hasImage) ok++; else ng++;

    var msg = 'http=' + http
      + ' main=' + sum.mainImageCount
      + ' lib=' + sum.libImageCount
      + ' sub=' + sum.subCodeImageCount
      + ' editing=' + (sum.editingFlag || '')
      + ' update=' + (sum.updateTime || '');

    YahooUpload__logItem_(ss, runId, meta[k].code, '', 'GETITEM', hasImage, msg, meta[k].url);
  }

  YahooUpload__notify_(ss, 'getItem検証：画像あり=' + ok + ' / 画像なし=' + ng);
}

function YahooUpload__parseGetItemSummary_(xmlText) {
  var s = String(xmlText || '');

  var out = {
    itemCode: YahooUpload__extractFirstTag_(s, 'ItemCode'),
    updateTime: YahooUpload__extractFirstTag_(s, 'UpdateTime'),
    editingFlag: YahooUpload__extractFirstTag_(s, 'EditingFlag'),
    mainImageCount: YahooUpload__countNonEmptyTagValues_(s, /<Image(?:\s[^>]*)?>([\s\S]*?)<\/Image>/g),
    libImageCount: YahooUpload__countNonEmptyTagValues_(s, /<LibImage\d+(?:\s[^>]*)?>([\s\S]*?)<\/LibImage\d+>/g),
    subCodeImageCount: 0
  };

  var a = s.match(/<SubCodeImage\b[^>]*\bexist_flag="1"[^>]*>/g);
  if (a) out.subCodeImageCount += a.length;

  var b = s.match(/<SubCodeImage\b[^>]*\bexist="1"[^>]*>/g);
  if (b) out.subCodeImageCount += b.length;

  if (out.subCodeImageCount === 0) {
    out.subCodeImageCount = YahooUpload__countNonEmptyTagValues_(s, /<SubCodeImage(?:\s[^>]*)?>([\s\S]*?)<\/SubCodeImage>/g);
  }

  return out;
}

function YahooUpload__countNonEmptyTagValues_(s, re) {
  var count = 0;
  var m;
  while ((m = re.exec(s)) !== null) {
    var v = String(m[1] || '').replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, '$1').trim();
    if (v) count++;
  }
  return count;
}

function YahooUpload__extractFirstTag_(s, tagName) {
  var re = new RegExp('<' + tagName + '>([\\s\\S]*?)<\\/' + tagName + '>', 'i');
  var m = s.match(re);
  if (!m) return '';
  return String(m[1] || '').replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, '$1').trim();
}

/** =============================
 * VPS POST（JSON）
 * ============================= */
function YahooUpload__postJsonDebug_(url, payloadObj, secret) {
  var t0 = Date.now();
  var resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payloadObj),
    headers: {
      'X-VPS-SECRET': String(secret || ''),
      'X-Requested-With': 'GAS',
      'X-Run-Id': String(payloadObj.run_id || '')
    },
    muteHttpExceptions: true,
    followRedirects: true
  });
  var elapsed = Date.now() - t0;

  var http = resp.getResponseCode();
  var body = String(resp.getContentText() || '');

  var json = null;
  try { json = JSON.parse(body); } catch (_) {}

  return { http: http, body: body, json: json, elapsed_ms: elapsed };
}

/** =============================
 * Token管理（Refresh）
 * ============================= */
function YahooUpload__getValidAccessToken_(force) {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty(YU_KEY_ACCESS_TOKEN);
  var exp = Number(props.getProperty(YU_KEY_EXPIRES_AT) || 0);

  if (!force && token && exp > Date.now() + 60000) return token;

  var ulock = LockService.getUserLock();
  ulock.waitLock(30000);

  try {
    token = props.getProperty(YU_KEY_ACCESS_TOKEN);
    exp = Number(props.getProperty(YU_KEY_EXPIRES_AT) || 0);
    if (!force && token && exp > Date.now() + 60000) return token;

    var cid = props.getProperty(YU_KEY_CLIENT_ID);
    var csec = props.getProperty(YU_KEY_CLIENT_SECRET);
    var refr = props.getProperty(YU_KEY_REFRESH_TOKEN);
    if (!cid || !csec || !refr) throw new Error('初期設定（client/secret/refresh）が未設定やで');

    var creds = Utilities.base64Encode(cid + ':' + csec);

    var resp = UrlFetchApp.fetch(YU_TOKEN_URL, {
      method: 'post',
      headers: { Authorization: 'Basic ' + creds },
      payload: { grant_type: 'refresh_token', refresh_token: refr },
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() !== 200) {
      throw new Error('トークン更新失敗: ' + resp.getContentText());
    }

    var json = JSON.parse(resp.getContentText());
    props.setProperty(YU_KEY_ACCESS_TOKEN, json.access_token);
    props.setProperty(YU_KEY_EXPIRES_AT, String(Date.now() + (Number(json.expires_in || 1800) * 1000)));

    if (json.refresh_token) props.setProperty(YU_KEY_REFRESH_TOKEN, json.refresh_token);

    YahooUpload__log_('token refreshed exp=' + props.getProperty(YU_KEY_EXPIRES_AT), 'TOKEN');
    return json.access_token;

  } finally {
    try { ulock.releaseLock(); } catch (_) {}
  }
}

/** =============================
 * SS / log / url helpers
 * ============================= */
function YahooUpload__ss_() {
  var id = YahooUpload.設定.スプレッドシートID;
  if (id && id.length > 10) return SpreadsheetApp.openById(id);
  return SpreadsheetApp.getActiveSpreadsheet();
}

function YahooUpload__notify_(ss, msg) {
  var cfg = YahooUpload.設定;
  var text = String(msg || '');
  YahooUpload__log_(text, 'YahooVPS');
  if (cfg.LOG && cfg.LOG.トースト) {
    try { (ss || SpreadsheetApp.getActiveSpreadsheet()).toast(text, 'YahooVPS', 10); } catch (_) {}
  }
}

function YahooUpload__log_(msg, tag) {
  var cfg = YahooUpload.設定;
  if (!(cfg.LOG && cfg.LOG.実行ログ)) return;
  try { Logger.log('[' + (tag || 'LOG') + '] ' + String(msg)); } catch (_) {}
}

function YahooUpload__getLastDataRowByCol_(sheet, col, headerRows) {
  var start = headerRows + 1;
  var maxRows = sheet.getMaxRows();
  var n = maxRows - headerRows;
  if (n <= 0) return headerRows;

  var values = sheet.getRange(start, col, n, 1).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    var v = YahooUpload__safeStr_(values[i][0]).trim();
    if (v !== '') return start + i;
  }
  return headerRows;
}

function YahooUpload__isYahooUrl_(val) {
  var s = String(val || '').toLowerCase();
  return (s.indexOf('shopping.c.yimg.jp') !== -1) ||
         (s.indexOf('item-shopping.c.yimg.jp') !== -1) ||
         (s.indexOf('yimg.jp') !== -1);
}

/** =============================
 * URL抽出
 * ============================= */
function YahooUpload__extractUrl_(val) {
  var s = YahooUpload__safeStr_(val);

  // 不可視文字を除去
  s = s.replace(/[\u200B-\u200D\uFEFF\u00AD]/g, '');

  var m = s.match(/https?:\/\/[^\s;"'<>]+/);
  if (m) return m[0].replace(/[.,;:!?)]+$/, '');

  var m2 = s.match(/\/\/[a-zA-Z0-9][^\s;"'<>]+/);
  if (m2) return ('https:' + m2[0]).replace(/[.,;:!?)]+$/, '');

  // ベアDrive ID
  var m3 = s.match(/^([a-zA-Z0-9_-]{25,})$/);
  if (m3) return 'DRIVE_ID:' + m3[1];

  return '';
}

function YahooUpload__extractUrls_(val) {
  var s = YahooUpload__safeStr_(val).trim();
  if (!s) return [];
  var parts = s.split(/[\r\n;]+/);
  var out = [];
  for (var i = 0; i < parts.length; i++) {
    var u = YahooUpload__extractUrl_(parts[i]);
    if (u) out.push(u);
  }
  var seen = {};
  return out.filter(function(x) {
    if (seen[x]) return false;
    seen[x] = 1;
    return true;
  });
}

function YahooUpload__makeItem_(url, code, filename) {
  var type = 'url';
  var src = url;

  if (String(url).indexOf('DRIVE_ID:') === 0) {
    type = 'id';
    src = url.substring(9);
    return { type: type, src: src, code: code, name: filename };
  }

  var m = String(url).match(/\/d\/([a-zA-Z0-9_-]+)/) || String(url).match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m) { type = 'id'; src = m[1]; }

  return { type: type, src: src, code: code, name: filename };
}

/** =============================
 * 安全な文字列変換
 * ============================= */
function YahooUpload__safeStr_(val) {
  if (val === null || val === undefined) return '';
  if (val === false) return '';
  if (val === 0) return '';
  if (val instanceof Date) return '';
  return String(val);
}

function YahooUpload__typeLabel_(val) {
  if (val === null || val === undefined) return 'null';
  if (val instanceof Date) return 'Date';
  if (typeof val === 'number') return 'num';
  if (typeof val === 'boolean') return 'bool';
  if (typeof val === 'string') return (val === '') ? 'empty' : 'str';
  return typeof val;
}

/** =============================
 * ログシート
 * ============================= */
function YahooUpload__ensureLogSheets_(ss) {
  var cfg = YahooUpload.設定;

  var runSh = ss.getSheetByName(cfg.LOG_SHEET_RUN) || ss.insertSheet(cfg.LOG_SHEET_RUN);
  var itemSh = ss.getSheetByName(cfg.LOG_SHEET_ITEM) || ss.insertSheet(cfg.LOG_SHEET_ITEM);

  if (runSh.getLastRow() === 0) {
    runSh.getRange(1, 1, 1, 10).setValues([[
      'ts', 'run_id', 'phase', 'chunk', 'codes', 'items', 'http', 'elapsed_ms', 'ok', 'message/json'
    ]]);
    runSh.setFrozenRows(1);
  }
  if (itemSh.getLastRow() === 0) {
    itemSh.getRange(1, 1, 1, 8).setValues([[
      'ts', 'run_id', 'code', 'row', 'type', 'ok', 'message', 'ref'
    ]]);
    itemSh.setFrozenRows(1);
  }
}

function YahooUpload__logRun_(ss, runId, phase, chunk, codes, items, http, ms, ok, msg, json) {
  var cfg = YahooUpload.設定;
  if (!(cfg.LOG && cfg.LOG.シートログ)) return;
  var sh = ss.getSheetByName(cfg.LOG_SHEET_RUN);
  if (!sh) return;

  var merged = String(msg || '');
  if (json) merged += '\n' + String(json);

  sh.appendRow([
    new Date(),
    String(runId || ''),
    String(phase || ''),
    Number(chunk || 0),
    Number(codes || 0),
    Number(items || 0),
    String(http || ''),
    Number(ms || 0),
    ok ? 'true' : 'false',
    merged
  ]);
}

function YahooUpload__logItem_(ss, runId, code, row, type, ok, message, ref) {
  var cfg = YahooUpload.設定;
  if (!(cfg.LOG && cfg.LOG.シートログ)) return;
  var sh = ss.getSheetByName(cfg.LOG_SHEET_ITEM);
  if (!sh) return;

  sh.appendRow([
    new Date(),
    String(runId || ''),
    String(code || ''),
    row ? String(row) : '',
    String(type || ''),
    ok ? 'true' : 'false',
    String(message || ''),
    String(ref || '')
  ]);
}

function YahooUpload__logItemSkips_(ss, runId, chunk, reason) {
  for (var i = 0; i < chunk.length; i++) {
    YahooUpload__logItem_(ss, runId, chunk[i].code, chunk[i].row, 'SKIP', false, reason, '');
  }
}

function YahooUpload__logItemDetailsFromResponse_(ss, runId, chunk, resJson) {
  if (!resJson) return;

  if (Array.isArray(resJson.details)) {
    for (var i = 0; i < resJson.details.length; i++) {
      var d = resJson.details[i] || {};
      var code = String(d.item_code || d.code || '').trim();
      if (!code) continue;

      YahooUpload__logItem_(
        ss, runId, code, '', 'VPS_DETAIL',
        (d.ok === true),
        String(d.status || d.message || d.error || ''),
        YahooUpload__stringifyForCell_(d, YahooUpload.設定.LOG_JSON_MAXLEN)
      );
    }
  }
}

function YahooUpload__stringifyForCell_(obj, maxLen) {
  var s = '';
  try { s = JSON.stringify(obj); } catch (e) { s = String(obj); }
  s = String(s || '');
  if (s.length > maxLen) s = s.slice(0, maxLen) + ' ...(truncated)';
  return s;
}

function YahooUpload__pickMsg_(json, body) {
  if (json) return String(json.message || json.status || json.error || json.err || '');
  return String(body || '').slice(0, 200);
}
function YahooUpload__trySubmit_(ss, runId, sellerId, secret, accessToken, targets) {
  var cfg = YahooUpload.設定;
  var PACK_URL = 'https://img.niyantarose.com/_api/imgcv_api_pack.php';

  var seen = {};
  var items = [];

  (targets || []).forEach(function(t) {
    var c = String((t && t.code) || '').trim();
    if (!c || seen[c]) return;
    seen[c] = 1;
    items.push(c);
  });

  if (items.length === 0) {
    return { http: 0, ok: false, msg: 'NO_ITEMS', itemsCount: 0, processed: 0, succeeded: 0, failed: 0, jsonStr: '' };
  }

  var currentToken = accessToken;
  var res;

  for (var attempt = 0; attempt < 2; attempt++) {
    var payload = {
      seller_id:   sellerId,
      yahoo_token: currentToken,
      items:       items,
      commit:      1
    };

    res = YahooUpload__postJsonDebug_(PACK_URL, payload, secret);

    var isTokenError = false;
    if (res.json && res.json.error === 'token_expired') {
      isTokenError = true;
    }

    if (isTokenError && attempt === 0) {
      YahooUpload__log_('トークンエラー検出 → 強制リフレッシュしてリトライ', 'SUBMIT');
      currentToken = YahooUpload__getValidAccessToken_(true);
      continue;
    }
    break;
  }

  var ok = (res.json && res.json.ok) ? true : false;
  var msg = YahooUpload__pickMsg_(res.json, res.body);
  var processed = 0, succeeded = 0, failed = 0;

  if (res.json && res.json.results) {
    var r = res.json.results;
    Object.keys(r).forEach(function(code) {
      processed++;
      if (r[code].ok) succeeded++; else failed++;
    });
  }

  var logMsg = 'HTTP=' + res.http + ' ok=' + ok + ' msg=' + msg
    + ' | items送信=' + items.length
    + ' | 処理=' + processed + '件'
    + ' | 成功=' + succeeded + '件'
    + ' | 失敗=' + failed + '件';

  YahooUpload__log_(logMsg, 'SUBMIT');
  YahooUpload__log_('submit応答JSON: ' + JSON.stringify(res.json, null, 2), 'SUBMIT');

  YahooUpload__logRun_(
    ss, runId, 'SUBMIT', 0, items.length, items.length,
    res.http, res.elapsed_ms, ok, logMsg,
    YahooUpload__stringifyForCell_(res.json || { raw: res.body }, cfg.LOG_JSON_MAXLEN)
  );

  return {
    http:       res.http,
    ok:         ok,
    msg:        msg,
    itemsCount: items.length,
    processed:  processed,
    succeeded:  succeeded,
    failed:     failed,
    jsonStr:    JSON.stringify(res.json, null, 2).substring(0, 500)
  };
}