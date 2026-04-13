/******************************************************
 * 00_config.gs - 共通設定
 ******************************************************/

// ===== シート名 =====
const SHEET_商品入力 = '①商品入力シート';
const SHEET_Yahoo商品登録 = 'Yahoo商品登録シート';

// ===== ヘッダー行数 =====
const HEADER_商品入力 = 2;
const HEADER_Yahoo商品登録 = 1;

// ===== 配送マスタ =====
const SHEET_配送グループ設定 = '配送グループ設定';
const SHEET_配送種別テキスト = '配送種別テキスト';

// ===== 列（固定運用） =====
const COL_CODE = 3;  // C列
const COL_NAME = 2;  // B列

// Yahoo用カラム（あなたの現行に合わせる）
const COL_ABSTRACT = 9;   // I列
const COL_SP_ADDITIONAL = 17; // Q列
const COL_POSTAGE_SET = 18;  // R列
const COL_SHIP_WEIGHT = 11;  // K列（佐川=1000）

// 画像（S/T固定）
const COL_S = 19; // S列 メイン
const COL_T = 20; // T列 詳細（;連結）

// ===== Driveフォルダ =====
const 商品画像ルートフォルダID = '1Vk1cVgrXM5CNFqfBvAaBUJpheDn4f9Lz';
const 免責画像ルートフォルダID = '1qfwT1NC3wwKgOLx2rvnSH836Zc0hROIa';

// ===== 画像サーバ（VPS） =====
const SERVER_HOST = 'img.niyantarose.com';
const SERVER_URL_BASE = 'https://img.niyantarose.com/img/';
const SSH_API_ENDPOINT = 'https://img.niyantarose.com/ssh_execute.php';
const IMG_CONVERT_SCRIPT = '/usr/local/bin/download_image_from_url.sh';

// 変換上限（固まり防止）
const MAX_CONVERT_IMAGES_PER_RUN = 50;
const MAX_CONVERT_MILLIS = 1000 * 60 * 5; // 5分

/**
 * ScriptProperties から設定値を取得（必須チェック付き）
 * 例: Yahoo画像__設定値取得_('Y_SELLER_ID')
 */
function Yahoo画像__設定値取得_(key) {
  key = String(key || '').trim();
  if (!key) throw new Error('設定キーが空です');

  var props = PropertiesService.getScriptProperties();
  var v = props.getProperty(key);

  if (v == null || String(v).trim() === '') {
    throw new Error(
      '設定値が未設定です: ' + key + '\n' +
      'メニュー「① 初期設定（client/secret/refresh/seller）」を実行して設定してな'
    );
  }
  return String(v).trim();
}

/**
 * アクセストークンのキャッシュをクリア
 * - refresh_token 生存確認 / 本体処理の前に呼ぶ
 */
function Yahoo画像__トークンキャッシュクリア_() {
  var props = PropertiesService.getScriptProperties();

  // よくあるキャッシュキー候補をまとめて消す（実装差異に強くする）
  var keys = [
    'Y_ACCESS_TOKEN',
    'Y_ACCESS_TOKEN_EXPIRES_AT',
    'Y_ACCESS_TOKEN_CREATED_AT',
    'Y_TOKEN_CACHE',
    'Y_TOKEN_CACHE_JSON',
    'Y_TOKEN_EXPIRES',
    'Y_TOKEN_EXPIRES_AT'
  ];

  keys.forEach(function(k) {
    try { props.deleteProperty(k); } catch (e) {}
  });

  Logger.log('🧹 トークンキャッシュをクリアしました');
}

/**
 * 文字列から URL だけを抽出して配列で返す
 * - セミコロン区切り / 空白 / 改行に対応
 * - http/https のみ対象
 */
function Yahoo画像__URLだけ抽出_配列_(text) {
  if (!text) return [];

  return String(text)
    .split(/[\s;]+/)
    .map(function(s){ return s.trim(); })
    .filter(function(s){ return /^https?:\/\//i.test(s); });
}
/**
 * Yahooの商品画像URLか？（S/Tに入ってるYahooの画像URL判定）
 * 例: https://item-shopping.c.yimg.jp/i/n/.... のような /i/ を含む形式を想定
 */
function Yahoo画像__Yahoo形式URLか_(url) {
  url = String(url || '').trim();
  if (!url) return false;

  // Yahooの画像URLはだいたい /i/ を含む（item-shopping.c.yimg.jp / shopping.c.yimg.jp 等）
  if (!/^https?:\/\//i.test(url)) return false;

  // 代表的ホストに寄せつつ、将来のホスト違いにも耐えるため /i/ パスで判定
  return /\/i\/[^\/]+\//i.test(url) || /item-shopping\.c\.yimg\.jp/i.test(url);
}
/**
 * 画像ソース収集
 * - 優先順:
 *   1) Drive: 画像ルート/code フォルダ内の画像（番号順）
 *   2) 既存S/T列のURL（http/httpsのみ）
 * - 免責キーワードに一致する画像は後ろへ回す
 *
 * @param {Folder} rootFolder 画像ルートフォルダ
 * @param {string} code 商品コード
 * @param {string} existingS 既存S列文字列
 * @param {string} existingT 既存T列文字列
 * @param {number} rowNum ログ用行番号
 * @return {Array} sources [{種類:'drive'|'url', ...}]
 */
function Yahoo画像__画像ソース収集_(rootFolder, code, existingS, existingT, rowNum) {
  code = String(code || '').trim();
  if (!code) return [];

  var cfg = (typeof Yahoo画像 !== 'undefined' && Yahoo画像.設定) ? Yahoo画像.設定 : {};
  var maxTotal = Number(cfg.最大画像枚数 || 21);
  var maxDetail = Number(cfg.最大詳細枚数 || 20);

  // 免責判定（設定が無い時も落ちないように）
  var keywords = (cfg.免責キーワード && cfg.免責キーワード.length) ? cfg.免責キーワード : ['notice_', '免責', 'menseki', 'disclaimer'];
  function isDisclaimer(nameOrUrl) {
    var s = String(nameOrUrl || '').toLowerCase();
    for (var i = 0; i < keywords.length; i++) {
      if (s.indexOf(String(keywords[i]).toLowerCase()) !== -1) return true;
    }
    return false;
  }

  // --- 1) Drive優先 ---
  var driveList = [];
  try {
    // 既にあなたのコード内に Yahoo画像__Drive画像取得_ があるので利用
    if (typeof Yahoo画像__Drive画像取得_ === 'function') {
      driveList = Yahoo画像__Drive画像取得_(rootFolder, code) || [];
    } else {
      // 万一無い場合の最低限フォールバック（codeフォルダ内の画像全部）
      var it = rootFolder.getFoldersByName(code);
      if (it.hasNext()) {
        var folder = it.next();
        var files = folder.getFiles();
        while (files.hasNext()) {
          var f = files.next();
          var name = String(f.getName() || '');
          if (!/\.(jpe?g|png|webp|gif)$/i.test(name)) continue;
          driveList.push({
            種類: 'drive',
            表示名: name,
            ファイル: f,
            番号: null,
            免責フラグ: isDisclaimer(name)
          });
        }
      }
    }
  } catch (e) {
    Logger.log('  [WARN] Drive画像取得で例外（行' + rowNum + ' ' + code + '）: ' + (e.message || e));
    driveList = [];
  }

  // Driveがあればそれを採用（免責は末尾）
  if (driveList && driveList.length) {
    var normal = [];
    var disc = [];
    for (var i = 0; i < driveList.length; i++) {
      (driveList[i].免責フラグ ? disc : normal).push(driveList[i]);
    }

    var out = normal.concat(disc).slice(0, maxTotal);
    Logger.log('  [SOURCE] Drive=' + out.length + '（行' + rowNum + ' ' + code + '）');
    return out;
  }

  // --- 2) URLフォールバック（S/TのURLを使う） ---
  var urls = [];
  // Sは1枚
  var sArr = (typeof Yahoo画像__URLだけ抽出_配列_ === 'function') ? Yahoo画像__URLだけ抽出_配列_(existingS) : [];
  if (sArr.length) urls.push(sArr[0]);

  // Tは最大詳細枚数まで
  var tArr = (typeof Yahoo画像__URLだけ抽出_配列_ === 'function') ? Yahoo画像__URLだけ抽出_配列_(existingT) : [];
  for (var j = 0; j < tArr.length && urls.length < (1 + maxDetail) && urls.length < maxTotal; j++) {
    urls.push(tArr[j]);
  }

  // URLが無いなら空
  if (!urls.length) {
    Logger.log('  [SOURCE] 画像ソースなし（DriveもURLも無し）行' + rowNum + ' ' + code);
    return [];
  }

  // sources化（免責URLは末尾）
  var urlSources = urls.map(function(u, idx){
    return {
      種類: 'url',
      表示名: 'url_' + (idx + 1),
      URL: u,
      番号: idx + 1,
      免責フラグ: isDisclaimer(u)
    };
  });

  var n2 = [], d2 = [];
  for (var k = 0; k < urlSources.length; k++) {
    (urlSources[k].免責フラグ ? d2 : n2).push(urlSources[k]);
  }

  var out2 = n2.concat(d2).slice(0, maxTotal);
  Logger.log('  [SOURCE] URL=' + out2.length + '（行' + rowNum + ' ' + code + '）');
  return out2;
}
/**
 * 免責画像か判定（ファイル名で判定）
 * Yahoo画像.設定.免責キーワード を見て判定
 */
function Yahoo画像__免責判定_(name) {
  var cfg = (typeof Yahoo画像 !== 'undefined' && Yahoo画像.設定) ? Yahoo画像.設定 : {};
  var keys = (cfg.免責キーワード && cfg.免責キーワード.length)
    ? cfg.免責キーワード
    : ['notice_', '免責', 'menseki', 'disclaimer'];

  var s = String(name || '').toLowerCase();
  for (var i = 0; i < keys.length; i++) {
    if (s.indexOf(String(keys[i]).toLowerCase()) !== -1) return true;
  }
  return false;
}
/**
 * 画像ソースを一意に識別するキーを返す（重複アップロード防止）
 * - drive: fileId を使う（最強）
 * - url  : URL を正規化して使う
 */
function Yahoo画像__ソース識別子取得_(src) {
  if (!src) return '';

  // Driveソース（Yahoo画像__Drive画像取得_ の形式）
  if (src.種類 === 'drive') {
    try {
      if (src.ファイル && typeof src.ファイル.getId === 'function') {
        return 'drive:' + src.ファイル.getId();
      }
    } catch (e) {}
    // フォールバック（ファイル名）
    return 'driveName:' + String(src.表示名 || '');
  }

  // URLソース（Yahoo画像__画像ソース収集_ の形式）
  if (src.種類 === 'url') {
    var u = String(src.URL || '').trim();
    if (!u) return '';
    // 小さく正規化（末尾スラッシュ/空白除去）
    u = u.replace(/\s+/g, '');
    return 'url:' + u;
  }

  // 予備
  return String(src.種類 || '') + ':' + String(src.表示名 || '');
}
/******************************************************
 * Yahoo画像 共通関数：アップロード〜反映（最小完結パック）
 ******************************************************/

/**
 * ソース（drive/url）からアップロード用Blobを作る
 * @param {Object} src {種類:'drive'|'url', ファイル or URL}
 * @param {string} uploadName Yahoo側でのファイル名（例: CODE.jpg）
 * @return {Blob}
 */
function Yahoo画像__Blob取得_(src, uploadName) {
  if (!src) throw new Error('Blob取得: srcが空');

  uploadName = String(uploadName || 'image.jpg');

  // Drive
  if (src.種類 === 'drive') {
    if (!src.ファイル || typeof src.ファイル.getBlob !== 'function') {
      throw new Error('Blob取得: Driveファイルが不正');
    }
    var blob = src.ファイル.getBlob();
    // contentTypeはDrive側に任せる。名前だけ付ける
    return blob.setName(uploadName);
  }

  // URL
  if (src.種類 === 'url') {
    var url = String(src.URL || '').trim();
    if (!/^https?:\/\//i.test(url)) throw new Error('Blob取得: URLが不正: ' + url);

    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    var http = res.getResponseCode();
    if (http !== 200) throw new Error('Blob取得: URL取得失敗 HTTP=' + http + ' url=' + url);

    var blob2 = res.getBlob();
    return blob2.setName(uploadName);
  }

  throw new Error('Blob取得: 未対応の種類=' + src.種類);
}

/** トークンキャッシュ削除（前に渡したやつが無い場合の保険） */
function Yahoo画像__トークンキャッシュクリア_() {
  var props = PropertiesService.getScriptProperties();
  ['Y_ACCESS_TOKEN','Y_ACCESS_TOKEN_EXPIRES_AT','Y_ACCESS_TOKEN_CREATED_AT'].forEach(function(k){
    try { props.deleteProperty(k); } catch(e) {}
  });
  Logger.log('🧹 トークンキャッシュをクリアしました');
}

/**
 * アクセストークン取得（refresh_token から自動更新）
 * ScriptProperties 必須:
 *  - Y_CLIENT_ID / Y_CLIENT_SECRET / Y_REFRESH_TOKEN
 * キャッシュ:
 *  - Y_ACCESS_TOKEN / Y_ACCESS_TOKEN_EXPIRES_AT
 */
function Yahoo画像__アクセストークン取得_() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('Y_ACCESS_TOKEN');
  var expAt = Number(props.getProperty('Y_ACCESS_TOKEN_EXPIRES_AT') || '0');

  // まだ有効ならキャッシュを返す（60秒マージン）
  if (token && expAt && Date.now() < (expAt - 60*1000)) {
    return token;
  }

  var clientId = props.getProperty('Y_CLIENT_ID');
  var clientSecret = props.getProperty('Y_CLIENT_SECRET');
  var refreshToken = props.getProperty('Y_REFRESH_TOKEN');

  if (!clientId || !clientSecret || !refreshToken) {
    throw new Error('トークン取得: client/secret/refresh が未設定。メニュー①を実行してな');
  }

  var url = 'https://auth.login.yahoo.co.jp/yconnect/v2/token';

  var payload = {
    grant_type: 'refresh_token',
    refresh_token: refreshToken
  };

  var basic = Utilities.base64Encode(clientId + ':' + clientSecret);

  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      Authorization: 'Basic ' + basic
    },
    payload: payload,
    muteHttpExceptions: true
  });

  var http = res.getResponseCode();
  var body = res.getContentText() || '';

  if (http !== 200) {
    throw new Error('トークン更新失敗 HTTP=' + http + ' body=' + body.substring(0, 500));
  }

  var json = JSON.parse(body);
  var access = String(json.access_token || '');
  var expiresIn = Number(json.expires_in || 0);

  if (!access) throw new Error('トークン更新: access_token が空');

  props.setProperty('Y_ACCESS_TOKEN', access);
  // expires_in は秒。マージンを引いて保存
  props.setProperty('Y_ACCESS_TOKEN_EXPIRES_AT', String(Date.now() + (Math.max(60, expiresIn - 60) * 1000)));

  return access;
}

/**
 * API実行：401等ならトークン取り直してもう一回
 * @param {function(string):HTTPResponse} fetcher accessTokenを受けてUrlFetchApp.fetchを返す関数
 * @return {HTTPResponse}
 */
function Yahoo画像__API実行_トークン自動更新_(fetcher) {
  var access = Yahoo画像__アクセストークン取得_();
  var res = fetcher(access);

  var http = res.getResponseCode();
  if (http === 401) {
    // トークン無効の可能性→キャッシュクリア→再取得→再実行
    Yahoo画像__トークンキャッシュクリア_();
    access = Yahoo画像__アクセストークン取得_();
    res = fetcher(access);
  }
  return res;
}

/**
 * 画像アップロード（並列）
 * - reqs: [{blob: Blob}, ...]
 * @return {string[]} Yahoo画像URL配列（同順）
 */
/**
 * 画像アップロード（確実版：1枚ずつfetch）
 * - reqs: [{blob: Blob}, ...]
 * @return {string[]} Yahoo画像URL配列（同順）
 */
/**
 * 画像アップロード（並列）
 * - seller_id は「GETパラメータで渡す」必要がある（bodyでは受け取られない）
 * - reqs: [{blob: Blob}, ...]
 * return: Yahoo画像URL配列（reqs順）
 */
/**
 * 画像アップロード（並列）
 * - seller_id は URL クエリで渡す（bodyでは受け取られない）
 * - reqs: [{blob: Blob}, ...]
 * return: Yahoo画像URL配列（reqs順）
 */
function Yahoo画像__画像アップロード_並列_(sellerId, reqs) {
  var uploadUrl =
    'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/uploadItemImage' +
    '?seller_id=' + encodeURIComponent(sellerId);

  // reqs: [{blob: Blob}, ...]
  var pending = reqs.map(function(r, idx){
    return { idx: idx, blob: r.blob };
  });

  var results = new Array(reqs.length);
  var maxRetry = 5;
  var baseWait = 1000;

  // ★ fetchAllで一度に投げる本数
  var MAX_FETCHALL = 8;

  // 最初のtoken
  var token = Yahoo画像__アクセストークン取得_();

  for (var t = 0; t <= maxRetry; t++) {
    if (pending.length === 0) break;

    var nextPending = [];
    var sawAuthExpire = false;

    // ★ pendingを MAX_FETCHALL 本ずつ処理
    for (var offset = 0; offset < pending.length; offset += MAX_FETCHALL) {
      var chunk = pending.slice(offset, offset + MAX_FETCHALL);

      // ★ GASのpayloadオブジェクトに任せる（multipart自動生成）
      var requests = chunk.map(function(p){
        var bytes = p.blob.getBytes();
        if (!bytes || bytes.length === 0) {
          throw new Error('Blobが0byteです idx=' + p.idx + ' name=' + (p.blob.getName ? p.blob.getName() : ''));
        }

        return {
          url: uploadUrl,
          method: 'post',
          headers: { Authorization: 'Bearer ' + token },
          payload: { file: p.blob },  // ★これだけでOK
          muteHttpExceptions: true
        };
      });

      var responses = UrlFetchApp.fetchAll(requests);

      for (var i = 0; i < responses.length; i++) {
        var res = responses[i];
        var p = chunk[i];

        var http = res.getResponseCode();
        var body = res.getContentText() || '';

        if (http === 200) {
          var yurl = Yahoo画像__XML解析_URLを取得_(body);
          if (!yurl) throw new Error('upload応答からURL取得失敗 body=' + body.substring(0, 200));
          results[p.idx] = yurl;
          continue;
        }

        // 認証切れ系
        var authExpired = (http === 401) ||
                          (body.indexOf('AccessToken has been expired') !== -1) ||
                          (body.indexOf('px-04102') !== -1);
        if (authExpired) {
          sawAuthExpire = true;
          nextPending.push(p);
          continue;
        }

        // im-02001（file必須）/ ed-00006 / 503 は再試行対象
        var isIm02001 = body.indexOf('im-02001') !== -1 || body.indexOf('fileは必須') !== -1;
        var isRetryable = (http === 503) || (http === 400 && Yahoo画像__isEd00006_(body)) || isIm02001;

        if (isRetryable) {
          nextPending.push(p);
          continue;
        }

        throw new Error('uploadItemImage失敗 HTTP=' + http + ' body=' + body.substring(0, 400));
      }
    }

    if (sawAuthExpire) {
      Yahoo画像__トークンキャッシュクリア_();
      token = Yahoo画像__アクセストークン取得_();
    }

    // ★ 単発送信で救済
    if (nextPending.length > 0) {
      var rescued = [];
      for (var k = 0; k < nextPending.length; k++) {
        var pp = nextPending[k];

        try {
          var res2 = UrlFetchApp.fetch(uploadUrl, {
            method: 'post',
            headers: { Authorization: 'Bearer ' + token },
            payload: { file: pp.blob },  // ★これだけでOK
            muteHttpExceptions: true
          });

          var h2 = res2.getResponseCode();
          var b2 = res2.getContentText() || '';

          if (h2 === 200) {
            var y2 = Yahoo画像__XML解析_URLを取得_(b2);
            if (!y2) throw new Error('単発upload URL取得失敗 body=' + b2.substring(0, 200));
            results[pp.idx] = y2;
            // 救済成功 → pendingから除外
          } else {
            rescued.push(pp);
          }
        } catch (e2) {
          rescued.push(pp);
        }
      }
      nextPending = rescued;
    }

    pending = nextPending;

    if (pending.length > 0) {
      var wait = baseWait * Math.pow(1.5, t);
      if (wait > 10000) wait = 10000;
      Utilities.sleep(wait);
    }
  }

  if (pending.length > 0) {
    throw new Error('uploadItemImage失敗（リトライ上限） 残=' + pending.length);
  }

  // 念のため：結果が埋まってないところがあれば落とす
  for (var z = 0; z < results.length; z++) {
    if (!results[z]) throw new Error('upload結果URLが不足（null/空） idx=' + z);
  }

  return results;
}

/**
 * submitItem（商品反映）
 */
function Yahoo画像__商品反映_(sellerId, code) {
  sellerId = String(sellerId || '').trim();
  code = String(code || '').trim();
  if (!sellerId || !code) throw new Error('submitItem: sellerId/codeが空');

  var url = 'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/submitItem';
  var payload = { seller_id: sellerId, item_code: code };

  var res = Yahoo画像__API実行_トークン自動更新_(function(access){
    return UrlFetchApp.fetch(url, {
      method: 'post',
      headers: { Authorization: 'Bearer ' + access },
      payload: payload,
      muteHttpExceptions: true
    });
  });

  var http = res.getResponseCode();
  var body = res.getContentText() || '';

  if (http !== 200) throw new Error('submitItem失敗 HTTP=' + http + ' body=' + body.substring(0, 500));
  if (body.indexOf('<Error>') !== -1) {
    // it-07004等を呼び出し側で握りつぶす設計なので、ここはエラー文字列をそのまま投げる
    throw new Error(body.substring(0, 500));
  }
  return { success: true, body: body };
}

function Yahoo画像__deleteItemImage_生レスポンス_(sellerId, imageIds) {
  var url = 'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/deleteItemImage';
  var payload = { seller_id: sellerId, image_id: imageIds.join(',') };

  var res = Yahoo画像__API実行_トークン自動更新_(function(accessToken){
    return UrlFetchApp.fetch(url, {
      method: 'post',
      headers: { Authorization: 'Bearer ' + accessToken },
      payload: payload,
      muteHttpExceptions: true
    });
  });

  var http = res.getResponseCode();
  var body = res.getContentText() || '';

  if (http === 200) {
    var ok = (body.match(/<Status>OK<\/Status>/g) || []).length;
    return { http: 200, okCount: ok, code: '', message: '', retryable: false, body: body };
  }

  // エラー解析（XML）
  var code = '';
  var message = '';
  try {
    var m1 = body.match(/<Code>([^<]+)<\/Code>/);
    if (m1) code = m1[1];
    var m2 = body.match(/<Message><!\[CDATA\[([\s\S]*?)\]\]><\/Message>/);
    if (m2) message = m2[1];
  } catch (e) {}

  // ★ ed-00006 は一時ロック＝リトライ対象
  var retryable = (code === 'ed-00006');

  return { http: http, okCount: 0, code: code, message: message, retryable: retryable, body: body };
}
function Yahoo画像__deleteItemImage_安全削除_ed00006耐性_(sellerId, imageIds) {
  imageIds = (imageIds || [])
    .map(String)
    .map(function(s){ return s.trim(); })
    .filter(Boolean);

  if (!imageIds.length) return { ok: 0, badIds: [], tried: 0 };

  var totalOk = 0;
  var badIds = [];
  var tried = 0;

  var batchSize = 100;
  var maxRetry = 6; // ed-00006多発時に少し粘る

  for (var p = 0; p < imageIds.length; p += batchSize) {
    var batch = imageIds.slice(p, p + batchSize);
    tried += batch.length;

    // まずバッチで試す → ed-00006なら待ってリトライ
    var attempt = 0;
    while (attempt <= maxRetry) {
      var r = Yahoo画像__deleteItemImage_生レスポンス_(sellerId, batch);

      if (r.http === 200) {
        totalOk += r.okCount;
        break;
      }

      if (r.retryable) {
        var wait = Math.min(8000, 1500 * Math.pow(1.6, attempt)); // 1.5s→最大8s
        Logger.log('      [LOCK] ed-00006 のため待機 ' + wait + 'ms → リトライ ' + (attempt+1) + '/' + maxRetry);
        Utilities.sleep(wait);
        attempt++;
        continue;
      }

      // ed-00006以外は 1件ずつ切り分け
      Logger.log('      [WARN] deleteItemImage バッチ失敗 HTTP=' + r.http + ' code=' + (r.code||'') + ' msg=' + (r.message||''));
      for (var i = 0; i < batch.length; i++) {
        var id1 = batch[i];

        var a2 = 0;
        var ok1 = false;

        // 1件でも ed-00006なら待って粘る（ただし“badIds”には入れない）
        while (a2 <= maxRetry) {
          var r1 = Yahoo画像__deleteItemImage_生レスポンス_(sellerId, [id1]);
          if (r1.http === 200) { totalOk += r1.okCount; ok1 = true; break; }

          if (r1.retryable) {
            Utilities.sleep(Math.min(8000, 1500 * Math.pow(1.6, a2)));
            a2++;
            continue;
          }

          // それ以外は不正扱い
          badIds.push(id1);
          ok1 = true; // 終了
          break;
        }

        // ロックが解けずに尽きた場合：保留（badIdsにしない）
        if (!ok1) {
          Logger.log('        [HOLD] ロックが解けず削除保留: ' + id1);
        }
      }

      break; // 切り分けに入ったら次バッチへ
    }
  }

  return { ok: totalOk, badIds: badIds, tried: tried };
}

function runOnSheet_(sheetName, fn) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const current = ss.getActiveSheet();            // いまのシート（戻す用）
  const target = ss.getSheetByName(sheetName);
  if (!target) throw new Error('シートが見つかりません: ' + sheetName);

  ss.setActiveSheet(target);                      // 対象シートに切り替え
  try {
    fn();                                         // 実行
  } finally {
    ss.setActiveSheet(current);                   // 元に戻す
  }
}

function IMG_alert_(msg) {
  msg = String(msg || '');

  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    Logger.log('UI alert skipped (no UI): ' + msg + ' / ' + e);
  }
}



