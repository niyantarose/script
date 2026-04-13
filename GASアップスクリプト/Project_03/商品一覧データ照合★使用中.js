  // ========================== Yahoo! CSV 自動ダウンロード & シート更新 ==========================
  // 2025-07-04  EU 403 抑制メール版
  //   ・EU/EEA ブロック（GDPR ページ）による失敗ではメール送信しない
  //   ・runAllProcesses 全体失敗／トリガー生成失敗／トークン更新失敗のみ通知
  // 2025-07-24  商品コード自動追加機能追加
  //   ・CSV在庫データから商品一覧シートに存在しない商品コードを自動追加
  // 2025-07-30  リフレッシュトークンエラー修正
  //   ・refresh_token パラメータの不足エラーを修正
  //   ・エラーハンドリングを改善
  // 2025-07-30  コールバックURL修正
  //   ・登録済みのコールバックURLに合わせて修正
  // ---------------------------------------------------------------------------------------------

  // ---------- 定数 ----------
  var CLIENT_ID         = 'dj00aiZpPUQ4TEx2bDllVXl2eCZzPWNvbnN1bWVyc2VjcmV0Jng9YTM-';
  var CLIENT_SECRET     = 'uLjglH91MOAljwDwR7tGVubXa1UJ54fsEYTpdbii';
  var SELLER_ID         = 'niyantarose';
  var ALERT_EMAIL       = 'niyantarose2@gmail.com';
  var CALLBACK_URL      = 'https://script.google.com/macros/s/AKfycbyEPYRYq2SebdFWYD3qfxvTF4LxJqbTNPLBP_Y0M7sbQIYGGQ5O0HTd0wqiVlbUhPQi/exec';

  var PRODUCT_CSV_TYPE  = 1;   // 商品
  var STOCK_CSV_TYPE    = 2;   // 在庫

  var MAX_RETRIES       = 12;
  var INITIAL_DELAY_MS  = 1000;
  var MAX_DELAY_MS      = 8000;

  var MIN_REQUEST_INTERVAL_MS = 30000;      // 30 秒ルール
  var RATE_LIMIT_PROPERTY_KEY = 'LAST_REQUEST_TIME';

  var TOKEN_URL         = 'https://auth.login.yahoo.co.jp/yconnect/v2/token';

  // ---------- 共通ユーティリティ ----------
  function log(msg){ Logger.log(msg); }
  function sendAlert(sub, body){
    try{ MailApp.sendEmail(ALERT_EMAIL, sub, body); }
    catch(e){ log('メール送信失敗:'+e.message); }
  }
  // 特定の場面だけ呼ぶ
  function logAndAlert(title, body){ log(title+'\n'+body); sendAlert(title, body); }

  // ---------- 認可関連 ----------
  function getValidAccessToken(){
    var p       = PropertiesService.getScriptProperties();
    var token   = p.getProperty('ACCESS_TOKEN');
    var expires = Number(p.getProperty('TOKEN_EXPIRES_AT')||0);
    if(token && Date.now()<expires-5*60*1000) return token;
    
    var refreshToken = p.getProperty('REFRESH_TOKEN');
    if(!refreshToken || refreshToken.trim() === '') {
      throw new Error('REFRESH_TOKEN が設定されていません。初回認証を実行してください。');
    }
    
    return refreshAccessToken(refreshToken);
  }

  // ---------- 認可関連：リフレッシュトークンでアクセストークンを更新 ----------
function refreshAccessToken(rt) {
  /* 0️⃣ バリデーション --------------------------------------------------- */
  if (!rt || rt.trim() === '') {
    throw new Error('REFRESH_TOKEN が空です。最初に generateAuthUrl() で認証してください');
  }

  /* 1️⃣ リフレッシュトークンを使って再取得 ------------------------------ */
  var resp = UrlFetchApp.fetch(TOKEN_URL, {
    method        : 'post',
    contentType   : 'application/x-www-form-urlencoded',
    muteHttpExceptions: true,
    payload: {
      grant_type    : 'refresh_token',
      refresh_token : rt,
      client_id     : CLIENT_ID,
      client_secret : CLIENT_SECRET
    }
  });

  var code = resp.getResponseCode();
  var body = resp.getContentText();
  log('refreshAccessToken → HTTP ' + code);

  /* 2️⃣ エラー時のハンドリング ----------------------------------------- */
  if (code !== 200) {
    // 400＝期限切れなど → 再認証 URL をメールで送る
    if (code === 400) {
      var url = generateAuthUrl();
      sendAlert('【Yahoo トークン失効】再認証が必要です',
        '下記 URL を開いて認証コードを取得し\n' +
        'setAuthorizationCode("コード") を実行してください\n\n' + url +
        '\n\nAPI応答：\n' + body);
    }
    throw new Error('refreshAccessToken 失敗 (' + code + '): ' + body);
  }

  /* 3️⃣ トークン保存 ----------------------------------------------------- */
  var j = JSON.parse(body);
  if (!j.access_token) { throw new Error('access_token が取得できません: ' + body); }

  var props = {
    ACCESS_TOKEN     : j.access_token,
    TOKEN_EXPIRES_AT : String(Date.now() + (j.expires_in || 3600) * 1000) // ms
  };
  if (j.refresh_token) { props.REFRESH_TOKEN = j.refresh_token; } // 更新される時だけ

  PropertiesService.getScriptProperties().setProperties(props);
  log('✅ アクセストークン更新完了　有効期限: ' + (j.expires_in || 3600) + ' 秒');

  return j.access_token;
}

  // ---------- リクエスト間隔チェック ----------
  function checkRequestInterval(){
    var prop = PropertiesService.getScriptProperties();
    var last = Number(prop.getProperty(RATE_LIMIT_PROPERTY_KEY)||0);
    var now  = Date.now();
    if(now-last<MIN_REQUEST_INTERVAL_MS){
      Utilities.sleep(MIN_REQUEST_INTERVAL_MS-(now-last));
    }
    prop.setProperty(RATE_LIMIT_PROPERTY_KEY,String(Date.now()));
  }

  // ---------- CSV取得 ----------
  function fetchCsv(type){
    log('📥 CSV取得開始: type='+type);
    if(type!==PRODUCT_CSV_TYPE && type!==STOCK_CSV_TYPE)
      throw new Error('無効なtype:'+type);

    checkRequestInterval();
    var token=getValidAccessToken();

    // downloadRequest
    var req=UrlFetchApp.fetch(
      'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/downloadRequest',
      { method:'post',
        contentType:'application/x-www-form-urlencoded',
        headers:{ Authorization:'Bearer '+token,
                  'User-Agent':'GoogleAppsScript; Yahoo AppID: '+CLIENT_ID },
        payload:'seller_id='+SELLER_ID+'&type='+type });

    if(req.getResponseCode()!==200){
      log('downloadRequest HTTP:'+req.getResponseCode());
      throw new Error('downloadRequest failed');
    }

    var xml  = XmlService.parse(req.getContentText()).getRootElement();
    var fileKey = xml.getChild('Result').getChildText('FileKey');
    if(fileKey) log('file_key='+fileKey); else log('file_key missing');

    // downloadSubmit
    var submitURL = 'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/downloadSubmit'
                    +'?seller_id='+SELLER_ID+'&type='+type+(fileKey?'&file_key='+fileKey:'');
    var delay=INITIAL_DELAY_MS;
    for(var i=0;i<MAX_RETRIES;i++){
      var resp = UrlFetchApp.fetch(submitURL,{
        headers:{ Authorization:'Bearer '+token,
                  'User-Agent':'GoogleAppsScript; Yahoo AppID: '+CLIENT_ID },
        muteHttpExceptions:true});
      var code = resp.getResponseCode();
      if(code===200){ log('CSV取得成功'); return resp.getBlob().getDataAsString('Shift_JIS'); }
      if(code===401){ 
        log('認証エラー（401）- トークンを再取得します');
        token=refreshAccessToken(PropertiesService.getScriptProperties().getProperty('REFRESH_TOKEN')); 
        continue; 
      }
      if([202,403,500].includes(code)){
        log('リトライ中... ('+code+') 待機時間: '+delay+'ms');
        Utilities.sleep(delay); delay=Math.min(delay*1.4,MAX_DELAY_MS); continue;
      }
      throw new Error('downloadSubmit '+code);
    }
    throw new Error('downloadSubmit retry exceeded');
  }

  // ---------- シート更新 ----------
  function updateProductSheet() {
    log('📊 CSV商品データ 更新開始');
    var data = Utilities.parseCsv(fetchCsv(PRODUCT_CSV_TYPE));
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CSV商品データ');
    if (!sheet) throw new Error('CSV商品データ sheet missing');
    sheet.clear();
    if(data.length){
      sheet.getRange(1, 1, data.length, data[0].length)
          .setNumberFormat('@').setValues(data);
    }
    log('📊 CSV商品データ 更新完了 '+data.length+'行');
  }

  function updateStockSheet() {
    log('📊 CSV在庫データ 更新開始');
    var data = Utilities.parseCsv(fetchCsv(STOCK_CSV_TYPE));
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CSV在庫データ');
    if (!sheet) throw new Error('CSV在庫データ sheet missing');
    sheet.clear();
    if(data.length){
      sheet.getRange(1, 1, data.length, data[0].length)
          .setNumberFormat('@').setValues(data);
    }
    log('📊 CSV在庫データ 更新完了 '+data.length+'行');
  }

  // ---------- 商品一覧更新＋JAN反映 ----------
  function updateProductList() {
    log('📋 商品一覧更新開始');
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var prod = ss.getSheetByName('商品一覧');
    var inv  = ss.getSheetByName('CSV在庫データ');
    var jan  = ss.getSheetByName('CSV商品データ');
    
    if (!prod || !inv || !jan) {
      log('必要なシートが見つかりません');
      return;
    }
    
    var h = 1; // ヘッダー行数
    
    // invData: [プレフィックス, サフィックス候補, 数量]
    var invData = inv.getRange(h + 1, 1, Math.max(1, inv.getLastRow() - h), 3).getValues();
    var rows = invData
      .map(function(r) {
        var prefix   = String(r[0] || '').trim();
        var suffixCsv = String(r[1] || '').trim();
        var qty      = Number(r[2]) || 0;
        
        // 出力するコードはサフィックス候補があればそれ、なければプレフィックス
        var code     = suffixCsv || prefix;
        if(!code) return null; // 空のコードは除外
        
        // 末尾が小文字 a ～ z ならサフィックス、それ以外はサフィックス無し
        var m        = code.match(/([a-z])$/);
        var suf      = m ? m[1] : null;
        
        // ステータス判定
        var stat;
        if (qty === 0) {
          stat = '欠品';
        } else if (suf === 'a') {
          stat = '即納';
        } else if (suf === 'b') {
          stat = 'お取り寄せ';
        } else if (suf && suf >= 'c') {
          stat = qty <= 20 ? '即納' : 'お取り寄せ';
        } else {
          // 大文字サフィックスやサフィックス無しは在庫あれば即納
          stat = '即納';
        }
        
        return { code: code, stat: stat, qty: qty };
      })
      .filter(function(o) { return o && o.code; }); // null と空コードを除外
    
    var n = rows.length;
    if(n === 0) {
      log('更新対象の商品データがありません');
      return;
    }
    
    // A列: 商品ID
    var A = rows.map(function(row, i) {
      return ['PROD-' + Utilities.formatString('%07d', i + 1)];
    });
    
    // B列: 商品コード
    var B = rows.map(function(r) { return [r.code]; });
    
    // C列: ステータス
    var C = rows.map(function(r) { return [r.stat]; });
    
    // D列: JANコード（CSV商品データの10列目から取得）
    var janVals;
    try {
      janVals = jan.getRange(h + 1, 10, n, 1).getValues();
    } catch(e) {
      log('JANコード取得エラー: ' + e.message);
      janVals = rows.map(function() { return ['']; }); // 空で埋める
    }
    
    // F列: 数量
    var F = rows.map(function(r) { return [r.qty]; });
    
    // 既存データをクリア
    if (prod.getLastRow() > h) {
      prod.getRange(h + 1, 1, prod.getLastRow() - h, 6).clear();
    }
    
    // データを設定
    prod.getRange(h + 1, 1, n, 1).setValues(A);                    // A列: 商品ID
    var bRange = prod.getRange(h + 1, 2, n, 1);
    bRange.setNumberFormat('@').setValues(B);                       // B列: 商品コード
    prod.getRange(h + 1, 3, n, 1).setValues(C);                    // C列: ステータス
    prod.getRange(h + 1, 4, n, 1).setValues(janVals);              // D列: JANコード
    prod.getRange(h + 1, 6, n, 1).setValues(F);                    // F列: 数量
    
    log('📋 商品一覧更新完了 '+n+'行');
  }

  // ---------- バッチ実行 ----------
  function runAllProcesses(){
    log('🚀 runAllProcesses 開始');
    var okProd=true, okStock=true, okProductList=true;

    try{ updateProductSheet(); }catch(e){ okProd=false; log('CSV商品データ更新失敗:'+e.message); }
    try{ updateStockSheet();   }catch(e){ okStock=false; log('CSV在庫データ更新失敗:'+e.message); }
    
    // 商品一覧更新処理を実行
    try{ updateProductList(); }catch(e){ okProductList=false; log('商品一覧更新失敗:'+e.message); }

    var summary='CSV商品:'+(okProd?'OK':'NG')+' CSV在庫:'+(okStock?'OK':'NG')+' 商品一覧:'+(okProductList?'OK':'NG');
    log('🏁 runAllProcesses 終了 '+summary);

    // EU/EEA ブロック（GDPR ページ）はメール抑制
    var logText = Logger.getLog();
    var euBlock = logText.indexOf('欧州経済領域')>-1 ||
                  logText.indexOf('<title>【お知らせ】欧州経済領域')>-1;
    if(!euBlock && (!okProd || !okStock || !okProductList)){
      sendAlert('runAllProcesses 失敗', summary);
    }
  }

  // ---------- トリガー ----------
  function createCsvTriggers(){
    try{
      ScriptApp.getProjectTriggers().forEach(function(t){ ScriptApp.deleteTrigger(t); });
      ScriptApp.newTrigger('runAllProcesses')
              .timeBased()
              .everyMinutes(30)  // 30分毎に実行
              .create();
      log('🕒 トリガー設定完了');
    }catch(err){
      logAndAlert('createCsvTriggers 失敗', err.message);
      throw err;
    }
  }

  // ---------- 手動テスト ----------
  function manualTest(){ runAllProcesses(); }

  // ---------- 個別テスト用関数 ----------
  function testUpdateProductList(){ updateProductList(); }

  // ---------- デバッグ用関数 ----------
  function checkTokenStatus(){
    var p = PropertiesService.getScriptProperties();
    var accessToken = p.getProperty('ACCESS_TOKEN');
    var refreshToken = p.getProperty('REFRESH_TOKEN');
    var expiresAt = p.getProperty('TOKEN_EXPIRES_AT');
    
    log('=== トークン状態確認 ===');
    log('ACCESS_TOKEN: ' + (accessToken ? '設定済み (長さ: ' + accessToken.length + ')' : '未設定'));
    log('REFRESH_TOKEN: ' + (refreshToken ? '設定済み (長さ: ' + refreshToken.length + ')' : '未設定'));
    log('TOKEN_EXPIRES_AT: ' + (expiresAt ? new Date(Number(expiresAt)).toString() : '未設定'));
    log('現在時刻: ' + new Date().toString());
    log('登録済みコールバックURL: ' + CALLBACK_URL);
    
    if(!refreshToken) {
      log('⚠️ REFRESH_TOKEN が設定されていません。初回認証を実行してください。');
    }
  }

/* ---------- 初回認証用 URL を発行 ---------- */
function generateAuthUrl() {
  var authUrl =
      'https://auth.login.yahoo.co.jp/yconnect/v2/authorization' +
      '?response_type=code' +
      '&client_id='      + encodeURIComponent(CLIENT_ID) +
      '&redirect_uri='   + encodeURIComponent(CALLBACK_URL) +
      '&scope='          + encodeURIComponent('openid') +
      '&state='          + Utilities.getUuid() +
      // ↓↓↓ ★リフレッシュトークンを長期発行してもらうための追加パラメータ
      '&access_type=offline' +      // 長期 refresh_token を要求
      '&prompt=consent';           // 毎回同意画面を出す（確実に発行させる）

  log('=== 認証が必要です ===');
  log('1. この URL をブラウザで開く:\n' + authUrl);
  log('2. Yahoo! にログインして「同意」');
  log('3. リダイレクト後に表示される認証コードをコピー');
  log('4. setAuthorizationCode("H49JPjKa") を実行');
  return authUrl;
}

/* ---------- 認証コードを渡してトークンを取得 ---------- */
function setAuthorizationCode(authCode) {        // ★authCode は変数
  if (!authCode) {
    log('使用例:  setAuthorizationCode("xxxxxx")');
    return;
  }

  log('認証コードでトークンを取得中…');

  var response = UrlFetchApp.fetch(TOKEN_URL, {
    method             : 'post',
    contentType        : 'application/x-www-form-urlencoded',
    muteHttpExceptions : true,
    payload : {
      grant_type    : 'authorization_code',
      code          : authCode,
      redirect_uri  : CALLBACK_URL,
      client_id     : CLIENT_ID,
      client_secret : CLIENT_SECRET
    }
  });

  var code = response.getResponseCode();
  var body = response.getContentText();
  log('認証レスポンス: HTTP ' + code);

  if (code !== 200) {
    log('❌ 認証失敗: ' + body);
    return;
  }

  var j = JSON.parse(body);
  if (!j.access_token || !j.refresh_token) {
    throw new Error('access_token / refresh_token が取得できません:\n' + body);
  }

  PropertiesService.getScriptProperties().setProperties({
    ACCESS_TOKEN     : j.access_token,
    REFRESH_TOKEN    : j.refresh_token,
    TOKEN_EXPIRES_AT : String(Date.now() + (j.expires_in || 3600) * 1000)
  });

  log('✅ 認証成功！トークンを保存しました');
  log('これで manualTest() を実行して動作確認してください。');
}
// ---------- 認証ワンショット（新コード反映） ----------
function authOnce() {
  setAuthorizationCode("nYR7x57J");  // ← スクショの新しい認可コード
}



  // ---------- 簡単な認証状態リセット ----------
  function resetTokens() {
    PropertiesService.getScriptProperties().deleteProperty('ACCESS_TOKEN');
    PropertiesService.getScriptProperties().deleteProperty('REFRESH_TOKEN');
    PropertiesService.getScriptProperties().deleteProperty('TOKEN_EXPIRES_AT');
    log('全てのトークンをリセットしました。generateAuthUrl() から再開してください。');
  }

  // ---------- Web App用コールバック処理 ----------
  function doGet(e) {
    var code = e.parameter.code;
    var state = e.parameter.state;
    var error = e.parameter.error;
    
    if (error) {
      return HtmlService.createHtmlOutput('認証エラー: ' + error);
    }
    
    if (code) {
      return HtmlService.createHtmlOutput(
        '<h2>認証コード取得完了</h2>' +
        '<p>以下の認証コードをGoogle Apps Scriptで使用してください：</p>' +
        '<code style="background-color: #f0f0f0; padding: 10px; display: block; margin: 10px 0; font-size: 16px;">' + code + '</code>' +
        '<p>Google Apps Scriptで以下を実行してください：</p>' +
        '<code>setAuthorizationCode("' + code + '")</code>'
      );
    }
    
    return HtmlService.createHtmlOutput('認証に失敗しました。再度お試しください。');
  }