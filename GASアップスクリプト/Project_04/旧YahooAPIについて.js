// Yahoo Shopping API設定
const YAHOO設定 = {
  クライアントID: 'dj00aiZpPUQ4TEx2bDllVXl2eCZzPWNvbnNlbWVyc2VjcmV0Jng9YTM-',
  クライアントシークレット: 'uLjglH91MOAljwDwR7tGVubXa1UJ54fsEYTpdbii',
  トークンURL: 'https://auth.login.yahoo.co.jp/yconnect/v2/token'
};

/**
 * 現在のアクセストークンを確認（ログ出力）
 */
function 現在のアクセストークンを確認() {
  const props = PropertiesService.getScriptProperties();
  const accessToken = props.getProperty('YAHOO_ACCESS_TOKEN');
  const refreshToken = props.getProperty('YAHOO_REFRESH_TOKEN');
  
  console.log('=== Yahoo トークン情報 ===');
  console.log('アクセストークン: ' + (accessToken || '未設定'));
  console.log('リフレッシュトークン: ' + (refreshToken ? refreshToken.substring(0, 30) + '...' : '未設定'));
  
  return accessToken;
}

/**
 * リフレッシュトークンを初期設定
 */
function リフレッシュトークンを初期設定() {
  const refreshToken = 'A8ObS2kCAIRmiHxmgoAOUSD08f0nilpyBO_u7nepDyh_obWo8uKnGIGRcXma_arfchQg5ofl6lKz54ZZjtcGAX_REyJESbvmvOEpvbdX9Fw6QTtzfEe0q-W31P2-_eefeRZDytrZ1weD4YhwbJnJXB3HwT6TK1CNbQpQbIK6SRGtlBlU~1';
  PropertiesService.getScriptProperties().setProperty('YAHOO_REFRESH_TOKEN', refreshToken);
  console.log('リフレッシュトークンを設定しました');
}

/**
 * アクセストークンを更新して取得
 */
function アクセストークンを更新して取得() {
  const props = PropertiesService.getScriptProperties();
  const refreshToken = props.getProperty('YAHOO_REFRESH_TOKEN');
  
  if (!refreshToken) {
    throw new Error('YAHOO_REFRESH_TOKEN が未設定です。先に「リフレッシュトークンを初期設定」を実行してください');
  }
  
  const credentials = Utilities.base64Encode(YAHOO設定.クライアントID + ':' + YAHOO設定.クライアントシークレット);
  
  const response = UrlFetchApp.fetch(YAHOO設定.トークンURL, {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    headers: {
      'Authorization': 'Basic ' + credentials
    },
    payload: {
      'grant_type': 'refresh_token',
      'refresh_token': refreshToken
    },
    muteHttpExceptions: true
  });
  
  const http = response.getResponseCode();
  const text = response.getContentText();
  
  console.log('HTTP: ' + http);
  console.log('レスポンス: ' + text);
  
  if (http !== 200) {
    throw new Error('トークン更新失敗 HTTP=' + http + '\n' + text);
  }
  
  const json = JSON.parse(text);
  
  props.setProperty('YAHOO_ACCESS_TOKEN', json.access_token);
  if (json.refresh_token) {
    props.setProperty('YAHOO_REFRESH_TOKEN', json.refresh_token);
    console.log('リフレッシュトークンも更新されました');
  }
  
  console.log('=== 新しいアクセストークン ===');
  console.log(json.access_token);
  
  return json.access_token;
}

/**
 * アクセストークンを取得（なければ自動更新）
 */
function アクセストークンを取得() {
  const props = PropertiesService.getScriptProperties();
  let accessToken = props.getProperty('YAHOO_ACCESS_TOKEN');
  
  if (!accessToken) {
    console.log('アクセストークンがないため更新します');
    accessToken = アクセストークンを更新して取得();
  }
  
  return accessToken;
}

/**
 * トークンをクリア
 */
function トークンをクリア() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('YAHOO_ACCESS_TOKEN');
  props.deleteProperty('YAHOO_REFRESH_TOKEN');
  console.log('トークン情報をクリアしました');
}