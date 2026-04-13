/* ───── Yahoo! API 定数  ★自分の値をプロパティに設定 ───── */
const Y_SELLER_ID     = 'niyantarose';  // ★ ストアアカウント
const Y_CLIENT_ID     = PropertiesService.getScriptProperties().getProperty('Y_CLIENT_ID');
const Y_CLIENT_SECRET = PropertiesService.getScriptProperties().getProperty('Y_CLIENT_SECRET');
const Y_REFRESH_TOKEN = PropertiesService.getScriptProperties().getProperty('Y_REFRESH_TOKEN');
const Y_TOKEN_URL     = 'https://auth.login.yahoo.co.jp/yconnect/v2/token';
const Y_SETSTOCK_URL  = 'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/setStock';

function refreshYahooToken() {
  const payload = {
    grant_type   : 'refresh_token',
    refresh_token: Y_REFRESH_TOKEN,
    client_id    : Y_CLIENT_ID,
    client_secret: Y_CLIENT_SECRET
  };
  const res = UrlFetchApp.fetch(Y_TOKEN_URL, {method:'post', payload:payload});
  const json = JSON.parse(res.getContentText());
  PropertiesService.getScriptProperties().setProperty('Y_ACCESS_TOKEN', json.access_token);
  return json.access_token;
}

function pushStockToYahoo() {
  const token = PropertiesService.getScriptProperties().getProperty('Y_ACCESS_TOKEN') || refreshYahooToken();

  // 直近5分の更新分を取得
  const since = new Date(Date.now() - 5 * 60 * 1000).toISOString();
  const where = encodeURIComponent(`LastUpdated,gt,${since}`);
  const url   = `${NOCO_ENDPOINT}?where=${where}`;
  const list  = JSON.parse(UrlFetchApp.fetch(url, {headers:{'xc-auth':NOCO_TOKEN}})).list;

  if (!list.length) return;

  const itemCodes  = list.map(r=>r.ProductID).join(',');
  const quantities = list.map(r=>r.Quantity ).join(',');

  const headers = {
    'Authorization':'Bearer ' + token,
    'Content-Type' :'application/x-www-form-urlencoded'
  };
  const params = {
    seller_id : Y_SELLER_ID,
    item_code : itemCodes,
    quantity  : quantities
  };

  const res = UrlFetchApp.fetch(Y_SETSTOCK_URL, {method:'post', headers:headers, payload:params});
  Logger.log(res.getContentText());
}
