// ==================================================
// 設定・定数
// ==================================================
const Y_TOKEN_URL = 'https://auth.login.yahoo.co.jp/yconnect/v2/token';

// Script Properties のキー名
const KEY_CLIENT_ID            = 'Y_CLIENT_ID';
const KEY_CLIENT_SECRET        = 'Y_CLIENT_SECRET';
const KEY_REFRESH_TOKEN        = 'Y_REFRESH_TOKEN';
const KEY_ACCESS_TOKEN         = 'Y_ACCESS_TOKEN';
const KEY_EXPIRES_AT           = 'Y_ACCESS_EXPIRES_AT';
const KEY_REFRESH_SET_AT       = 'Y_REFRESH_SET_AT';

// 必要に応じてここだけ自分の値にして使う
const DEFAULT_Y_CLIENT_ID      = 'dj00aiZpPUQ4TEx2bDllVXl2eCZzPWNvbnN1bWVyc2VjcmV0Jng9YTM-';
const DEFAULT_Y_CLIENT_SECRET  = 'uLjglH91MOAljwDwR7tGVubXa1UJ54fsEYTpdbii';

// リフレッシュトークンの目安日数
const REFRESH_TOKEN_LIFETIME_DAYS = 28;


// ==================================================
// 1. メインのAPI呼び出し関数
// ==================================================
function fetchYahooApi(url, options = {}) {
  try {
    let accessToken = getValidAccessToken_();

    if (!options.headers) options.headers = {};
    options.headers['Authorization'] = `Bearer ${accessToken}`;
    options.muteHttpExceptions = true;

    let res = UrlFetchApp.fetch(url, options);
    let code = res.getResponseCode();

    // 期限内でも401が返るケースの救済
    if (code === 401) {
      console.warn('401エラー検知: トークンを強制更新して再試行します');
      accessToken = refreshAccessToken_();
      options.headers['Authorization'] = `Bearer ${accessToken}`;
      res = UrlFetchApp.fetch(url, options);
      code = res.getResponseCode();
    }

    if (code < 200 || code >= 300) {
      throw new Error(`Yahoo API Error: HTTP ${code} - ${res.getContentText()}`);
    }

    return res.getContentText();

  } catch (e) {
    console.error(e);
    Logger.log(String(e));
    try {
      const ss = SpreadsheetApp.getActive();
      if (ss) ss.toast(String(e).slice(0, 100), 'Yahoo APIエラー', 10);
    } catch (_) {}
    throw e;
  }
}


// ==================================================
// 2. 内部処理用関数
// ==================================================
function getValidAccessToken_() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty(KEY_ACCESS_TOKEN);
  const expiresAt = Number(props.getProperty(KEY_EXPIRES_AT) || 0);

  const needRefresh = !token || !expiresAt || Date.now() > (expiresAt - 60 * 1000);

  if (needRefresh) {
    return refreshAccessToken_();
  }
  return token;
}

function refreshAccessToken_() {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30 * 1000);
  } catch (e) {
    throw new Error('トークン更新のロック取得に失敗しました（同時実行が多い可能性があります）: ' + e);
  }

  try {
    const props = PropertiesService.getScriptProperties();

    // ロック待ちの間に他が更新したか確認
    const token = props.getProperty(KEY_ACCESS_TOKEN);
    const expiresAt = Number(props.getProperty(KEY_EXPIRES_AT) || 0);
    if (token && expiresAt && Date.now() <= (expiresAt - 60 * 1000)) {
      return token;
    }

    const clientId = props.getProperty(KEY_CLIENT_ID);
    const clientSecret = props.getProperty(KEY_CLIENT_SECRET);
    const refreshToken = props.getProperty(KEY_REFRESH_TOKEN);

    if (!clientId || !clientSecret || !refreshToken) {
      throw new Error(
        '認証情報不足です。\n\n' +
        '先に「Yahoo認証_初期設定_クライアント情報保存」を実行して CLIENT_ID / CLIENT_SECRET を保存し、\n' +
        'その後「Yahoo認証_ボタン実行」で refresh token を設定してください。'
      );
    }

    const credentials = Utilities.base64Encode(`${clientId}:${clientSecret}`);

    const res = UrlFetchApp.fetch(Y_TOKEN_URL, {
      method: 'post',
      contentType: 'application/x-www-form-urlencoded',
      headers: { Authorization: `Basic ${credentials}` },
      payload: {
        grant_type: 'refresh_token',
        refresh_token: refreshToken,
      },
      muteHttpExceptions: true,
    });

    const httpCode = res.getResponseCode();
    const text = res.getContentText();

    if (httpCode !== 200) {
      let msg = `トークン更新失敗 HTTP=${httpCode}\n${text}`;

      if (text.includes('invalid_grant')) {
        msg += '\n\n【重要】原因: refresh_token が失効または無効です。' +
               '\nブラウザで再認可を行い、「Yahoo認証_ボタン実行」で新しいリフレッシュトークンを設定してください。';
      }

      throw new Error(msg);
    }

    const json = JSON.parse(text);

    if (!json.access_token) {
      throw new Error('アクセストークン取得失敗: access_token がレスポンスにありません。');
    }

    props.setProperty(KEY_ACCESS_TOKEN, json.access_token);

    if (json.expires_in) {
      props.setProperty(KEY_EXPIRES_AT, String(Date.now() + Number(json.expires_in) * 1000));
    }

    // Yahoo側が refresh token をローテーションして返した場合は保存し直す
    if (json.refresh_token) {
      props.setProperty(KEY_REFRESH_TOKEN, json.refresh_token);
      props.setProperty(KEY_REFRESH_SET_AT, String(Date.now()));
      console.log('Refresh Token was rotated and updated.');
    }

    return json.access_token;

  } finally {
    lock.releaseLock();
  }
}


// ==================================================
// 3. 初期設定
// ==================================================
/**
 * 初回だけ実行
 * CLIENT_ID / CLIENT_SECRET を保存する
 */
function Yahoo認証_初期設定_クライアント情報保存() {
  const props = PropertiesService.getScriptProperties();

  if (!DEFAULT_Y_CLIENT_ID || DEFAULT_Y_CLIENT_ID === 'ここにCLIENT_ID') {
    throw new Error('DEFAULT_Y_CLIENT_ID を自分の CLIENT_ID に書き換えてください。');
  }
  if (!DEFAULT_Y_CLIENT_SECRET || DEFAULT_Y_CLIENT_SECRET === 'ここにCLIENT_SECRET') {
    throw new Error('DEFAULT_Y_CLIENT_SECRET を自分の CLIENT_SECRET に書き換えてください。');
  }

  props.setProperties({
    [KEY_CLIENT_ID]: DEFAULT_Y_CLIENT_ID,
    [KEY_CLIENT_SECRET]: DEFAULT_Y_CLIENT_SECRET,
  });

  SpreadsheetApp.getUi().alert(
    'Yahoo認証 初期設定完了\n\n' +
    'CLIENT_ID / CLIENT_SECRET を保存しました。\n' +
    '次に「Yahoo認証_ボタン実行」で refresh token を設定してください。'
  );
}

/**
 * 事故防止: 古い関数は使わない
 */
function setupYahooAuth() {
  throw new Error('この関数は使用しません。「Yahoo認証_ボタン実行」を使ってください。');
}


// ==================================================
// 4. ボタン運用用
// ==================================================
/**
 * ボタンに割り当てる本命関数
 */
function Yahoo認証_ボタン実行() {
  const ui = SpreadsheetApp.getUi();

  const res = ui.prompt(
    'Yahoo リフレッシュトークン設定',
    '新しい refresh token を貼り付けてください。',
    ui.ButtonSet.OK_CANCEL
  );

  if (res.getSelectedButton() !== ui.Button.OK) {
    ui.alert('キャンセルしました。');
    return;
  }

  const refreshToken = String(res.getResponseText() || '').trim();
  if (!refreshToken) {
    ui.alert('refresh token が空です。処理を中止しました。');
    return;
  }

  Yahoo認証_リフレッシュトークン設定実行_(refreshToken);
}

/**
 * 入力された refresh token を保存し、その場で access token も取得
 */
function Yahoo認証_リフレッシュトークン設定実行_(refreshToken) {
  const props = PropertiesService.getScriptProperties();

  if (!refreshToken) {
    throw new Error('refresh token が空です。');
  }

  const clientId = props.getProperty(KEY_CLIENT_ID);
  const clientSecret = props.getProperty(KEY_CLIENT_SECRET);
  if (!clientId || !clientSecret) {
    throw new Error(
      'CLIENT_ID / CLIENT_SECRET が未設定です。\n' +
      '先に「Yahoo認証_初期設定_クライアント情報保存」を実行してください。'
    );
  }

  props.setProperty(KEY_REFRESH_TOKEN, String(refreshToken).trim());
  props.setProperty(KEY_REFRESH_SET_AT, String(Date.now()));

  // access token は取り直すので一旦消す
  props.deleteProperty(KEY_ACCESS_TOKEN);
  props.deleteProperty(KEY_EXPIRES_AT);

  const newAccessToken = refreshAccessToken_();

  const accessPreview = newAccessToken
    ? newAccessToken.substring(0, 20) + '...'
    : '(取得失敗)';

  SpreadsheetApp.getUi().alert(
    'Yahoo認証 完了\n\n' +
    '新しい refresh token を保存し、\n' +
    'access token も更新しました。\n\n' +
    'access token: ' + accessPreview
  );

  return newAccessToken;
}

/**
 * 状態確認用ボタン
 */
function Yahoo認証_状態確認_ボタン用() {
  const props = PropertiesService.getScriptProperties();

  const accessToken = props.getProperty(KEY_ACCESS_TOKEN);
  const refreshToken = props.getProperty(KEY_REFRESH_TOKEN);
  const expiresAt = Number(props.getProperty(KEY_EXPIRES_AT) || 0);
  const refreshSetAt = Number(props.getProperty(KEY_REFRESH_SET_AT) || 0);

  const now = Date.now();
  const isAccessExpired = !expiresAt || now > expiresAt;
  const accessExpiresDate = expiresAt ? new Date(expiresAt).toLocaleString('ja-JP') : '未設定';

  let refreshInfo = '未設定';
  if (refreshSetAt) {
    const refreshSetDate = new Date(refreshSetAt).toLocaleString('ja-JP');
    const elapsedDays = Math.floor((now - refreshSetAt) / (1000 * 60 * 60 * 24));
    const remainDays = REFRESH_TOKEN_LIFETIME_DAYS - elapsedDays;

    refreshInfo =
      refreshSetDate +
      `\n経過日数: ${elapsedDays}日` +
      `\n目安残日数: ${remainDays >= 0 ? remainDays : 0}日`;
  }

  SpreadsheetApp.getUi().alert(
    'Yahoo認証 状態確認\n\n' +
    'アクセストークン: ' + (accessToken ? '設定あり' : '未設定') + '\n' +
    'アクセストークン有効期限: ' + accessExpiresDate + (isAccessExpired ? '（期限切れ）' : '（有効）') + '\n\n' +
    'リフレッシュトークン: ' + (refreshToken ? '設定あり' : '未設定') + '\n' +
    'リフレッシュトークン設定日時:\n' + refreshInfo
  );
}

/**
 * 強制的にアクセストークンを更新したい時のボタン用
 */
function Yahoo認証_アクセストークン更新_ボタン用() {
  const token = refreshAccessToken_();
  SpreadsheetApp.getUi().alert(
    'Yahoo認証\n\nアクセストークンを更新しました。\n\n' +
    '先頭20文字: ' + token.substring(0, 20) + '...'
  );
  return token;
}

/**
 * 認証情報を全部消す時用
 */
function Yahoo認証_全クリア_ボタン用() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert(
    'Yahoo認証情報を削除',
    '保存済みの Yahoo 認証情報を削除します。よろしいですか？',
    ui.ButtonSet.OK_CANCEL
  );

  if (res !== ui.Button.OK) {
    ui.alert('キャンセルしました。');
    return;
  }

  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(KEY_REFRESH_TOKEN);
  props.deleteProperty(KEY_ACCESS_TOKEN);
  props.deleteProperty(KEY_EXPIRES_AT);
  props.deleteProperty(KEY_REFRESH_SET_AT);

  ui.alert('Yahoo認証情報を削除しました。');
}


// ==================================================
// 5. ログ確認用
// ==================================================
function 現在のトークン状態を確認() {
  const props = PropertiesService.getScriptProperties();

  const accessToken = props.getProperty(KEY_ACCESS_TOKEN);
  const refreshToken = props.getProperty(KEY_REFRESH_TOKEN);
  const expiresAt = Number(props.getProperty(KEY_EXPIRES_AT) || 0);
  const refreshSetAt = Number(props.getProperty(KEY_REFRESH_SET_AT) || 0);

  const now = Date.now();
  const isExpired = !expiresAt || now > expiresAt;
  const expiresDate = expiresAt ? new Date(expiresAt).toLocaleString('ja-JP') : '未設定';

  console.log('=== Yahoo トークン状態 ===');
  console.log('アクセストークン: ' + (accessToken ? '設定あり' : '未設定'));
  console.log('有効期限: ' + expiresDate + (isExpired ? ' ← 期限切れ' : ' ← 有効'));
  console.log('リフレッシュトークン: ' + (refreshToken ? refreshToken.substring(0, 30) + '...' : '未設定'));

  if (refreshSetAt) {
    const elapsedDays = Math.floor((now - refreshSetAt) / (1000 * 60 * 60 * 24));
    console.log('リフレッシュトークン設定日時: ' + new Date(refreshSetAt).toLocaleString('ja-JP'));
    console.log('リフレッシュトークン経過日数: ' + elapsedDays + '日');
  } else {
    console.log('リフレッシュトークン設定日時: 未設定');
  }

  return accessToken;
}

function 新しいアクセストークンを取得() {
  const token = refreshAccessToken_();
  console.log('=== 新しいアクセストークン ===');
  console.log(token);
  return token;
}