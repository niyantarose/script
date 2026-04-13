/******************************************************
 * Yahoo Shopping API - トークン取得/更新（修正版）
 * - Yahoo認証を別タブで確実に開く
 * - code -> token 交換（refresh_token含む）
 * - refresh_token -> access_token 更新
 ******************************************************/

const 設定 = {
  クライアントID: 'dj00aiZpPUQ4TEx2bDllVXl2eCZzPWNvbnN1bWVyc2VjcmV0Jng9YTM-',
  クライアントシークレット: 'uLjglH91MOAljwDwR7tGVubXa1UJ54fsEYTpdbii',
  トークンURL: 'https://auth.login.yahoo.co.jp/yconnect/v2/token',
  認可URL: 'https://auth.login.yahoo.co.jp/yconnect/v2/authorization',
  スコープ: 'openid',
};

/** WebアプリのURL(/exec)を返す */
function ウェブアプリURLを取得() {
  return ScriptApp.getService().getUrl();
}

/** トップ画面HTML */
function トップ画面を作る_() {
  const webUrl = ウェブアプリURLを取得();

  const authUrl =
    設定.認可URL
    + '?response_type=code'
    + '&client_id=' + encodeURIComponent(設定.クライアントID)
    + '&redirect_uri=' + encodeURIComponent(webUrl)
    + '&scope=' + encodeURIComponent(設定.スコープ);

  const refreshUrl = webUrl + '?action=refresh';

  return `
<!doctype html><html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Yahoo トークン管理</title>
<style>
  body { 
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
    padding: 20px; 
    max-width: 700px; 
    margin: 0 auto;
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    min-height: 100vh;
  }
  .container {
    background: white;
    padding: 30px;
    border-radius: 12px;
    box-shadow: 0 20px 60px rgba(0,0,0,0.3);
  }
  h2 { color: #333; margin-top: 0; }
  .card { 
    background: #f8f9fa; 
    padding: 20px; 
    margin: 16px 0; 
    border-left: 4px solid #667eea; 
    border-radius: 8px;
  }
  .card-success { border-left-color: #2e7d32; }
  .card h3 { margin-top: 0; color: #667eea; }
  .card-success h3 { color: #2e7d32; }
  .btn { 
    display: inline-block; 
    padding: 14px 28px; 
    margin: 8px 0; 
    text-decoration: none; 
    border-radius: 6px; 
    font-weight: 600; 
    cursor: pointer; 
    border: none; 
    font-size: 16px;
    width: 100%;
    transition: transform 0.2s, box-shadow 0.2s;
  }
  .btn:hover { 
    transform: translateY(-2px);
    box-shadow: 0 5px 20px rgba(0,0,0,0.2);
  }
  .btn-primary { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: #fff; }
  .btn-success { background: linear-gradient(135deg, #4caf50 0%, #45a049 100%); color: #fff; }
  .info {
    background: #e3f2fd;
    padding: 16px;
    border-left: 4px solid #2196f3;
    margin: 20px 0;
    border-radius: 4px;
    font-size: 14px;
    color: #1565c0;
  }
</style>
</head><body>
<div class="container">
  <h2>🔐 Yahoo Shopping API - トークン管理</h2>

  <div class="card">
    <h3>🆕 新規取得</h3>
    <p>初めてリフレッシュトークンを取得する場合、または完全に新しいトークンが必要な場合</p>
    <button class="btn btn-primary" onclick="openYahooAuth()">Yahoo認証で新規取得</button>
  </div>

  <div class="card card-success">
    <h3>🔄 トークン更新</h3>
    <p>既存のリフレッシュトークンを使って新しいアクセストークンを取得する場合</p>
    <button class="btn btn-success" onclick="location.href='${refreshUrl}'">リフレッシュトークンで更新</button>
  </div>

  <div class="info">
    <strong>💡 使い分け:</strong><br>
    • <strong>新規取得</strong>: リフレッシュトークンが失効した、または初めて取得する<br>
    • <strong>トークン更新</strong>: リフレッシュトークンを使ってアクセストークンを更新（簡単・高速）
  </div>
</div>

<script>
  function openYahooAuth() {
    const authUrl = '${authUrl}';
    // window.open で確実に別タブで開く
    const newWindow = window.open(authUrl, '_blank', 'noopener,noreferrer');
    
    if (!newWindow || newWindow.closed || typeof newWindow.closed == 'undefined') {
      // ポップアップがブロックされた場合
      alert('ポップアップがブロックされました。\\n\\nブラウザの設定でポップアップを許可するか、\\n下のリンクをCtrl+クリック（Mac: Cmd+クリック）で開いてください。');
      
      // フォールバック: リンクを表示
      const link = document.createElement('a');
      link.href = authUrl;
      link.target = '_blank';
      link.rel = 'noopener noreferrer';
      link.textContent = 'Yahoo認証ページを開く（クリック）';
      link.style.cssText = 'display:block;margin:20px 0;padding:15px;background:#ff9800;color:white;text-align:center;text-decoration:none;border-radius:6px;font-weight:bold;';
      document.body.appendChild(link);
    }
  }
</script>
</body></html>`;
}

/** refresh_token入力画面 */
function リフレッシュ入力画面を作る_() {
  const webUrl = ウェブアプリURLを取得();
  return `
<!doctype html><html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>refresh_tokenで更新</title>
<style>
  body { 
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
    padding: 20px;
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    min-height: 100vh;
  }
  .container {
    max-width: 600px;
    margin: 0 auto;
    background: white;
    padding: 30px;
    border-radius: 12px;
    box-shadow: 0 20px 60px rgba(0,0,0,0.3);
  }
  h2 { color: #333; margin-top: 0; }
  textarea {
    width: 100%;
    height: 120px;
    padding: 12px;
    border: 2px solid #e0e0e0;
    border-radius: 6px;
    font-family: 'Courier New', monospace;
    font-size: 14px;
    resize: vertical;
  }
  textarea:focus {
    outline: none;
    border-color: #667eea;
  }
  button {
    background: linear-gradient(135deg, #4caf50 0%, #45a049 100%);
    color: white;
    padding: 14px 28px;
    border: none;
    border-radius: 6px;
    font-size: 16px;
    font-weight: 600;
    cursor: pointer;
    width: 100%;
    margin-top: 10px;
    transition: transform 0.2s;
  }
  button:hover {
    transform: translateY(-2px);
  }
  a {
    color: #667eea;
    text-decoration: none;
    font-weight: 600;
  }
  a:hover {
    text-decoration: underline;
  }
</style>
</head>
<body>
<div class="container">
  <h2>🔄 refresh_token で access_token 更新</h2>
  <form onsubmit="go(event)">
    <label for="rt" style="display:block;margin-bottom:8px;color:#555;font-weight:600;">リフレッシュトークン:</label>
    <textarea id="rt" placeholder="refresh_token を貼り付けてください" required></textarea>
    <button type="submit">🔑 トークン更新</button>
  </form>
  <p style="margin-top:20px;"><a href="${webUrl}">← トップに戻る</a></p>
</div>
<script>
  function go(e){
    e.preventDefault();
    const rt = document.getElementById('rt').value.trim();
    if (!rt) {
      alert('refresh_token を入力してください');
      return;
    }
    location.href = '${webUrl}?action=refresh&refresh_token=' + encodeURIComponent(rt);
  }
</script>
</body></html>`;
}

/** doGet：Webアプリ入口 */
function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  const action = p.action || '';
  const code = p.code || '';

  // 更新モード
  if (action === 'refresh') {
    const refreshToken = p.refresh_token || '';
    if (!refreshToken) {
      return HtmlService.createHtmlOutput(リフレッシュ入力画面を作る_());
    }
    try {
      const tokens = リフレッシュで更新する_(refreshToken);
      return HtmlService.createHtmlOutput(成功画面を作る_(tokens, '更新成功'));
    } catch (err) {
      return HtmlService.createHtmlOutput(エラー画面を作る_(String(err)));
    }
  }

  // code 受け取りモード
  if (code) {
    try {
      const tokens = 認可コードから取得する_(code);
      return HtmlService.createHtmlOutput(成功画面を作る_(tokens, '取得成功'));
    } catch (err) {
      return HtmlService.createHtmlOutput(エラー画面を作る_(String(err)));
    }
  }

  // トップ
  return HtmlService.createHtmlOutput(トップ画面を作る_());
}

/** code -> token */
function 認可コードから取得する_(code) {
  const redirectUri = ウェブアプリURLを取得();
  const basic = Utilities.base64Encode(`${設定.クライアントID}:${設定.クライアントシークレット}`);

  const res = UrlFetchApp.fetch(設定.トークンURL, {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    headers: { Authorization: `Basic ${basic}` },
    payload: {
      grant_type: 'authorization_code',
      code,
      redirect_uri: redirectUri,
    },
    muteHttpExceptions: true,
  });

  const http = res.getResponseCode();
  const text = res.getContentText();
  if (http !== 200) throw new Error(`トークン取得失敗 HTTP=${http}\n${text}`);

  const json = JSON.parse(text);

  // 保存（必要なら）
  const props = PropertiesService.getScriptProperties();
  if (json.refresh_token) props.setProperty('Y_REFRESH_TOKEN', json.refresh_token);
  if (json.access_token) props.setProperty('Y_ACCESS_TOKEN', json.access_token);

  return json;
}

/** refresh_token -> access_token */
function リフレッシュで更新する_(refreshToken) {
  const basic = Utilities.base64Encode(`${設定.クライアントID}:${設定.クライアントシークレット}`);

  const res = UrlFetchApp.fetch(設定.トークンURL, {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    headers: { Authorization: `Basic ${basic}` },
    payload: {
      grant_type: 'refresh_token',
      refresh_token: refreshToken,
    },
    muteHttpExceptions: true,
  });

  const http = res.getResponseCode();
  const text = res.getContentText();
  if (http !== 200) throw new Error(`更新失敗 HTTP=${http}\n${text}`);

  const json = JSON.parse(text);

  // 保存（必要なら）
  const props = PropertiesService.getScriptProperties();
  if (json.access_token) props.setProperty('Y_ACCESS_TOKEN', json.access_token);

  return json;
}

/** 成功画面 */
function 成功画面を作る_(tokens, title) {
  const webUrl = ウェブアプリURLを取得();
  const access = tokens.access_token ? String(tokens.access_token) : '(なし)';
  const refresh = tokens.refresh_token ? String(tokens.refresh_token) : '(返ってこない場合あり)';

  return `
<!doctype html><html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${title}</title>
<style>
  body {
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    padding: 20px;
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    min-height: 100vh;
  }
  .container {
    max-width: 800px;
    margin: 0 auto;
    background: white;
    padding: 30px;
    border-radius: 12px;
    box-shadow: 0 20px 60px rgba(0,0,0,0.3);
  }
  h2 { color: #2e7d32; margin-top: 0; }
  h3 { color: #333; margin-top: 20px; }
  .token-box {
    background: #f5f5f5;
    padding: 15px;
    border-radius: 6px;
    word-break: break-all;
    font-family: 'Courier New', monospace;
    font-size: 13px;
    border: 2px solid #e0e0e0;
    margin-top: 10px;
  }
  .btn-copy {
    background: #4caf50;
    color: white;
    padding: 8px 20px;
    border: none;
    border-radius: 6px;
    cursor: pointer;
    margin-top: 10px;
    font-weight: 600;
  }
  .btn-copy:hover {
    background: #45a049;
  }
  .warning {
    background: #fff3cd;
    border-left: 4px solid #ffc107;
    padding: 15px;
    margin-top: 20px;
    border-radius: 4px;
    color: #856404;
  }
  a {
    color: #667eea;
    text-decoration: none;
    font-weight: 600;
  }
  a:hover {
    text-decoration: underline;
  }
</style>
</head>
<body>
<div class="container">
  <h2>✅ ${title}</h2>
  
  <h3>Access Token:</h3>
  <div class="token-box" id="access">${escapeHtml_(access)}</div>
  <button class="btn-copy" onclick="copy('access')">📋 コピー</button>

  <h3>Refresh Token:</h3>
  <div class="token-box" id="refresh">${escapeHtml_(refresh)}</div>
  <button class="btn-copy" onclick="copy('refresh')">📋 コピー</button>

  <div class="warning">
    <strong>⚠️ 重要:</strong><br>
    • Refresh Tokenは大切に保管してください（漏洩すると第三者があなたのアカウントにアクセスできます）<br>
    • このトークンをスクリプトの設定に保存してください<br>
    • 有効期限: Access Token は約1時間、Refresh Token は長期間有効
  </div>

  <p style="margin-top:30px;text-align:center;">
    <a href="${webUrl}">🏠 トップに戻る</a>
  </p>
</div>

<script>
  function copy(id) {
    const el = document.getElementById(id);
    const text = el.textContent;
    navigator.clipboard.writeText(text).then(() => {
      const btn = event.target;
      const orig = btn.textContent;
      btn.textContent = '✓ コピーしました!';
      setTimeout(() => btn.textContent = orig, 2000);
    }).catch(err => {
      alert('コピー失敗: ' + err);
    });
  }
</script>
</body></html>`;
}

/** エラー画面 */
function エラー画面を作る_(msg) {
  const webUrl = ウェブアプリURLを取得();
  return `
<!doctype html><html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>エラー</title>
<style>
  body {
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    padding: 20px;
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    min-height: 100vh;
  }
  .container {
    max-width: 600px;
    margin: 0 auto;
    background: white;
    padding: 30px;
    border-radius: 12px;
    box-shadow: 0 20px 60px rgba(0,0,0,0.3);
  }
  h2 { color: #c62828; margin-top: 0; }
  .error-box {
    background: #ffebee;
    padding: 15px;
    border-radius: 6px;
    border: 2px solid #f44336;
    word-break: break-all;
    font-family: 'Courier New', monospace;
    font-size: 13px;
    color: #c62828;
    white-space: pre-wrap;
  }
  a {
    color: #667eea;
    text-decoration: none;
    font-weight: 600;
  }
  a:hover {
    text-decoration: underline;
  }
</style>
</head>
<body>
<div class="container">
  <h2>❌ エラーが発生しました</h2>
  <div class="error-box">${escapeHtml_(msg)}</div>
  <p style="margin-top:20px;"><a href="${webUrl}">🏠 トップに戻る</a> | <a href="${webUrl}?action=refresh">🔄 やり直す</a></p>
</div>
</body></html>`;
}

function escapeHtml_(s) {
  return String(s).replace(/[&<>"']/g, c => ({
    "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"
  }[c]));
}
