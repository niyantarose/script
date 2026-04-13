// Yahoo Shopping API - リフレッシュトークン取得スクリプト
// Google Apps Scriptにこのコードを貼り付けてデプロイしてください

// 設定
const CONFIG = {
  clientId: 'dj00aiZpPUQ4TEx2bDllVXl2eCZzPWNvbnN1bWVyc2VjcmV0Jng9YTM-',
  clientSecret: 'uLjglH91MOAljwDwR7tGVubXa1UJ54fsEYTpdbii',
  // ⚠️ 重要: デプロイしたら、このURLを実際のデプロイURLに書き換えてください
  redirectUri: 'https://script.google.com/macros/s/AKfycbzlug48mUX0YcBjWtaX6RHMNRh0IH3pdPU8FgO4G7_XK2wE0x_DBfPha_YrhObYVq743A/exec'
};

// Webアプリとしてアクセスされた時の処理
function doGet(e) {
  const code = e.parameter.code;
  const action = e.parameter.action;
  
  // リフレッシュトークン更新モード
  if (action === 'refresh') {
    const refreshToken = e.parameter.refresh_token;
    
    if (refreshToken) {
      try {
        const tokens = refreshAccessToken(refreshToken);
        return HtmlService.createHtmlOutput(createSuccessPage(tokens))
          .setTitle('Yahoo Shopping API - トークン更新成功');
      } catch (error) {
        return HtmlService.createHtmlOutput(createErrorPage(error.toString()))
          .setTitle('Yahoo Shopping API - エラー');
      }
    } else {
      return HtmlService.createHtmlOutput(createRefreshTokenInputPage())
        .setTitle('Yahoo Shopping API - トークン更新');
    }
  }
  
  // 認証コードからトークン取得モード
  if (code) {
    try {
      const tokens = getTokensFromAuthCode(code);
      return HtmlService.createHtmlOutput(createSuccessPage(tokens))
        .setTitle('Yahoo Shopping API - トークン取得成功');
    } catch (error) {
      return HtmlService.createHtmlOutput(createErrorPage(error.toString()))
        .setTitle('Yahoo Shopping API - エラー');
    }
  } else {
    // メインメニュー
    return HtmlService.createHtmlOutput(createMainMenu())
      .setTitle('Yahoo Shopping API - トークン管理');
  }
}

// 認証コードからトークンを取得
function getTokensFromAuthCode(authCode) {
  const tokenUrl = 'https://auth.login.yahoo.co.jp/yconnect/v2/token';
  
  const credentials = Utilities.base64Encode(CONFIG.clientId + ':' + CONFIG.clientSecret);
  
  const payload = {
    'grant_type': 'authorization_code',
    'code': authCode,
    'redirect_uri': CONFIG.redirectUri
  };
  
  const options = {
    'method': 'post',
    'headers': {
      'Authorization': 'Basic ' + credentials,
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    'payload': payload,
    'muteHttpExceptions': true
  };
  
  const response = UrlFetchApp.fetch(tokenUrl, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();
  
  if (responseCode !== 200) {
    throw new Error('トークン取得失敗 (' + responseCode + '): ' + responseText);
  }
  
  return JSON.parse(responseText);
}

// リフレッシュトークンから新しいアクセストークンを取得
function refreshAccessToken(refreshToken) {
  const tokenUrl = 'https://auth.login.yahoo.co.jp/yconnect/v2/token';
  
  const credentials = Utilities.base64Encode(CONFIG.clientId + ':' + CONFIG.clientSecret);
  
  const payload = {
    'grant_type': 'refresh_token',
    'refresh_token': refreshToken
  };
  
  const options = {
    'method': 'post',
    'headers': {
      'Authorization': 'Basic ' + credentials,
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    'payload': payload,
    'muteHttpExceptions': true
  };
  
  const response = UrlFetchApp.fetch(tokenUrl, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();
  
  if (responseCode !== 200) {
    throw new Error('トークン更新失敗 (' + responseCode + '): ' + responseText);
  }
  
  return JSON.parse(responseText);
}

// メインメニューのHTMLを生成
function createMainMenu() {
  const scriptUrl = CONFIG.redirectUri;
  const authUrl = 'https://auth.login.yahoo.co.jp/yconnect/v2/authorization?' +
    'response_type=code' +
    '&client_id=' + encodeURIComponent(CONFIG.clientId) +
    '&redirect_uri=' + encodeURIComponent(scriptUrl) +
    '&scope=openid';
  
  const refreshUrl = scriptUrl + '?action=refresh';
  
  return `
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Yahoo Shopping API - トークン管理</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
            padding: 20px;
        }
        .container {
            background: white;
            border-radius: 12px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            max-width: 600px;
            width: 100%;
            padding: 40px;
        }
        h1 {
            color: #333;
            margin-bottom: 10px;
            font-size: 28px;
            text-align: center;
        }
        .subtitle {
            color: #666;
            margin-bottom: 30px;
            text-align: center;
            font-size: 14px;
        }
        .option {
            background: #f8f9fa;
            border-left: 4px solid #667eea;
            padding: 25px;
            margin-bottom: 20px;
            border-radius: 8px;
            transition: transform 0.2s;
        }
        .option:hover {
            transform: translateX(5px);
        }
        .option-title {
            font-weight: bold;
            color: #667eea;
            margin-bottom: 10px;
            font-size: 18px;
        }
        .option-desc {
            color: #666;
            margin-bottom: 15px;
            font-size: 14px;
            line-height: 1.5;
        }
        .btn {
            display: inline-block;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 12px 30px;
            border-radius: 6px;
            text-decoration: none;
            font-size: 16px;
            font-weight: 600;
            transition: transform 0.2s, box-shadow 0.2s;
            text-align: center;
            width: 100%;
            display: block;
        }
        .btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 5px 20px rgba(102, 126, 234, 0.4);
        }
        .btn-secondary {
            background: linear-gradient(135deg, #4caf50 0%, #45a049 100%);
        }
        .info {
            background: #e3f2fd;
            border-left: 4px solid #2196f3;
            padding: 15px;
            margin-top: 30px;
            border-radius: 4px;
            font-size: 13px;
            color: #1565c0;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>🔐 Yahoo Shopping API</h1>
        <p class="subtitle">トークン管理ツール</p>
        
        <div class="option">
            <div class="option-title">🆕 新規取得</div>
            <div class="option-desc">
                初めてリフレッシュトークンを取得する場合、または完全に新しいトークンが必要な場合
            </div>
            <a href="${authUrl}" class="btn">Yahoo認証で新規取得</a>
        </div>
        
        <div class="option">
            <div class="option-title">🔄 トークン更新</div>
            <div class="option-desc">
                既存のリフレッシュトークンを使って新しいアクセストークンを取得する場合
            </div>
            <a href="${refreshUrl}" class="btn btn-secondary">リフレッシュトークンで更新</a>
        </div>
        
        <div class="info">
            <strong>💡 使い分け:</strong><br>
            • <strong>新規取得</strong>: リフレッシュトークンが失効した、または初めて取得する<br>
            • <strong>トークン更新</strong>: リフレッシュトークンを使ってアクセストークンを更新（簡単・高速）
        </div>
    </div>
</body>
</html>
  `;
}

// 認証ページのHTMLを生成（廃止予定だが互換性のため残す）
function createAuthPage() {
  return createMainMenu();
}

// 成功ページのHTMLを生成
function createSuccessPage(tokens) {
  const scriptUrl = CONFIG.redirectUri;
  
  return `
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>トークン取得成功</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
            padding: 20px;
        }
        .container {
            background: white;
            border-radius: 12px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            max-width: 800px;
            width: 100%;
            padding: 40px;
        }
        h1 {
            color: #2e7d32;
            margin-bottom: 20px;
            font-size: 28px;
        }
        .token-section {
            background: #f8f9fa;
            border-left: 4px solid #4caf50;
            padding: 20px;
            margin-bottom: 20px;
            border-radius: 4px;
        }
        .token-title {
            font-weight: bold;
            color: #2e7d32;
            margin-bottom: 10px;
            font-size: 16px;
        }
        .token-display {
            background: white;
            padding: 15px;
            border-radius: 4px;
            word-break: break-all;
            font-family: 'Courier New', monospace;
            font-size: 13px;
            color: #333;
            border: 1px solid #a5d6a7;
            margin-top: 10px;
        }
        .copy-btn {
            background: #4caf50;
            color: white;
            border: none;
            padding: 10px 25px;
            border-radius: 6px;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            margin-top: 10px;
            transition: background 0.2s;
        }
        .copy-btn:hover {
            background: #45a049;
        }
        .info {
            background: #fff3cd;
            border: 1px solid #ffc107;
            padding: 15px;
            border-radius: 4px;
            margin-top: 20px;
            font-size: 13px;
            color: #856404;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>✅ トークン取得成功!</h1>
        
        <div class="token-section">
            <div class="token-title">Access Token:</div>
            <div class="token-display" id="accessToken">${tokens.access_token}</div>
            <button class="copy-btn" onclick="copyToken('accessToken')">📋 コピー</button>
        </div>
        
        <div class="token-section">
            <div class="token-title">Refresh Token (これをスクリプトに設定):</div>
            <div class="token-display" id="refreshToken">${tokens.refresh_token}</div>
            <button class="copy-btn" onclick="copyToken('refreshToken')">📋 コピー</button>
        </div>
        
        <div class="info">
            <strong>⏰ 有効期限:</strong> ${tokens.expires_in}秒 (約${Math.floor(tokens.expires_in / 3600)}時間)<br>
            <strong>📅 取得日時:</strong> ${new Date().toLocaleString('ja-JP')}
        </div>
        
        <div style="text-align: center; margin-top: 30px;">
            <a href="${scriptUrl}" style="display: inline-block; background: #667eea; color: white; padding: 12px 30px; border-radius: 6px; text-decoration: none; font-weight: 600;">🏠 トップに戻る</a>
        </div>
    </div>
    
    <script>
        function copyToken(elementId) {
            const text = document.getElementById(elementId).textContent;
            navigator.clipboard.writeText(text).then(() => {
                const btn = event.target;
                const originalText = btn.textContent;
                btn.textContent = '✓ コピーしました!';
                setTimeout(() => {
                    btn.textContent = originalText;
                }, 2000);
            }).catch(err => {
                alert('コピーに失敗しました: ' + err);
            });
        }
    </script>
</body>
</html>
  `;
}

// リフレッシュトークン入力ページのHTMLを生成
function createRefreshTokenInputPage() {
  const scriptUrl = CONFIG.redirectUri;
  
  return `
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>リフレッシュトークンで更新</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
            padding: 20px;
        }
        .container {
            background: white;
            border-radius: 12px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            max-width: 600px;
            width: 100%;
            padding: 40px;
        }
        h1 {
            color: #333;
            margin-bottom: 20px;
            font-size: 28px;
        }
        .form-group {
            margin-bottom: 20px;
        }
        label {
            display: block;
            margin-bottom: 8px;
            color: #555;
            font-weight: 600;
            font-size: 14px;
        }
        textarea {
            width: 100%;
            padding: 15px;
            border: 2px solid #e0e0e0;
            border-radius: 6px;
            font-size: 14px;
            font-family: 'Courier New', monospace;
            resize: vertical;
            min-height: 100px;
            transition: border-color 0.3s;
        }
        textarea:focus {
            outline: none;
            border-color: #667eea;
        }
        .btn {
            background: linear-gradient(135deg, #4caf50 0%, #45a049 100%);
            color: white;
            border: none;
            padding: 15px 30px;
            border-radius: 6px;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
            width: 100%;
            transition: transform 0.2s, box-shadow 0.2s;
        }
        .btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 5px 20px rgba(76, 175, 80, 0.4);
        }
        .btn:disabled {
            background: #ccc;
            cursor: not-allowed;
            transform: none;
        }
        .info {
            background: #e3f2fd;
            border-left: 4px solid #2196f3;
            padding: 15px;
            margin-bottom: 20px;
            border-radius: 4px;
            font-size: 13px;
            color: #1565c0;
        }
        .back-link {
            text-align: center;
            margin-top: 20px;
        }
        .back-link a {
            color: #667eea;
            text-decoration: none;
            font-size: 14px;
        }
        .back-link a:hover {
            text-decoration: underline;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>🔄 トークン更新</h1>
        
        <div class="info">
            <strong>📌 使い方:</strong><br>
            既存のリフレッシュトークンを入力して、新しいアクセストークンを取得します。<br>
            これはYahoo認証を経由せず、すぐに更新できます。
        </div>
        
        <form onsubmit="updateToken(event)">
            <div class="form-group">
                <label for="refreshToken">リフレッシュトークン:</label>
                <textarea id="refreshToken" placeholder="リフレッシュトークンをここに貼り付けてください" required></textarea>
            </div>
            
            <button type="submit" class="btn" id="submitBtn">🔑 トークン更新</button>
        </form>
        
        <div class="back-link">
            <a href="${scriptUrl}">← トップに戻る</a>
        </div>
    </div>
    
    <script>
        function updateToken(event) {
            event.preventDefault();
            
            const refreshToken = document.getElementById('refreshToken').value.trim();
            const btn = document.getElementById('submitBtn');
            
            if (!refreshToken) {
                alert('リフレッシュトークンを入力してください');
                return;
            }
            
            btn.disabled = true;
            btn.textContent = '更新中...';
            
            // リダイレクト
            const url = '${scriptUrl}?action=refresh&refresh_token=' + encodeURIComponent(refreshToken);
            window.location.href = url;
        }
    </script>
</body>
</html>
  `;
}

// エラーページのHTMLを生成
function createErrorPage(error) {
  const scriptUrl = CONFIG.redirectUri;
  
  return `
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>エラー</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
            padding: 20px;
        }
        .container {
            background: white;
            border-radius: 12px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            max-width: 600px;
            width: 100%;
            padding: 40px;
        }
        h1 {
            color: #c62828;
            margin-bottom: 20px;
            font-size: 28px;
        }
        .error-box {
            background: #ffebee;
            border: 2px solid #f44336;
            padding: 20px;
            border-radius: 6px;
            color: #c62828;
            white-space: pre-line;
            font-size: 14px;
        }
        .btn {
            display: inline-block;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 12px 30px;
            border-radius: 6px;
            text-decoration: none;
            font-size: 16px;
            font-weight: 600;
            margin-top: 20px;
            transition: transform 0.2s;
        }
        .btn:hover {
            transform: translateY(-2px);
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>❌ エラーが発生しました</h1>
        <div class="error-box">${error}</div>
        <div style="display: flex; gap: 10px; margin-top: 20px;">
            <a href="${scriptUrl}" class="btn" style="flex: 1; text-align: center;">🏠 トップに戻る</a>
            <a href="${scriptUrl}?action=refresh" class="btn" style="flex: 1; background: #4caf50; text-align: center;">🔄 やり直す</a>
        </div>
    </div>
</body>
</html>
  `;
}
