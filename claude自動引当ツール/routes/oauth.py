"""Yahoo OAuth 2.0 認証フロー
アクセストークン・リフレッシュトークンをブラウザ経由で取得して .env に保存する
"""
import os, re, base64, json
import urllib.parse, urllib.request
from flask import Blueprint, request, redirect, render_template_string

bp = Blueprint('oauth', __name__, url_prefix='/oauth')

YAHOO_AUTH_URL  = 'https://auth.login.yahoo.co.jp/yconnect/v2/authorization'
YAHOO_TOKEN_URL = 'https://auth.login.yahoo.co.jp/yconnect/v2/token'
CALLBACK_URL    = 'http://localhost:5001/oauth/callback'

# Yahoo ショッピング注文APIに必要なスコープ
SCOPES = 'openid profile email'


# ─── 認証開始 ──────────────────────────────────────────────────
@bp.route('/start')
def oauth_start():
    client_id = os.getenv('YAHOO_CLIENT_ID', '')
    if not client_id:
        return '<h2>❌ YAHOO_CLIENT_ID が .env に設定されていません</h2>', 500

    params = {
        'response_type': 'code',
        'client_id':     client_id,
        'redirect_uri':  CALLBACK_URL,
        'scope':         SCOPES,
        'bail':          '1',
    }
    auth_url = YAHOO_AUTH_URL + '?' + urllib.parse.urlencode(params)
    return redirect(auth_url)


# ─── コールバック受け取り ──────────────────────────────────────
@bp.route('/callback')
def oauth_callback():
    error = request.args.get('error')
    if error:
        desc = request.args.get('error_description', '')
        return _html('❌ 認証エラー', f'<p><b>{error}</b>: {desc}</p>', ok=False)

    code = request.args.get('code')
    if not code:
        return _html('❌ コードなし', f'<p>パラメータ: {dict(request.args)}</p>', ok=False)

    # アクセストークン取得
    client_id     = os.getenv('YAHOO_CLIENT_ID', '')
    client_secret = os.getenv('YAHOO_CLIENT_SECRET', '')
    credentials   = base64.b64encode(f'{client_id}:{client_secret}'.encode()).decode()

    post_data = urllib.parse.urlencode({
        'grant_type':   'authorization_code',
        'code':         code,
        'redirect_uri': CALLBACK_URL,
    }).encode('utf-8')

    req = urllib.request.Request(
        YAHOO_TOKEN_URL,
        data=post_data,
        headers={
            'Authorization': f'Basic {credentials}',
            'Content-Type':  'application/x-www-form-urlencoded',
        }
    )
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            token_data = json.loads(resp.read().decode('utf-8'))
    except urllib.error.HTTPError as e:
        body = e.read().decode('utf-8', errors='replace')
        return _html('❌ トークン取得失敗', f'<p>HTTP {e.code}</p><pre>{body}</pre>', ok=False)
    except Exception as e:
        return _html('❌ 通信エラー', f'<pre>{e}</pre>', ok=False)

    refresh_token = token_data.get('refresh_token', '')
    access_token  = token_data.get('access_token', '')

    if not refresh_token:
        return _html('⚠️ リフレッシュトークンなし',
                     f'<p>レスポンス内容:</p><pre>{json.dumps(token_data, indent=2)}</pre>',
                     ok=False)

    # .env に保存
    _update_env('YAHOO_REFRESH_TOKEN', refresh_token)

    return _html('✅ 認証成功！', f'''
        <p>リフレッシュトークンを <code>.env</code> に保存しました。</p>
        <p>このウィンドウを閉じてください。</p>
        <hr>
        <p style="font-size:12px;color:#888">
          Access Token (参考): {access_token[:30]}...<br>
          Refresh Token (先頭20文字): {refresh_token[:20]}...
        </p>
    ''', ok=True)


# ─── .env 更新ヘルパー ─────────────────────────────────────────
def _update_env(key, value):
    env_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), '.env')
    try:
        with open(env_path, 'r', encoding='utf-8') as f:
            content = f.read()
        if re.search(rf'^{key}=', content, re.MULTILINE):
            content = re.sub(rf'^{key}=.*$', f'{key}={value}', content, flags=re.MULTILINE)
        else:
            content += f'\n{key}={value}\n'
        with open(env_path, 'w', encoding='utf-8') as f:
            f.write(content)
    except Exception as e:
        raise RuntimeError(f'.env 更新失敗: {e}')


# ─── HTMLテンプレート ──────────────────────────────────────────
def _html(title, body, ok=True):
    color = '#2e7d32' if ok else '#c62828'
    return render_template_string(f'''
<!DOCTYPE html><html lang="ja"><head>
<meta charset="utf-8">
<title>{title}</title>
<style>
  body {{ font-family: sans-serif; max-width:600px; margin:60px auto; padding:20px; }}
  h2 {{ color: {color}; }}
  pre {{ background:#f5f5f5; padding:12px; border-radius:4px; overflow:auto; }}
  code {{ background:#f5f5f5; padding:2px 6px; border-radius:3px; }}
</style>
</head><body>
<h2>{title}</h2>
{body}
<p><a href="/">← ダッシュボードへ戻る</a></p>
</body></html>
''')
