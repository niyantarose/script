#!/usr/bin/env python3
"""Test Yahoo Japan Shopping API with Bearer token — try multiple endpoints."""
import os, sys, json, base64
import urllib.request, urllib.parse, urllib.error

CLIENT_ID     = 'dj00aiZpPUQ4TEx2bDllVXl2eCZzPWNvbnN1bWVyc2VjcmV0Jng9YTM-'
CLIENT_SECRET = 'uLjglH91MOAljwDwR7tGVubXa1UJ54fsEYTpdbii'
SELLER_ID     = 'niyantarose'
REFRESH_TOKEN = os.environ.get('YAHOO_REFRESH_TOKEN', '')

def get_access_token():
    credentials = base64.b64encode(f'{CLIENT_ID}:{CLIENT_SECRET}'.encode()).decode()
    data = urllib.parse.urlencode({
        'grant_type':    'refresh_token',
        'refresh_token': REFRESH_TOKEN,
    }).encode()
    req = urllib.request.Request(
        'https://auth.login.yahoo.co.jp/yconnect/v2/token',
        data=data,
        headers={
            'Authorization': f'Basic {credentials}',
            'Content-Type':  'application/x-www-form-urlencoded',
        }
    )
    with urllib.request.urlopen(req, timeout=20) as r:
        d = json.loads(r.read())
    print(f'[token] access_token prefix: {d.get("access_token","")[:20]}...')
    if d.get('refresh_token'):
        print(f'[token] new refresh_token returned, saving...')
        env_path = '/home/ubuntu/zaiko-tool/app/.env'
        try:
            txt = open(env_path).read()
            import re
            txt = re.sub(r'^YAHOO_REFRESH_TOKEN=.*$', f'YAHOO_REFRESH_TOKEN={d["refresh_token"]}', txt, flags=re.MULTILINE)
            open(env_path, 'w').write(txt)
        except Exception as e:
            print(f'  save failed: {e}')
    return d['access_token']

def try_get(name, url, params, token):
    full = url + '?' + urllib.parse.urlencode(params)
    req  = urllib.request.Request(full, headers={'Authorization': f'Bearer {token}'})
    print(f'\n--- {name} ---')
    print(f'    URL: {full[:100]}')
    try:
        with urllib.request.urlopen(req, timeout=20) as r:
            body = r.read().decode('utf-8', errors='replace')
        print(f'    Status: 200 OK')
        print(f'    Body (500): {body[:500]}')
    except urllib.error.HTTPError as e:
        body = e.read().decode('utf-8', errors='replace')
        print(f'    Status: {e.code}')
        print(f'    Body (300): {body[:300]}')

BASE = 'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1'

token = get_access_token()

# orderCount — known working
try_get('orderCount', f'{BASE}/orderCount',
        {'sellerId': SELLER_ID, 'output': 'json'}, token)

# searchOrder GET (might be wrong method)
try_get('searchOrder GET', f'{BASE}/searchOrder',
        {'sellerId': SELLER_ID, 'output': 'json'}, token)

# orderList GET
try_get('orderList GET', f'{BASE}/orderList',
        {'sellerId': SELLER_ID, 'output': 'json'}, token)

# orderInfo GET (needs orderId but tests path recognition)
try_get('orderInfo GET (no orderId)', f'{BASE}/orderInfo',
        {'sellerId': SELLER_ID, 'output': 'json'}, token)

# Try v2 base
try_get('v2/searchOrder', 'https://circus.shopping.yahooapis.jp/ShoppingWebService/V2/searchOrder',
        {'sellerId': SELLER_ID, 'output': 'json'}, token)
