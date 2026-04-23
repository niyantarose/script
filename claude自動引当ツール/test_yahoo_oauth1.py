#!/usr/bin/env python3
"""Test Yahoo Japan Shopping API with OAuth 1.0a HMAC-SHA1 signing."""
import os, time, uuid, hmac, hashlib, base64
from urllib.parse import quote, urlencode
import requests

CONSUMER_KEY    = os.environ.get('YAHOO_CLIENT_ID', 'dj00aiZpPUQ4TEx2bDllVXl2eCZzPWNvbnN1bWVyc2VjcmV0Jng9YTM-')
CONSUMER_SECRET = os.environ.get('YAHOO_CLIENT_SECRET', 'uLjglH91MOAljwDwR7tGVubXa1UJ54fsEYTpdbii')
SELLER_ID       = os.environ.get('YAHOO_SELLER_ID', 'niyantarose')
REFRESH_TOKEN   = os.environ.get('YAHOO_REFRESH_TOKEN', '')

def make_oauth1_header(method, url, extra_params):
    ts    = str(int(time.time()))
    nonce = uuid.uuid4().hex
    oauth_params = {
        'oauth_consumer_key':     CONSUMER_KEY,
        'oauth_nonce':            nonce,
        'oauth_signature_method': 'HMAC-SHA1',
        'oauth_timestamp':        ts,
        'oauth_version':          '1.0',
    }
    all_params = {**extra_params, **oauth_params}
    sorted_qs  = '&'.join(
        quote(k, safe='') + '=' + quote(str(v), safe='')
        for k, v in sorted(all_params.items())
    )
    base_str     = '&'.join([method.upper(), quote(url, safe=''), quote(sorted_qs, safe='')])
    signing_key  = quote(CONSUMER_SECRET, safe='') + '&'  # no access token secret
    sig          = base64.b64encode(
        hmac.new(signing_key.encode(), base_str.encode(), hashlib.sha1).digest()
    ).decode()
    oauth_params['oauth_signature'] = sig
    auth_header = 'OAuth ' + ', '.join(
        quote(k, safe='') + '="' + quote(str(v), safe='') + '"'
        for k, v in sorted(oauth_params.items())
    )
    return auth_header


def test_endpoint(name, method, url, params=None, body=None):
    print(f"\n{'='*60}")
    print(f"TEST: {name}")
    print(f"URL:  {url}")
    extra = dict(params or {})
    auth  = make_oauth1_header(method, url, extra)
    headers = {
        'Authorization': auth,
        'Content-Type':  'application/xml',
    }
    try:
        if method == 'GET':
            r = requests.get(url, params=params, headers=headers, timeout=15)
        else:
            r = requests.post(url, params=params, data=body, headers=headers, timeout=15)
        print(f"Status: {r.status_code}")
        print(f"Response (first 500 chars): {r.text[:500]}")
    except Exception as e:
        print(f"Error: {e}")


BASE = 'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1'

# Test 1: orderCount (known to return 200 with Bearer — let's try OAuth1)
test_endpoint(
    'orderCount (OAuth1)',
    'GET',
    f'{BASE}/orderCount',
    params={'sellerId': SELLER_ID},
)

# Test 2: orderList
test_endpoint(
    'orderList (OAuth1)',
    'GET',
    f'{BASE}/orderList',
    params={'sellerId': SELLER_ID, 'sts': 'ordered'},
)

# Test 3: searchOrder (POST)
search_body = f'''<?xml version="1.0" encoding="UTF-8"?>
<Req>
  <Search>
    <Condition>
      <SellerId>{SELLER_ID}</SellerId>
      <OrderTimeFrom>20260101000000</OrderTimeFrom>
    </Condition>
    <Setting>
      <Results>10</Results>
      <Start>1</Start>
    </Setting>
  </Search>
</Req>'''
test_endpoint(
    'searchOrder POST (OAuth1)',
    'POST',
    f'{BASE}/searchOrder',
    params={'sellerId': SELLER_ID},
    body=search_body,
)
