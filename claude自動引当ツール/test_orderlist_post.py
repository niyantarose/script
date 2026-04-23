#!/usr/bin/env python3
"""Test orderList with POST + XML on the VPS."""
import sys
sys.path.insert(0, '/home/ubuntu/zaiko-tool/app')

import os
# load .env
env_path = '/home/ubuntu/zaiko-tool/app/.env'
for line in open(env_path):
    line = line.strip()
    if line and not line.startswith('#') and '=' in line:
        k, v = line.split('=', 1)
        os.environ.setdefault(k.strip(), v.strip())

from services.yahoo_api import YahooAPI
api = YahooAPI()

print('Testing orderList POST+XML...')
try:
    result = api.search_orders(days=30, hits=5)
    print(f'SUCCESS! Keys: {list(result.keys())}')
    import json
    print(json.dumps(result, ensure_ascii=False, indent=2)[:1000])
except Exception as e:
    print(f'ERROR: {e}')
