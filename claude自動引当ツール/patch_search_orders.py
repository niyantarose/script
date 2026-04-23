#!/usr/bin/env python3
"""Patch search_orders to use POST + XML as required by Yahoo orderList API."""
import re

path = '/home/ubuntu/zaiko-tool/app/services/yahoo_api.py'
with open(path, 'r', encoding='utf-8') as f:
    src = f.read()

OLD = '''    def search_orders(self, days: int = 7, start: int = 1, hits: int = 100) -> dict:
        """直近 days 日の受注を取得して辞書で返す。"""
        access_token = self.get_access_token()

        end_dt   = datetime.now()
        start_dt = end_dt - timedelta(days=days)

        params = {
            'sellerId':      self.seller_id,
            'orderTimeFrom': start_dt.strftime('%Y%m%d%H%M%S'),
            'orderTimeTo':   end_dt.strftime('%Y%m%d%H%M%S'),
            'condition':     '3',   # 全ステータス（元の動作確認済みパラメータ）
            'hits':          str(hits),
            'start':         str(start),
            'output':        'json',
        }

        url = YAHOO_ORDER_SEARCH_URL + '?' + urllib.parse.urlencode(params)
        req = urllib.request.Request(
            url,
            headers={'Authorization': f\'Bearer {access_token}\'}
        )

        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                return json.loads(resp.read().decode(\'utf-8\'))
        except urllib.error.HTTPError as e:
            body = e.read().decode(\'utf-8\', errors=\'replace\')
            raise RuntimeError(f\'受注検索 HTTP {e.code}: {body}\')'''

NEW = '''    def search_orders(self, days: int = 7, start: int = 1, hits: int = 100) -> dict:
        """直近 days 日の受注を取得して辞書で返す。
        Yahoo orderList API: POST + XML ボディ形式。"""
        from xml.etree import ElementTree as ET
        access_token = self.get_access_token()

        end_dt   = datetime.now()
        start_dt = end_dt - timedelta(days=days)

        xml_body = (
            f\'<Req>\\'
            f\'<Search>\\'
            f\'<Result>{hits}</Result>\\'
            f\'<Start>{start}</Start>\\'
            f\'<Condition>\\'
            f\'<SellerId>{self.seller_id}</SellerId>\\'
            f\'<OrderTimeFrom>{start_dt.strftime("%Y%m%d%H%M%S")}</OrderTimeFrom>\\'
            f\'<OrderTimeTo>{end_dt.strftime("%Y%m%d%H%M%S")}</OrderTimeTo>\\'
            f\'</Condition>\\'
            f\'</Search>\\'
            f\'</Req>\'
        )

        req = urllib.request.Request(
            YAHOO_ORDER_SEARCH_URL,
            data=xml_body.encode(\'utf-8\'),
            headers={
                \'Authorization\': f\'Bearer {access_token}\',
                \'Content-Type\':  \'application/xml\',
            }
        )

        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                raw = resp.read().decode(\'utf-8\', errors=\'replace\')
            # XML レスポンスを辞書に変換
            try:
                import xmltodict
                return xmltodict.parse(raw)
            except Exception:
                root = ET.fromstring(raw)
                return {\'_raw_xml\': raw}
        except urllib.error.HTTPError as e:
            body = e.read().decode(\'utf-8\', errors=\'replace\')
            raise RuntimeError(f\'受注検索 HTTP {e.code}: {body}\')'''

if OLD in src:
    src = src.replace(OLD, NEW)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(src)
    print('PATCHED OK')
else:
    print('NOT FOUND - manual edit needed')
    print('First 200 chars of search_orders:')
    i = src.find('def search_orders')
    print(repr(src[i:i+300]))
