#!/usr/bin/env python3
"""Add get_order_count() to yahoo_api.py and update orders route."""
import re, os

# ── 1. yahoo_api.py ──────────────────────────────────────────────────────────
api_path = '/home/ubuntu/zaiko-tool/app/services/yahoo_api.py'
with open(api_path, 'r', encoding='utf-8') as f:
    src = f.read()

NEW_METHOD = '''
    def get_order_count(self) -> dict:
        """Yahoo orderCount API (GET) for real-time status counts.
        Returns e.g. {'NewOrder': '0', 'WaitShipping': '157', 'Reserve': '467', ...}"""
        from xml.etree import ElementTree as ET
        access_token = self.get_access_token()
        url = (
            'https://circus.shopping.yahooapis.jp/ShoppingWebService/V1/orderCount'
            '?sellerId=' + urllib.parse.quote(self.seller_id)
        )
        headers = {'Authorization': 'Bearer ' + access_token}
        try:
            headers.update(self._sws_headers())
        except Exception:
            pass
        req = urllib.request.Request(url, headers=headers)
        try:
            with urllib.request.urlopen(req, timeout=15) as resp:
                raw = resp.read().decode('utf-8', errors='replace')
            root = ET.fromstring(raw)
            count = root.find('.//Count')
            if count is None:
                return {}
            return {child.tag: (child.text or '0') for child in count}
        except Exception as e:
            import logging
            logging.warning('orderCount error: %s', e)
            return {}
'''

# Insert before get_order_detail or at end of class
anchor = '    def get_order_detail('
if anchor in src:
    src = src.replace(anchor, NEW_METHOD + '\n' + anchor, 1)
    print('Inserted get_order_count before get_order_detail')
elif 'def get_order_count' not in src:
    # append before last line
    src = src.rstrip() + '\n' + NEW_METHOD + '\n'
    print('Appended get_order_count')
else:
    print('get_order_count already exists — skipping yahoo_api.py')

with open(api_path, 'w', encoding='utf-8') as f:
    f.write(src)

# ── 2. orders.py ─────────────────────────────────────────────────────────────
orders_path = '/home/ubuntu/zaiko-tool/app/routes/orders.py'
with open(orders_path, 'r', encoding='utf-8') as f:
    osrc = f.read()

# Replace ORDER_STATUS_LABELS and tab count logic
# We need to:
#   a) keep ORDER_STATUS_LABELS for badge rendering (yahoo_order_status int codes)
#   b) add ORDERCOUNT_LABELS for the new API-based tabs
#   c) replace os_counts GROUP BY with api call

OLD_LABELS = "ORDER_STATUS_LABELS = {"
if OLD_LABELS not in osrc:
    print('ORDER_STATUS_LABELS not found — check orders.py manually')
else:
    print('orders.py found, updating tab counts...')

# Add import for YahooAPI near top if not already there
if 'from services.yahoo_api import YahooAPI' not in osrc and 'YahooAPI' not in osrc:
    osrc = osrc.replace(
        'from flask import',
        'from services.yahoo_api import YahooAPI\nfrom flask import',
        1
    )

# Add ORDERCOUNT_LABELS constant after ORDER_STATUS_LABELS block
ORDERCOUNT_LABELS_CODE = """
ORDERCOUNT_LABELS = {
    'NewOrder':    '新規注文',
    'WaitPayment': '入金待ち',
    'WaitShipping': '出荷待ち',
    'Shipping':    '出荷処理中',
    'Reserve':     '予約中',
    'Holding':     '保留',
}
"""

if 'ORDERCOUNT_LABELS' not in osrc:
    # insert after ORDER_STATUS_LABELS closing brace
    pattern = r"(ORDER_STATUS_LABELS\s*=\s*\{[^}]+\})"
    m = re.search(pattern, osrc)
    if m:
        osrc = osrc[:m.end()] + ORDERCOUNT_LABELS_CODE + osrc[m.end():]
        print('Inserted ORDERCOUNT_LABELS')
    else:
        print('Could not find ORDER_STATUS_LABELS block via regex')

# Now update index() to use get_order_count() for tab counts
# Find the os_counts GROUP BY block and replace it
OLD_OS_COUNTS = """    os_counts = dict(
        db.session.query(Order.yahoo_order_status, func.count(Order.id))
        .group_by(Order.yahoo_order_status)
        .all()
    )"""

NEW_OS_COUNTS = """    # Use Yahoo orderCount API for real-time tab counts
    try:
        _api = YahooAPI()
        _raw_counts = _api.get_order_count()
    except Exception:
        _raw_counts = {}
    os_counts = _raw_counts  # keyed by orderCount field names e.g. 'NewOrder'"""

if OLD_OS_COUNTS in osrc:
    osrc = osrc.replace(OLD_OS_COUNTS, NEW_OS_COUNTS, 1)
    print('Replaced os_counts GROUP BY with API call')
else:
    # Try without exact whitespace
    if "db.session.query(Order.yahoo_order_status, func.count(Order.id))" in osrc:
        print('os_counts block found but whitespace differs — manual fix needed')
    else:
        print('os_counts GROUP BY block NOT found — may already be updated or different code')

# Pass ORDERCOUNT_LABELS to template
OLD_RENDER = "order_status_labels=ORDER_STATUS_LABELS,"
NEW_RENDER  = "order_status_labels=ORDER_STATUS_LABELS,\n        ordercount_labels=ORDERCOUNT_LABELS,"
if OLD_RENDER in osrc and 'ordercount_labels' not in osrc:
    osrc = osrc.replace(OLD_RENDER, NEW_RENDER, 1)
    print('Added ordercount_labels to render_template call')

with open(orders_path, 'w', encoding='utf-8') as f:
    f.write(osrc)

print('Done patching orders.py')
