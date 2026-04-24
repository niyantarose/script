#!/usr/bin/env python3
"""Fix os_counts to use orderCount API with fallback to DB."""
orders_path = '/home/ubuntu/zaiko-tool/app/routes/orders.py'
with open(orders_path, 'r', encoding='utf-8') as f:
    src = f.read()

# Add ORDERCOUNT_TO_DBSTATUS mapping after ORDERCOUNT_LABELS if not there
MAPPING_CODE = """
ORDERCOUNT_TO_DBSTATUS = {
    'Reserve':      '1',
    'NewOrder':     '2',
    'Holding':      '3',
    'WaitShipping': '4',
}
"""
if 'ORDERCOUNT_TO_DBSTATUS' not in src:
    if 'ORDERCOUNT_LABELS' in src:
        import re
        m = re.search(r'(ORDERCOUNT_LABELS\s*=\s*\{[^}]+\})', src)
        if m:
            src = src[:m.end()] + MAPPING_CODE + src[m.end():]
            print('Inserted ORDERCOUNT_TO_DBSTATUS')
    else:
        print('ORDERCOUNT_LABELS not found, inserting at ORDER_STATUS_LABELS')
        import re
        m = re.search(r'(ORDER_STATUS_LABELS\s*=\s*\{[^}]+\})', src)
        if m:
            src = src[:m.end()] + MAPPING_CODE + src[m.end():]
            print('Inserted ORDERCOUNT_TO_DBSTATUS after ORDER_STATUS_LABELS')
else:
    print('ORDERCOUNT_TO_DBSTATUS already exists')

# Replace the DB-based os_counts block
OLD = """    # ステータス集計（タブ用）
    from sqlalchemy import func
    os_counts_raw = (
        db.session.query(Order.yahoo_order_status, func.count(Order.id))
        .group_by(Order.yahoo_order_status)
        .all()
    )
    os_counts = {(row[0] or ''): row[1] for row in os_counts_raw}
    total_count = sum(os_counts.values())"""

NEW = """    # ステータス集計（タブ用） - Yahoo orderCount API を使用
    from sqlalchemy import func
    try:
        _api = YahooAPI()
        _raw_counts = _api.get_order_count()
        os_counts = {}
        for _api_key, _db_code in ORDERCOUNT_TO_DBSTATUS.items():
            _v = _raw_counts.get(_api_key, '0')
            os_counts[_db_code] = int(_v) if _v else 0
        total_count = sum(os_counts.values())
    except Exception:
        os_counts_raw = (
            db.session.query(Order.yahoo_order_status, func.count(Order.id))
            .group_by(Order.yahoo_order_status)
            .all()
        )
        os_counts = {(row[0] or ''): row[1] for row in os_counts_raw}
        total_count = sum(os_counts.values())"""

if OLD in src:
    src = src.replace(OLD, NEW, 1)
    print('Replaced os_counts with API call + fallback')
else:
    print('ERROR: OLD block not found exactly — check orders.py')
    # Print the relevant section for debugging
    i = src.find('os_counts_raw')
    print(repr(src[max(0,i-100):i+300]))

with open(orders_path, 'w', encoding='utf-8') as f:
    f.write(src)
print('Done')
