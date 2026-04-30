import csv
import io
import time
import unicodedata
from flask import Blueprint, current_app, render_template, request, jsonify, Response
from models import db
from models.order import Order
from models.order_item import OrderItem
from models.purchase import Purchase
from models.ems import Ems
from services.yahoo_api import YahooAPI
from datetime import datetime, timedelta
from sqlalchemy import and_, func, or_

bp = Blueprint('orders', __name__, url_prefix='/orders')

_CURRENT_YAHOO_STATUS_CACHE = {
    'expires_at': 0,
    'ids_by_status': {},
}
_CUSTOMER_ORDER_COUNT_CACHE = {}


# StoreCreator Pro風の画面ステータス（OrderStatus + ShipStatus + IsSeenで判定）
ORDER_STATUS_LABELS = {
    'new_order': '新規注文',
    'wait_payment': '入金待ち',
    'wait_shipping': '出荷待ち',
    'shipping': '出荷処理中',
    'new_reserve': '新規予約',
    'reserve': '予約中',
    'holding': '保留',
    'cancel': 'キャンセル',
    'shipped': '出荷済み',
    'completed': '完了',
}


ORDER_TAB_KEYS = (
    'new_order',
    'wait_payment',
    'wait_shipping',
    'shipping',
    'new_reserve',
    'reserve',
    'holding',
)


ORDER_TAB_LABELS = {
    key: ORDER_STATUS_LABELS[key]
    for key in ORDER_TAB_KEYS
}


SHIP_STATUS_LABELS = {
    '0': '出荷不可',
    '1': '出荷可',
    '2': '出荷処理中',
    '3': '出荷済み',
    '4': '着荷済み',
}


PAY_STATUS_LABELS = {
    '0': '未入金',
    '1': '入金済み',
}


SHIP_COMPANY_LABELS = {
    '999': 'その他',
    '100': '日本郵便',
    '103': 'ヤマト運輸',
    '104': '佐川急便',
    '1000': 'その他',
    '1001': 'ヤマト運輸',
    '1002': '佐川急便',
    '1003': '日本郵便',
    '1004': '西濃運輸',
    '1005': '西武運輸',
    '1006': '福山通運',
    '1007': '名鉄運輸',
    '1008': 'トナミ運輸',
    '1009': '第一貨物',
    '1010': '新潟運輸',
    '1011': '中越運送',
    '1012': '岡山県貨物運送',
    '1013': '久留米運送',
    '1014': '山陽自動車運送',
    '1015': '日本トラック',
    '1016': 'エコ配',
    '1017': 'EMS',
    '1018': 'DHL',
    '1019': 'FedEx',
    '1020': 'UPS',
    '1021': '日通通運',
    '1022': 'TNT',
    '1023': 'OCS',
    '1024': 'USPS',
    '1025': 'SFエクスプレス',
    '1026': 'Aramex',
    '1027': 'SGHグローバル・ジャパン',
    '1028': 'JPロジスティクス',
    '1029': 'MagicalMove',
    '5000': 'ASKUL LOGIST',
    '5001': '受取店舗',
    '0001': 'ヤマト運輸',
    '0002': '佐川急便',
    '0003': '日本郵便',
    '0004': '西濃運輸',
}


ORDERCOUNT_TO_UI_STATUS = {
    'NewOrder': 'new_order',
    'WaitPayment': 'wait_payment',
    'WaitShipping': 'wait_shipping',
    'Shipping': 'shipping',
    'NewReserve': 'new_reserve',
    'Reserve': 'reserve',
    'Holding': 'holding',
}


YAHOO_ID_FILTER_STATUS_KEYS = {
    'new_order',
    'new_reserve',
}


ALLOCATION_STATUS_CHOICES = [
    ('pending', '未発注'),
    ('ordered', '発注済'),
    ('korea_office', '韓国事務所にあり'),
    ('korea_shipping', '日本へ輸送中'),
    ('japan_arrived', '日本入荷済'),
    ('japan_stock', '即納在庫あり'),
]

ALLOCATION_STATUS_LABELS = dict(ALLOCATION_STATUS_CHOICES)


FULFILLMENT_STATUS_LABELS = {
    'shippable': '発送可能',
    'delivery_date_wait': 'お届け希望日待ち',
    'partial_stock_wait': '部分在庫出荷待ち',
    'shipping_wait': '出荷待ち',
}

FULFILLMENT_STATUS_CLASSES = {
    'shippable': 'badge-green',
    'delivery_date_wait': 'badge-blue',
    'partial_stock_wait': 'badge-orange',
    'shipping_wait': 'badge-gray',
}


def _normalize_name(value):
    text = unicodedata.normalize('NFKC', str(value or ''))
    return ''.join(ch for ch in text if ch.isalnum()).casefold()


def _split_customer_name(value):
    parts = str(value or '').strip().split()
    if len(parts) >= 2:
        return parts[0], ' '.join(parts[1:])
    return '', ''


def _first_text(*values):
    for value in values:
        if value is None:
            continue
        text = str(value).strip()
        if text:
            return text
    return ''


def _parse_api_bool(value):
    return str(value or '').strip().lower() in ('true', '1', 'yes')


def _api_order_time_value(order):
    raw = _first_text(order.get('OrderTime'), order.get('order_time'))
    try:
        return datetime.fromisoformat(raw.replace('Z', '+00:00')).timestamp()
    except Exception:
        return 0


def _api_order_ui_status_key(order):
    order_status = _first_text(order.get('OrderStatus'), order.get('order_status'))
    ship_status = _first_text(order.get('ShipStatus'), order.get('ship_status'))
    pay_status = _first_text(order.get('PayStatus'), order.get('pay_status'))
    is_seen = _parse_api_bool(_first_text(order.get('IsSeen'), order.get('is_seen')))

    if order_status == '1':
        return 'reserve' if is_seen else 'new_reserve'
    if order_status == '3':
        return 'holding'
    if order_status == '2':
        if not is_seen:
            return 'new_order'
        if pay_status and pay_status != '1':
            return 'wait_payment'
        if ship_status == '2':
            return 'shipping'
        if ship_status in ('0', '1', ''):
            return 'wait_shipping'
    return ''


def _current_yahoo_ids_by_status(os_counts, force_refresh=False, target_status=None):
    now = time.time()
    cached_ids = _CURRENT_YAHOO_STATUS_CACHE['ids_by_status']
    if (not force_refresh) and _CURRENT_YAHOO_STATUS_CACHE['expires_at'] > now:
        if target_status:
            return {target_status: cached_ids.get(target_status, [])}
        return cached_ids

    ids_by_status = {key: [] for key in ORDER_TAB_KEYS}
    try:
        api = YahooAPI()
        # 新規系タブは軽量取得（直近2日・先頭300件）で即応性を優先
        if target_status in YAHOO_ID_FILTER_STATUS_KEYS:
            result = api.search_orders(days=2, start=1, hits=300)
            orders = YahooAPI._extract_orders(result)
        else:
            orders = api.fetch_all_orders(days=7)
    except Exception:
        return None

    rows_by_status = {key: [] for key in ORDER_TAB_KEYS}
    for order in orders:
        key = _api_order_ui_status_key(order)
        order_id = _first_text(order.get('OrderId'), order.get('order_id'))
        if key in rows_by_status and order_id:
            rows_by_status[key].append((_api_order_time_value(order), order_id))

    for key, rows in rows_by_status.items():
        rows.sort(reverse=True)
        expected = os_counts.get(key)
        if key in YAHOO_ID_FILTER_STATUS_KEYS and key != 'new_order' and expected is not None:
            rows = rows[:max(0, int(expected))]
        ids_by_status[key] = [order_id for _ts, order_id in rows]

    _CURRENT_YAHOO_STATUS_CACHE['expires_at'] = now + 120
    _CURRENT_YAHOO_STATUS_CACHE['ids_by_status'] = ids_by_status
    if target_status:
        return {target_status: ids_by_status.get(target_status, [])}
    return ids_by_status


def _ensure_current_yahoo_orders_loaded(yahoo_order_ids, max_ids=120):
    """表示対象のYahoo注文IDがDBに無い場合だけ詳細APIから補完する。"""
    order_ids = []
    seen = set()
    for order_id in yahoo_order_ids or []:
        order_id = str(order_id or '').strip()
        if order_id and order_id not in seen:
            order_ids.append(order_id)
            seen.add(order_id)
    if not order_ids:
        return
    if max_ids and len(order_ids) > max_ids:
        order_ids = order_ids[:max_ids]

    existing_ids = {
        row[0] for row in
        db.session.query(Order.yahoo_order_id)
        .filter(Order.yahoo_order_id.in_(order_ids))
        .all()
    }
    missing_ids = [order_id for order_id in order_ids if order_id not in existing_ids]
    if not missing_ids:
        return

    try:
        from routes.import_data import _upsert_yahoo_order
        api = YahooAPI()
        for i in range(0, len(missing_ids), 100):
            for detail_raw in api.get_order_details_bulk(missing_ids[i:i + 100]):
                normalized = YahooAPI.normalize_order_info(detail_raw)
                if normalized.get('OrderId'):
                    _upsert_yahoo_order(normalized)
        db.session.commit()
    except Exception as e:
        db.session.rollback()
        current_app.logger.warning('current Yahoo order sync failed: %s', e)


def _yahoo_customer_order_counts_for_orders(orders, enabled=False):
    """少件数タブではYahoo検索APIの過去1年件数を優先する。"""
    if not enabled or not orders or len(orders) > 25:
        return {}

    api = YahooAPI()
    now = time.time()
    counts_by_order_id = {}
    queried = 0

    for order in orders:
        last_name, first_name = _split_customer_name(order.customer_name)
        cache_key = _normalize_name(f'{last_name}{first_name}')
        if not cache_key:
            continue

        cached = _CUSTOMER_ORDER_COUNT_CACHE.get(cache_key)
        if cached and cached['expires_at'] > now:
            counts_by_order_id[order.id] = cached['count']
            continue

        if queried:
            time.sleep(1.05)
        try:
            count = api.count_orders_by_bill_name(last_name, first_name, days=365)
            _CUSTOMER_ORDER_COUNT_CACHE[cache_key] = {
                'count': count,
                'expires_at': now + 3600,
            }
            counts_by_order_id[order.id] = count
            queried += 1
        except Exception as e:
            current_app.logger.warning('Yahoo customer order count failed: %s', e)

    return counts_by_order_id


def order_ui_status_key(order):
    order_status = str(order.yahoo_order_status or '')
    ship_status = str(order.yahoo_ship_status or '')
    pay_status = str(order.yahoo_pay_status or '')
    is_seen = bool(order.is_seen)

    if order_status == '1':
        return 'reserve' if is_seen else 'new_reserve'
    if order_status == '3':
        return 'holding'
    if order_status == '4':
        return 'cancel'
    if order_status == '5':
        return 'completed'
    if order_status == '2':
        if not is_seen:
            return 'new_order'
        if pay_status and pay_status != '1':
            return 'wait_payment'
        if ship_status == '2':
            return 'shipping'
        if ship_status in ('0', '1', ''):
            return 'wait_shipping'
        if ship_status in ('3', '4'):
            return 'shipped'
    if ship_status in ('3', '4'):
        return 'shipped'
    return ''


def order_ui_status_label(order):
    return ORDER_STATUS_LABELS.get(order_ui_status_key(order), '—')


def ship_status_label(code):
    code = str(code or '')
    return SHIP_STATUS_LABELS.get(code, code or '—')


def pay_status_label(code):
    code = str(code or '')
    return PAY_STATUS_LABELS.get(code, '未入金')


def ship_company_label(code):
    code = str(code or '')
    return SHIP_COMPANY_LABELS.get(code, code or '—')


def _apply_ui_status_filter(query, key):
    if key == 'new_order':
        return query.filter(Order.yahoo_order_status == '2', Order.is_seen.is_(False))
    if key == 'wait_payment':
        return query.filter(
            Order.yahoo_order_status == '2',
            or_(Order.is_seen.is_(True), Order.is_seen.is_(None)),
            or_(Order.yahoo_pay_status != '1', Order.yahoo_pay_status.is_(None), Order.yahoo_pay_status == ''),
        )
    if key == 'wait_shipping':
        return query.filter(
            Order.yahoo_order_status == '2',
            or_(Order.is_seen.is_(True), Order.is_seen.is_(None)),
            or_(Order.yahoo_pay_status == '1', Order.yahoo_pay_status.is_(None), Order.yahoo_pay_status == ''),
            or_(Order.yahoo_ship_status.in_(('0', '1')), Order.yahoo_ship_status.is_(None), Order.yahoo_ship_status == ''),
        )
    if key == 'shipping':
        return query.filter(Order.yahoo_order_status == '2', Order.yahoo_ship_status == '2')
    if key == 'shipped':
        return query.filter(Order.yahoo_ship_status.in_(('3', '4')))
    if key == 'completed':
        return query.filter(Order.yahoo_order_status == '5')
    if key == 'new_reserve':
        return query.filter(Order.yahoo_order_status == '1', Order.is_seen.is_(False))
    if key == 'reserve':
        return query.filter(Order.yahoo_order_status == '1', or_(Order.is_seen.is_(True), Order.is_seen.is_(None)))
    if key == 'holding':
        return query.filter(Order.yahoo_order_status == '3')
    if key == 'cancel':
        return query.filter(Order.yahoo_order_status == '4')
    if key:
        return query.filter(Order.yahoo_order_status == key)
    return query


def _active_order_ids():
    order_ids = []
    for order in Order.query.with_entities(
        Order.id,
        Order.yahoo_order_status,
        Order.yahoo_ship_status,
        Order.yahoo_pay_status,
        Order.is_seen,
    ).all():
        if order_ui_status_key(order) in ORDER_TAB_KEYS:
            order_ids.append(order.id)
    return order_ids


def inventory_type_label(item):
    raw = str(item.inventory_type or '').strip()
    if raw in ('即納', 'お取り寄せ'):
        return raw
    # 新規取込直後など、まだ在庫区分が空/旧値の場合だけ初期判定する。
    return 'お取り寄せ' if (item.product_sub_code or '').strip().lower().endswith('b') else '即納'


def allocation_status_key(item, purchase_item_ids=None, ems_state_by_item=None):
    purchase_item_ids = purchase_item_ids or set()
    ems_state_by_item = ems_state_by_item or {}

    ems_state = ems_state_by_item.get(item.id)
    if ems_state == 'japan_arrived':
        return 'japan_arrived'
    if item.status in ALLOCATION_STATUS_LABELS and item.status != 'pending':
        return item.status
    if ems_state in ('korea_office', 'korea_shipping'):
        return ems_state
    if inventory_type_label(item) == '即納' or item.status in ('fully_allocated', 'allocated_sokunou', 'shipped'):
        return 'japan_stock'
    if item.id in purchase_item_ids or item.status in ('provisional_allocated', 'partial_waiting', 'priority_hold'):
        return 'ordered'
    return 'pending'


def allocation_status_label(item, purchase_item_ids=None, ems_state_by_item=None):
    key = allocation_status_key(item, purchase_item_ids, ems_state_by_item)
    return ALLOCATION_STATUS_LABELS.get(key, '—')


def fulfillment_status_key(order, items, purchase_item_ids=None, ems_state_by_item=None):
    if not items:
        return 'shipping_wait'

    ready_keys = {'japan_stock', 'japan_arrived'}
    item_statuses = [
        allocation_status_key(item, purchase_item_ids, ems_state_by_item)
        for item in items
    ]
    ready_count = sum(1 for key in item_statuses if key in ready_keys)

    if ready_count == len(item_statuses):
        if order.desired_delivery_date:
            days_until_delivery = (order.desired_delivery_date - datetime.now().date()).days
            if days_until_delivery >= 5:
                return 'delivery_date_wait'
        return 'shippable'
    if ready_count > 0:
        return 'partial_stock_wait'
    return 'shipping_wait'


def fulfillment_status_label(order, items, purchase_item_ids=None, ems_state_by_item=None):
    key = fulfillment_status_key(order, items, purchase_item_ids, ems_state_by_item)
    return FULFILLMENT_STATUS_LABELS.get(key, '出荷待ち')


def fulfillment_status_class(order, items, purchase_item_ids=None, ems_state_by_item=None):
    key = fulfillment_status_key(order, items, purchase_item_ids, ems_state_by_item)
    return FULFILLMENT_STATUS_CLASSES.get(key, 'badge-gray')


def _order_ids_for_allocation_filter(status_key):
    from models.ems_item import EmsItem

    order_ids = set()
    if status_key == 'pending':
        order_ids.update(row[0] for row in db.session.query(OrderItem.order_id).filter(
            OrderItem.status.in_(('pending', 'shortage'))
        ).all())
    elif status_key == 'ordered':
        order_ids.update(row[0] for row in db.session.query(OrderItem.order_id).filter(
            OrderItem.status.in_(('ordered', 'provisional_allocated', 'partial_waiting', 'priority_hold'))
        ).all())
        order_ids.update(row[0] for row in db.session.query(OrderItem.order_id).join(
            Purchase, Purchase.order_item_id == OrderItem.id
        ).all())
    elif status_key == 'korea_office':
        order_ids.update(row[0] for row in db.session.query(OrderItem.order_id).filter(
            OrderItem.status == 'korea_office'
        ).all())
        order_ids.update(row[0] for row in db.session.query(OrderItem.order_id).join(
            EmsItem, EmsItem.order_item_id == OrderItem.id
        ).join(Ems, EmsItem.ems_id == Ems.id).filter(
            and_(Ems.status != 'arrived', Ems.arrived_at.is_(None), Ems.shipped_at.is_(None))
        ).all())
    elif status_key == 'korea_shipping':
        order_ids.update(row[0] for row in db.session.query(OrderItem.order_id).filter(
            OrderItem.status == 'korea_shipping'
        ).all())
        order_ids.update(row[0] for row in db.session.query(OrderItem.order_id).join(
            EmsItem, EmsItem.order_item_id == OrderItem.id
        ).join(Ems, EmsItem.ems_id == Ems.id).filter(
            and_(Ems.status != 'arrived', Ems.arrived_at.is_(None), Ems.shipped_at.isnot(None))
        ).all())
    elif status_key == 'japan_arrived':
        order_ids.update(row[0] for row in db.session.query(OrderItem.order_id).filter(
            OrderItem.status == 'japan_arrived'
        ).all())
        order_ids.update(row[0] for row in db.session.query(OrderItem.order_id).join(
            EmsItem, EmsItem.order_item_id == OrderItem.id
        ).join(Ems, EmsItem.ems_id == Ems.id).filter(or_(
            Ems.status == 'arrived',
            Ems.arrived_at.isnot(None),
        )).all())
    elif status_key == 'japan_stock':
        order_ids.update(row[0] for row in db.session.query(OrderItem.order_id).filter(or_(
            OrderItem.status.in_(('japan_stock', 'fully_allocated', 'allocated_sokunou', 'shipped')),
            OrderItem.inventory_type == '即納',
        )).all())
    elif status_key in ALLOCATION_STATUS_LABELS:
        order_ids.update(row[0] for row in db.session.query(OrderItem.order_id).filter(
            OrderItem.status == status_key
        ).all())
    return list(order_ids)


@bp.route('/')
def index():
    q              = request.args.get('q', '').strip()
    status_filter  = request.args.get('status', '')   # item status（内部引当）
    os_filter      = request.args.get('os', '')       # yahoo_order_status
    if os_filter and os_filter not in ORDER_TAB_KEYS:
        os_filter = ''

    # ステータス集計（タブ用）
    os_counts = {key: 0 for key in ORDER_TAB_KEYS}
    count_orders = Order.query.all()
    for count_order in count_orders:
        key = order_ui_status_key(count_order)
        if key in os_counts:
            os_counts[key] = os_counts.get(key, 0) + 1
    try:
        raw_counts = YahooAPI().get_order_count()
        for api_key, ui_key in ORDERCOUNT_TO_UI_STATUS.items():
            if ui_key in os_counts:
                os_counts[ui_key] = int(raw_counts.get(api_key) or 0)
    except Exception:
        pass
    # 描画速度優先: orderList全件取得は新規タブ表示時のみ実行する。
    needs_current_yahoo_ids = os_filter in YAHOO_ID_FILTER_STATUS_KEYS
    # 新規注文タブはF5で即時反映したいため、キャッシュを使わず毎回取り直す。
    force_refresh_ids = os_filter == 'new_order'
    target_status = os_filter if os_filter in YAHOO_ID_FILTER_STATUS_KEYS else None
    ids_by_status = _current_yahoo_ids_by_status(
        os_counts,
        force_refresh=force_refresh_ids,
        target_status=target_status,
    ) if needs_current_yahoo_ids else None
    total_count = sum(os_counts.values())

    # 商品コードで絞る場合は order_items から引く
    if q:
        item_order_ids = [
            oi.order_id for oi in
            OrderItem.query.filter(OrderItem.product_code.contains(q)).all()
        ]
        query = Order.query.filter(
            or_(
                Order.yahoo_order_id.contains(q),
                Order.customer_name.contains(q),
                Order.id.in_(item_order_ids),
            )
        )
    else:
        query = Order.query
    if status_filter:
        if status_filter in dict(ALLOCATION_STATUS_CHOICES):
            item_order_ids = _order_ids_for_allocation_filter(status_filter)
        else:
            item_order_ids = [
                oi.order_id for oi in
                OrderItem.query.filter_by(status=status_filter).all()
            ]
        query = query.filter(Order.id.in_(item_order_ids))
    if os_filter:
        current_yahoo_ids = ids_by_status.get(os_filter) if ids_by_status is not None else None
        if os_filter in YAHOO_ID_FILTER_STATUS_KEYS and current_yahoo_ids is not None:
            _ensure_current_yahoo_orders_loaded(current_yahoo_ids, max_ids=120)
            query = query.filter(Order.yahoo_order_id.in_(current_yahoo_ids))
        elif os_filter in YAHOO_ID_FILTER_STATUS_KEYS:
            # orderList が認証エラー等で現在IDを取れない時に、古い未読注文を混ぜない。
            query = query.filter(False)
        else:
            query = _apply_ui_status_filter(query, os_filter)
    else:
        active_order_ids = _active_order_ids()
        query = query.filter(Order.id.in_(active_order_ids))

    # ページネーション
    page      = int(request.args.get('page', 1))
    per_page  = 100
    total     = query.count()
    orders    = (query
        .order_by(Order.ordered_at.desc())
        .offset((page - 1) * per_page)
        .limit(per_page)
        .all()
    )
    total_pages = (total + per_page - 1) // per_page
    # 精度優先: 新規注文/新規予約タブはYahooの顧客別回数を優先する。
    # ただし件数が多いと遅くなるので、ページ内が少ない時のみ有効化する。
    yahoo_enabled = (os_filter in YAHOO_ID_FILTER_STATUS_KEYS) and len(orders) <= 25
    yahoo_customer_counts = _yahoo_customer_order_counts_for_orders(
        orders,
        enabled=yahoo_enabled,
    )

    # 明細を一括フェッチ（N+1回避）
    order_ids   = [o.id for o in orders]
    yahoo_order_ids = [o.yahoo_order_id for o in orders]
    items_bulk  = OrderItem.query.filter(OrderItem.order_id.in_(order_ids)).all() if order_ids else []
    items_by_order = {}
    for it in items_bulk:
        items_by_order.setdefault(it.order_id, []).append(it)
    all_item_ids = [it.id for it in items_bulk]
    all_item_id_set = set(all_item_ids)
    item_codes = [it.product_code for it in items_bulk if it.product_code]
    order_id_to_yahoo = {o.id: o.yahoo_order_id for o in orders}
    item_key_to_ids = {}
    for it in items_bulk:
        key = (order_id_to_yahoo.get(it.order_id, ''), it.product_code, it.product_sub_code or '')
        item_key_to_ids.setdefault(key, []).append(it.id)
        if it.product_sub_code:
            item_key_to_ids.setdefault((key[0], key[1], ''), []).append(it.id)

    purchase_item_ids = set()
    if all_item_ids:
        purchase_rows = Purchase.query.filter(or_(
            Purchase.order_item_id.in_(all_item_ids),
            Purchase.order_id.in_(yahoo_order_ids),
            Purchase.product_code.in_(item_codes),
        )).all()
        for p in purchase_rows:
            if p.order_item_id in all_item_id_set:
                purchase_item_ids.add(p.order_item_id)
                continue
            matched_ids = item_key_to_ids.get((p.order_id or '', p.product_code, p.product_sub_code or ''), [])
            if not matched_ids:
                matched_ids = item_key_to_ids.get((p.order_id or '', p.product_code, ''), [])
            purchase_item_ids.update(matched_ids)

    ems_by_item = {}  # item_id → set of ems_number
    ems_state_by_item = {}  # item_id → arrived / in_transit
    if all_item_ids:
        from models.ems_item import EmsItem as EI
        eis = EI.query.outerjoin(Ems, EI.ems_id == Ems.id).filter(or_(
            EI.order_item_id.in_(all_item_ids),
            Ems.order_id.in_(yahoo_order_ids),
            EI.product_code.in_(item_codes),
        )).all()
        for ei in eis:
            matched_ids = []
            if ei.order_item_id in all_item_id_set:
                matched_ids.append(ei.order_item_id)
            ems_order_id = ei.ems.order_id if ei.ems else ''
            if ems_order_id:
                matched_ids.extend(item_key_to_ids.get((ems_order_id, ei.product_code, ei.product_sub_code or ''), []))
                matched_ids.extend(item_key_to_ids.get((ems_order_id, ei.product_code, ''), []))
            if ei.ems and (ei.ems.status == 'arrived' or ei.ems.arrived_at):
                state = 'japan_arrived'
            elif ei.ems and ei.ems.shipped_at:
                state = 'korea_shipping'
            else:
                state = 'korea_office'
            state_priority = {'korea_office': 1, 'korea_shipping': 2, 'japan_arrived': 3}
            for item_id in set(matched_ids):
                if ei.ems and ei.ems.ems_number:
                    ems_by_item.setdefault(item_id, set()).add(ei.ems.ems_number)
                if state_priority.get(state, 0) > state_priority.get(ems_state_by_item.get(item_id), 0):
                    ems_state_by_item[item_id] = state

    # グルーピング
    grouped = []
    for o in orders:
        items = items_by_order.get(o.id, [])
        ems_nums = set()
        for it in items:
            ems_nums |= ems_by_item.get(it.id, set())
        grouped.append({
            'order': o,
            'order_items': items,
            'ems_numbers': '、'.join(sorted(ems_nums)) if ems_nums else '',
        })

    one_year_ago = datetime.now() - timedelta(days=365)
    customer_names_raw = (
        db.session.query(Order.id, Order.customer_name)
        .filter(and_(Order.customer_name.isnot(None), Order.customer_name != ''))
        .filter(Order.ordered_at >= one_year_ago)
        .all()
    )
    customer_order_counts = {}
    for _order_id, customer_name in customer_names_raw:
        key = _normalize_name(customer_name)
        if not key:
            continue
        customer_order_counts[key] = customer_order_counts.get(key, 0) + 1

    def customer_order_count(order):
        if yahoo_enabled:
            # Yahoo件数を使うモードでは、取れなかった場合は0とする（ローカル誤差を避ける）
            return yahoo_customer_counts.get(order.id, 0)
        if order.id in yahoo_customer_counts:
            return yahoo_customer_counts[order.id]
        key = _normalize_name(order.customer_name)
        if not key:
            return None
        return customer_order_counts.get(key, 1)

    return render_template('orders.html',
        grouped=grouped, q=q, status_filter=status_filter,
        os_filter=os_filter, os_counts=os_counts,
        total_count=total_count, total_pages=total_pages,
        page=page, per_page=per_page, filtered_total=total,
        order_status_labels=ORDER_TAB_LABELS,
        customer_order_count=customer_order_count,
        allocation_status_choices=ALLOCATION_STATUS_CHOICES,
        allocation_status_key=lambda item: allocation_status_key(item, purchase_item_ids, ems_state_by_item),
        fulfillment_status_label=lambda order, items: fulfillment_status_label(order, items, purchase_item_ids, ems_state_by_item),
        fulfillment_status_class=lambda order, items: fulfillment_status_class(order, items, purchase_item_ids, ems_state_by_item),
        inventory_type_label=inventory_type_label,
        order_ui_status_label=order_ui_status_label,
        ship_status_label=ship_status_label,
        pay_status_label=pay_status_label,
        ship_company_label=ship_company_label,
    )


@bp.route('/detail/<yahoo_order_id>')
def detail(yahoo_order_id):
    o = Order.query.filter_by(yahoo_order_id=yahoo_order_id).first_or_404()
    items = OrderItem.query.filter_by(order_id=o.id).order_by(OrderItem.id.asc()).all()
    one_year_ago = datetime.now() - timedelta(days=365)
    past_year_order_count = 0
    if o.customer_name:
        target_key = _normalize_name(o.customer_name)
        names = (
            db.session.query(Order.customer_name)
            .filter(and_(Order.customer_name.isnot(None), Order.customer_name != ''))
            .filter(Order.ordered_at >= one_year_ago)
            .all()
        )
        past_year_order_count = sum(
            1
            for (customer_name,) in names
            if _normalize_name(customer_name) == target_key
        )
    return render_template(
        'order_detail.html',
        order=o,
        items=items,
        past_year_order_count=past_year_order_count,
        order_ui_status_label=order_ui_status_label,
        ship_status_label=ship_status_label,
        pay_status_label=pay_status_label,
        ship_company_label=ship_company_label,
    )


@bp.route('/allocate', methods=['POST'])
def allocate():
    """引当実行（受注リストから）"""
    item_ids = request.json.get('item_ids', [])
    from services.allocation import run_auto_allocation
    stats = run_auto_allocation(item_ids=item_ids)
    db.session.commit()
    return jsonify({
        'results': [{'item_id': item_id} for item_id in item_ids],
        'stats': stats,
    })


@bp.route('/update/order/<int:order_id>', methods=['POST'])
def update_order(order_id):
    o = Order.query.get_or_404(order_id)
    data = request.get_json()
    field, value = data.get('field'), data.get('value', '')
    allowed = {'customer_name', 'yahoo_ship_status', 'status', 'delay_memo', 'desired_delivery_date'}
    if field not in allowed:
        return jsonify({'ok': False, 'error': 'field not allowed'}), 400
    try:
        if field == 'desired_delivery_date':
            o.desired_delivery_date = datetime.strptime(value, '%Y-%m-%d').date() if value else None
        else:
            setattr(o, field, value)
        o.updated_at = datetime.now()
        db.session.commit()
        return jsonify({'ok': True})
    except Exception as e:
        db.session.rollback()
        return jsonify({'ok': False, 'error': str(e)}), 500


@bp.route('/update/item/<int:item_id>', methods=['POST'])
def update_item(item_id):
    item = OrderItem.query.get_or_404(item_id)
    data = request.get_json() or {}
    field, value = data.get('field'), data.get('value', '')
    try:
        _apply_item_field_update(item, field, value)
        db.session.commit()
        return jsonify({'ok': True})
    except Exception as e:
        db.session.rollback()
        return jsonify({'ok': False, 'error': str(e)}), 500


def _apply_item_field_update(item, field, value):
    allowed = {'product_code', 'quantity', 'status', 'allocated_qty', 'inventory_type'}
    if field not in allowed:
        raise ValueError('field not allowed')
    if field in ('quantity', 'allocated_qty'):
        setattr(item, field, int(value))
    elif field == 'inventory_type':
        if value not in ('即納', 'お取り寄せ'):
            raise ValueError('inventory_type not allowed')
        item.inventory_type = value
    else:
        setattr(item, field, value)
    item.updated_at = datetime.now()


@bp.route('/update/items/bulk', methods=['POST'])
def update_items_bulk():
    data = request.get_json() or {}
    updates = data.get('updates') or []
    if not isinstance(updates, list) or not updates:
        return jsonify({'ok': False, 'error': 'updates required'}), 400
    if len(updates) > 200:
        return jsonify({'ok': False, 'error': 'too many updates'}), 400

    try:
        item_ids = []
        for upd in updates:
            if not isinstance(upd, dict):
                raise ValueError('invalid update payload')
            item_id = int(upd.get('item_id'))
            item_ids.append(item_id)

        items = OrderItem.query.filter(OrderItem.id.in_(item_ids)).all()
        item_map = {item.id: item for item in items}

        updated = 0
        for upd in updates:
            item_id = int(upd.get('item_id'))
            item = item_map.get(item_id)
            if not item:
                continue
            _apply_item_field_update(item, upd.get('field'), upd.get('value', ''))
            updated += 1

        db.session.commit()
        return jsonify({'ok': True, 'updated': updated})
    except Exception as e:
        db.session.rollback()
        return jsonify({'ok': False, 'error': str(e)}), 500


@bp.route('/csv')
def export_csv():
    orders = Order.query.order_by(Order.ordered_at.desc()).all()
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(['注文番号', '注文日', 'お客様名', '商品コード', '商品名', '数量',
                     '在庫種別', '引当ステータス', '引当数', '経過日数'])

    for o in orders:
        items = OrderItem.query.filter_by(order_id=o.id).all()
        for item in items:
            writer.writerow([
                o.yahoo_order_id,
                o.ordered_at.strftime('%Y/%m/%d'),
                o.customer_name or '',
                item.product_code,
                item.product_name or '',
                item.quantity,
                item.inventory_type,
                item.status_label,
                item.allocated_qty,
                o.days_elapsed,
            ])

    output.seek(0)
    return Response(
        output.getvalue().encode('utf-8-sig'),
        mimetype='text/csv',
        headers={'Content-Disposition': f'attachment; filename=orders_{datetime.now().strftime("%Y%m%d")}.csv'}
    )
