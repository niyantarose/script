from datetime import datetime

from models import db, Order, OrderItem, Alert, StockTransaction
from services.stock_ledger import (
    apply_order_out, sync_cancel_returns, get_balance, record_transaction,
)


def _make_order(yahoo_order_id, order_status='2', ship_status='0'):
    o = Order(yahoo_order_id=yahoo_order_id, ordered_at=datetime.now(),
              yahoo_order_status=order_status, yahoo_ship_status=ship_status)
    db.session.add(o)
    db.session.flush()
    return o


def _make_item(order, product_code='P-001', qty=2):
    oi = OrderItem(order_id=order.id, product_code=product_code,
                   quantity=qty, inventory_type='pending', status='pending')
    db.session.add(oi)
    db.session.flush()
    return oi


def test_order_out_subtracts_once(app):
    record_transaction('P-001', 'adjust', 10, 'seed:P-001', reason='期首')
    o = _make_order('order-1')
    oi = _make_item(o, qty=2)
    db.session.commit()

    assert apply_order_out([oi.id]) == 1
    db.session.commit()
    assert get_balance('P-001') == 8

    # 再実行しても二重に引かれない
    assert apply_order_out([oi.id]) == 0
    db.session.commit()
    assert get_balance('P-001') == 8


def test_cancelled_new_order_is_skipped(app):
    o = _make_order('order-2', order_status='4')  # 取込時点で既にキャンセル
    oi = _make_item(o)
    db.session.commit()
    assert apply_order_out([oi.id]) == 0
    assert get_balance('P-001') == 0


def test_cancel_return_restores_once(app):
    record_transaction('P-001', 'adjust', 10, 'seed:P-001', reason='期首')
    o = _make_order('order-3')
    oi = _make_item(o, qty=3)
    db.session.commit()
    apply_order_out([oi.id])
    db.session.commit()
    assert get_balance('P-001') == 7

    o.yahoo_order_status = '4'  # キャンセルに変化
    db.session.commit()
    r = sync_cancel_returns()
    db.session.commit()
    assert r['returned'] == 1
    assert get_balance('P-001') == 10

    r2 = sync_cancel_returns()  # 再実行しても戻しは1回だけ
    db.session.commit()
    assert r2['returned'] == 0
    assert get_balance('P-001') == 10


def test_shipped_cancel_alerts_instead_of_return(app):
    record_transaction('P-001', 'adjust', 10, 'seed:P-001', reason='期首')
    o = _make_order('order-4')
    oi = _make_item(o, qty=1)
    db.session.commit()
    apply_order_out([oi.id])
    db.session.commit()

    o.yahoo_order_status = '4'
    o.yahoo_ship_status = '3'  # 出荷済み
    db.session.commit()
    r = sync_cancel_returns()
    db.session.commit()
    assert r['returned'] == 0
    assert r['alerted'] == 1
    assert get_balance('P-001') == 9  # 戻っていない
    assert Alert.query.filter_by(alert_type='shipped_cancel').count() == 1

    sync_cancel_returns()  # アラートも重複しない
    db.session.commit()
    assert Alert.query.filter_by(alert_type='shipped_cancel').count() == 1


def test_zero_quantity_item_is_skipped(app):
    record_transaction('P-001', 'adjust', 10, 'seed:P-001', reason='期首')
    o = _make_order('order-6')
    oi = _make_item(o, qty=0)
    db.session.commit()
    assert apply_order_out([oi.id]) == 0
    db.session.commit()
    assert get_balance('P-001') == 10


def test_cancel_without_out_does_nothing(app):
    # 台帳導入前の古いキャンセル注文（out記録なし）には何もしない
    o = _make_order('order-5', order_status='4')
    _make_item(o)
    db.session.commit()
    r = sync_cancel_returns()
    db.session.commit()
    assert r == {'returned': 0, 'alerted': 0}
    assert StockTransaction.query.count() == 0
