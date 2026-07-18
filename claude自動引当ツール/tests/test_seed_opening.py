from models import db, Inventory, StockTransaction, MallSku
from scripts.seed_ledger_opening import (
    seed_opening_balances, seed_yahoo_mall_skus, backfill_order_out,
)
from services.stock_ledger import get_balance


def _setup_rows():
    db.session.add(Inventory(product_code='P-A', inventory_type='即納',
                             quantity=7, reserved_qty=0, available_qty=7))
    db.session.add(Inventory(product_code='P-B', inventory_type='即納',
                             quantity=0, reserved_qty=0, available_qty=0))
    db.session.add(Inventory(product_code='P-A', inventory_type='yahoo',
                             quantity=7, yahoo_stock=7))
    db.session.commit()


def test_seed_opening_is_idempotent(app):
    _setup_rows()
    r1 = seed_opening_balances()
    db.session.commit()
    assert r1['seeded'] == 1          # P-A のみ（P-B は 0 なのでスキップ）
    assert get_balance('P-A') == 7

    r2 = seed_opening_balances()      # 2回目は何も増えない
    db.session.commit()
    assert r2['seeded'] == 0
    assert get_balance('P-A') == 7
    assert StockTransaction.query.count() == 1


def test_seed_yahoo_mall_skus(app):
    _setup_rows()
    r1 = seed_yahoo_mall_skus()
    db.session.commit()
    assert r1['created'] == 1
    ms = MallSku.query.filter_by(mall='yahoo', external_code='P-A').first()
    assert ms.product_code == 'P-A'

    r2 = seed_yahoo_mall_skus()
    db.session.commit()
    assert r2['created'] == 0
    assert MallSku.query.count() == 1


def test_backfill_order_out_only_unshipped_active(app):
    from datetime import datetime
    from models import Order, OrderItem
    from scripts.seed_ledger_opening import backfill_order_out
    db.session.add(Inventory(product_code='P-A', inventory_type='即納',
                             quantity=7, reserved_qty=0, available_qty=7))
    o1 = Order(yahoo_order_id='bf-1', ordered_at=datetime.now(),
               yahoo_order_status='2', yahoo_ship_status='0')   # 未出荷・有効
    o2 = Order(yahoo_order_id='bf-2', ordered_at=datetime.now(),
               yahoo_order_status='2', yahoo_ship_status='3')   # 出荷済み
    o3 = Order(yahoo_order_id='bf-3', ordered_at=datetime.now(),
               yahoo_order_status='4', yahoo_ship_status='0')   # キャンセル
    db.session.add_all([o1, o2, o3])
    db.session.flush()
    for o, qty in ((o1, 2), (o2, 3), (o3, 1)):
        db.session.add(OrderItem(order_id=o.id, product_code='P-A', quantity=qty,
                                 inventory_type='pending', status='pending'))
    db.session.commit()

    from scripts.seed_ledger_opening import seed_opening_balances
    seed_opening_balances()
    r = backfill_order_out()
    db.session.commit()
    assert r['recorded'] == 1           # o1 の明細のみ
    assert get_balance('P-A') == 5      # 7 - 2

    r2 = backfill_order_out()           # 冪等
    db.session.commit()
    assert r2['recorded'] == 0
    assert get_balance('P-A') == 5
