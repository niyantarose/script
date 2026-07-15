from models import db, Inventory, StockTransaction, MallSku
from scripts.seed_ledger_opening import seed_opening_balances, seed_yahoo_mall_skus
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
