import pytest

from models import db, Inventory, StockTransaction
from services.stock_ledger import record_transaction, get_balance, recalc_inventory


def test_record_and_balance(app):
    tx, created = record_transaction('P-001', 'manual_in', 5, 'manual:t1', reason='入庫')
    db.session.commit()
    assert created is True
    assert get_balance('P-001') == 5

    record_transaction('P-001', 'manual_out', -2, 'manual:t2', reason='破損')
    db.session.commit()
    assert get_balance('P-001') == 3


def test_same_source_key_records_once(app):
    record_transaction('P-001', 'order_out', -1, 'yahoo:o1:1:out')
    db.session.commit()
    tx, created = record_transaction('P-001', 'order_out', -1, 'yahoo:o1:1:out')
    db.session.commit()
    assert created is False
    assert StockTransaction.query.count() == 1
    assert get_balance('P-001') == -1


def test_sign_validation(app):
    with pytest.raises(ValueError):
        record_transaction('P-001', 'manual_in', -3, 'manual:bad1')  # 入庫は＋のみ
    with pytest.raises(ValueError):
        record_transaction('P-001', 'order_out', 1, 'manual:bad2')   # 出庫は−のみ
    with pytest.raises(ValueError):
        record_transaction('P-001', 'adjust', 0, 'manual:bad3')      # 0は禁止


def test_recalc_updates_sokunou_inventory(app):
    inv = Inventory(product_code='P-002', inventory_type='即納',
                    quantity=10, reserved_qty=4, available_qty=6)
    db.session.add(inv)
    db.session.commit()

    record_transaction('P-002', 'adjust', 10, 'seed:P-002', reason='期首残高')
    record_transaction('P-002', 'order_out', -3, 'yahoo:o2:9:out')
    db.session.commit()

    got = Inventory.query.filter_by(product_code='P-002', inventory_type='即納').first()
    assert got.quantity == 7            # 10 - 3
    assert got.available_qty == 3       # 7 - reserved 4


def test_recalc_creates_missing_inventory_row(app):
    record_transaction('P-NEW', 'manual_in', 2, 'manual:t3')
    db.session.commit()
    got = Inventory.query.filter_by(product_code='P-NEW', inventory_type='即納').first()
    assert got is not None
    assert got.quantity == 2
