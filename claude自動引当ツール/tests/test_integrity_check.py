from models import db, Alert, Inventory
from services.stock_ledger import record_transaction, verify_cache_integrity


def test_mismatch_is_fixed_and_alerted(app):
    record_transaction('P-001', 'manual_in', 5, 'manual:i1')
    db.session.commit()

    # 台帳を通さない直接書き換え（あってはならない操作）を再現
    inv = Inventory.query.filter_by(product_code='P-001', inventory_type='即納').first()
    inv.quantity = 99
    db.session.commit()

    mismatches = verify_cache_integrity()
    db.session.commit()
    assert mismatches == [{'product_code': 'P-001', 'expected': 5, 'actual': 99}]
    assert inv.quantity == 5  # 修正済み
    assert Alert.query.filter_by(alert_type='ledger_mismatch').count() == 1

    # 再実行: 一致しているので何も起きない
    assert verify_cache_integrity() == []
    assert Alert.query.filter_by(alert_type='ledger_mismatch').count() == 1
