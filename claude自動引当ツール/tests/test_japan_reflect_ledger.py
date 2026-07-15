import pytest

from models import db, Inventory, JapanInventoryStaging, StockTransaction
from services.stock_ledger import get_balance


@pytest.fixture()
def jclient(app):
    from routes.japan_inventory import bp
    app.register_blueprint(bp)
    return app.test_client()


def _make_staging(product_code='P-001', qty=4):
    # SQLite は FK 未強制なので ems_item_id はダミーIDで良い
    jis = JapanInventoryStaging(ems_item_id=999, product_code=product_code,
                                quantity=qty, status='to_japan_stock')
    db.session.add(jis)
    db.session.commit()
    return jis


def test_reflect_records_receive_tx(app, jclient):
    jis = _make_staging(qty=4)
    res = jclient.post('/japan/reflect')
    assert res.status_code == 200
    assert res.get_json()['reflected_count'] == 1

    assert get_balance('P-001') == 4
    inv = Inventory.query.filter_by(product_code='P-001', inventory_type='即納').first()
    assert inv.quantity == 4

    tx = StockTransaction.query.filter_by(source_key=f'japan_staging:{jis.id}').first()
    assert tx is not None and tx.tx_type == 'receive'


def test_reflect_is_idempotent(app, jclient):
    jis = _make_staging(qty=4)
    jclient.post('/japan/reflect')
    # ステータスを強制的に戻して再実行しても、台帳は二重計上しない
    jis.status = 'to_japan_stock'
    db.session.commit()
    jclient.post('/japan/reflect')
    assert get_balance('P-001') == 4
