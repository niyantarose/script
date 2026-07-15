import pytest

from models import db, Inventory
from services.stock_ledger import get_balance, record_transaction


@pytest.fixture()
def lclient(app):
    from routes.ledger import bp
    app.register_blueprint(bp)
    return app.test_client()


def test_manual_in_and_out(app, lclient):
    res = lclient.post('/ledger/api/tx', json={
        'product_code': 'P-001', 'tx_type': 'manual_in', 'qty': 5,
        'reason': '手動入庫テスト'})
    assert res.status_code == 200
    assert get_balance('P-001') == 5

    res = lclient.post('/ledger/api/tx', json={
        'product_code': 'P-001', 'tx_type': 'manual_out', 'qty': 2,
        'reason': '破損'})
    assert res.status_code == 200
    assert get_balance('P-001') == 3


def test_manual_out_requires_reason(app, lclient):
    res = lclient.post('/ledger/api/tx', json={
        'product_code': 'P-001', 'tx_type': 'manual_out', 'qty': 1, 'reason': ''})
    assert res.status_code == 400
    assert get_balance('P-001') == 0


def test_invalid_qty_rejected(app, lclient):
    res = lclient.post('/ledger/api/tx', json={
        'product_code': 'P-001', 'tx_type': 'manual_in', 'qty': 0})
    assert res.status_code == 400
    res = lclient.post('/ledger/api/tx', json={
        'product_code': 'P-001', 'tx_type': 'manual_in', 'qty': -3})
    assert res.status_code == 400


def test_balances_search(app, lclient):
    db.session.add(Inventory(product_code='ABC-1', product_name='ダンダダン 1巻',
                             inventory_type='即納', quantity=3, reserved_qty=0,
                             available_qty=3, location='A-1'))
    db.session.commit()
    res = lclient.get('/ledger/api/balances?q=ダンダ')
    data = res.get_json()
    assert len(data['items']) == 1
    assert data['items'][0]['product_code'] == 'ABC-1'
    assert data['items'][0]['location'] == 'A-1'


def test_history(app, lclient):
    record_transaction('P-001', 'manual_in', 5, 'manual:h1', reason='入庫1')
    record_transaction('P-001', 'manual_out', -1, 'manual:h2', reason='出庫1')
    db.session.commit()
    res = lclient.get('/ledger/api/history/P-001')
    items = res.get_json()['items']
    assert len(items) == 2
    assert items[0]['reason'] == '出庫1'  # 新しい順
