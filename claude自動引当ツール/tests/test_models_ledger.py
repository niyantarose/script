import pytest
from sqlalchemy.exc import IntegrityError

from models import db, StockTransaction, MallSku


def test_stock_transaction_insert(app):
    tx = StockTransaction(product_code='P-001', tx_type='manual_in',
                          qty=3, source_key='manual:abc', reason='テスト入庫')
    db.session.add(tx)
    db.session.commit()
    assert StockTransaction.query.count() == 1


def test_source_key_unique_blocks_duplicate(app):
    db.session.add(StockTransaction(product_code='P-001', tx_type='order_out',
                                    qty=-1, source_key='yahoo:o1:1:out'))
    db.session.commit()
    db.session.add(StockTransaction(product_code='P-001', tx_type='order_out',
                                    qty=-1, source_key='yahoo:o1:1:out'))
    with pytest.raises(IntegrityError):
        db.session.commit()
    db.session.rollback()
    assert StockTransaction.query.count() == 1


def test_mall_sku_unique_per_mall(app):
    db.session.add(MallSku(mall='yahoo', external_code='EXT-1', product_code='P-001'))
    db.session.commit()
    db.session.add(MallSku(mall='yahoo', external_code='EXT-1', product_code='P-002'))
    with pytest.raises(IntegrityError):
        db.session.commit()
    db.session.rollback()
    # 別モールなら同じ external_code でも登録できる
    db.session.add(MallSku(mall='amazon', external_code='EXT-1', product_code='P-001'))
    db.session.commit()
    assert MallSku.query.count() == 2
