"""台帳の期首残高と Yahoo モールSKUマッピングをシードする。

使い方（プロジェクトルートで）:
    python scripts/seed_ledger_opening.py

何度実行しても安全（source_key / UNIQUE制約で冪等）。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from models import db, Inventory, MallSku, OrderItem
from services.stock_ledger import record_transaction, apply_order_out


def seed_opening_balances():
    """即納 Inventory の現在値を期首残高として台帳に記録する。"""
    seeded = skipped = 0
    rows = Inventory.query.filter_by(inventory_type='即納').all()
    for inv in rows:
        qty = inv.quantity or 0
        if qty == 0:
            skipped += 1
            continue
        _, created = record_transaction(
            inv.product_code, 'adjust', qty, f'seed:{inv.product_code}',
            product_sub_code=inv.product_sub_code,
            ref_type='manual', reason='期首残高（台帳導入時の初期値）',
        )
        seeded += 1 if created else 0
        skipped += 0 if created else 1
    return {'seeded': seeded, 'skipped': skipped}


def seed_yahoo_mall_skus():
    """Yahoo 在庫行から MallSku を1:1で作成する。"""
    created = skipped = 0
    rows = Inventory.query.filter_by(inventory_type='yahoo').all()
    for inv in rows:
        exists = MallSku.query.filter_by(
            mall='yahoo', external_code=inv.product_code).first()
        if exists:
            skipped += 1
            continue
        db.session.add(MallSku(
            mall='yahoo',
            external_code=inv.product_code,
            external_sub_code=inv.product_sub_code,
            product_code=inv.product_code,
            product_sub_code=inv.product_sub_code,
        ))
        created += 1
    return {'created': created, 'skipped': skipped}


def backfill_order_out():
    """切替時点の既存注文明細の出庫を台帳に記録する。

    apply_order_out がキャンセル済み・出荷済み・数量0以下を内部でスキップするため、
    全明細IDを渡せば「未出荷・未キャンセルの明細」だけが出庫記録される。冪等。
    """
    item_ids = [row[0] for row in db.session.query(OrderItem.id).all()]
    recorded = apply_order_out(item_ids)
    return {'recorded': recorded, 'checked': len(item_ids)}


if __name__ == '__main__':
    from app import app
    with app.app_context():
        r1 = seed_opening_balances()
        r2 = backfill_order_out()
        r3 = seed_yahoo_mall_skus()
        db.session.commit()
        print(f'期首残高: {r1} / 出庫バックフィル: {r2} / YahooモールSKU: {r3}')
