"""在庫足し引き台帳サービス。

すべての在庫増減はここを通して stock_transactions に記録する。
source_key の UNIQUE 制約により、同じ発生源の足し引きは何度呼んでも1回だけ。
record_transaction / 各同期関数は commit しない（呼び出し側で commit する）。
"""
from sqlalchemy import func

from models import db, Inventory, StockTransaction

# tx_type ごとの符号ルール
_POSITIVE_TYPES = {'receive', 'cancel_return', 'manual_in'}
_NEGATIVE_TYPES = {'order_out', 'manual_out'}
_ANY_SIGN_TYPES = {'adjust'}


def record_transaction(product_code, tx_type, qty, source_key, *,
                       product_sub_code=None, ref_type=None, ref_id=None,
                       reason=None):
    """台帳に1行記録し、即納Inventoryキャッシュを再計算する。

    Returns:
        (tx, created): source_key が既存なら (既存tx, False)
    Raises:
        ValueError: tx_type と qty の符号が矛盾、または qty=0
    """
    if tx_type in _POSITIVE_TYPES and qty <= 0:
        raise ValueError(f'{tx_type} の qty は正の値のみ: {qty}')
    if tx_type in _NEGATIVE_TYPES and qty >= 0:
        raise ValueError(f'{tx_type} の qty は負の値のみ: {qty}')
    if tx_type in _ANY_SIGN_TYPES and qty == 0:
        raise ValueError('adjust の qty に 0 は指定できません')
    if tx_type not in (_POSITIVE_TYPES | _NEGATIVE_TYPES | _ANY_SIGN_TYPES):
        raise ValueError(f'不明な tx_type: {tx_type}')

    existing = StockTransaction.query.filter_by(source_key=source_key).first()
    if existing:
        return existing, False

    tx = StockTransaction(
        product_code=product_code,
        product_sub_code=product_sub_code,
        tx_type=tx_type,
        qty=qty,
        ref_type=ref_type,
        ref_id=ref_id,
        source_key=source_key,
        reason=reason,
    )
    db.session.add(tx)
    db.session.flush()
    recalc_inventory(product_code, product_sub_code)
    return tx, True


def get_balance(product_code):
    """台帳上の現在庫（qty 合計）。"""
    total = db.session.query(func.coalesce(func.sum(StockTransaction.qty), 0)) \
        .filter(StockTransaction.product_code == product_code).scalar()
    return int(total or 0)


def recalc_inventory(product_code, product_sub_code=None):
    """即納 Inventory 行の quantity を台帳合計で上書きする（キャッシュ更新）。"""
    balance = get_balance(product_code)
    inv = Inventory.query.filter_by(
        product_code=product_code, inventory_type='即納').first()
    if not inv:
        inv = Inventory(
            product_code=product_code,
            product_sub_code=product_sub_code,
            inventory_type='即納',
            quantity=0, reserved_qty=0, available_qty=0,
        )
        db.session.add(inv)
    inv.quantity = balance
    inv.available_qty = max(0, balance - (inv.reserved_qty or 0))
    return inv
