from datetime import datetime
from models import db


class StockTransaction(db.Model):
    """在庫トランザクション台帳（追記型）。

    現在庫 = この台帳の qty 合計。行の書き換え・削除は禁止で、
    訂正は逆方向の adjust 行を追加する。
    """
    __tablename__ = 'stock_transactions'

    id               = db.Column(db.Integer, primary_key=True)
    product_code     = db.Column(db.String(100), nullable=False, index=True)
    product_sub_code = db.Column(db.String(100), nullable=True)
    tx_type          = db.Column(db.String(30), nullable=False)
    qty              = db.Column(db.Integer, nullable=False)  # 符号付き（＋入庫/−出庫）
    ref_type         = db.Column(db.String(30), nullable=True)   # order_item / japan_staging / stocktake / manual
    ref_id           = db.Column(db.Integer, nullable=True)
    # 発生元の一意キー。UNIQUE制約が二重計上を構造的に防ぐ
    source_key       = db.Column(db.String(200), nullable=False, unique=True)
    reason           = db.Column(db.Text, nullable=True)
    created_at       = db.Column(db.DateTime, default=datetime.now)

    TX_TYPE_LABELS = {
        'receive':       '入庫（仕入れ到着）',
        'order_out':     '注文出庫',
        'cancel_return': 'キャンセル戻し',
        'manual_in':     '手動入庫',
        'manual_out':    '手動出庫',
        'adjust':        '調整（棚卸・訂正）',
    }

    @property
    def tx_type_label(self):
        return self.TX_TYPE_LABELS.get(self.tx_type, self.tx_type)

    def to_dict(self):
        return {
            'id': self.id,
            'product_code': self.product_code,
            'product_sub_code': self.product_sub_code or '',
            'tx_type': self.tx_type,
            'tx_type_label': self.tx_type_label,
            'qty': self.qty,
            'ref_type': self.ref_type or '',
            'ref_id': self.ref_id,
            'source_key': self.source_key,
            'reason': self.reason or '',
            'created_at': self.created_at.strftime('%Y/%m/%d %H:%M') if self.created_at else '',
        }
