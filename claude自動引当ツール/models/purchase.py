from datetime import datetime
from models import db


class Purchase(db.Model):
    __tablename__ = 'purchases'

    id = db.Column(db.Integer, primary_key=True)
    # 発注NOとYahoo受注番号を直接保持（OrderItemが無くても取込可能）
    purchase_no  = db.Column(db.String(100), nullable=True, index=True)  # Wata240801_01
    order_id     = db.Column(db.String(100), nullable=True)              # Yahoo受注番号
    # OrderItemへの紐付け（nullable: 受注と紐付かない発注も保存できる）
    order_item_id = db.Column(db.Integer, db.ForeignKey('order_items.id'), nullable=True)
    product_code = db.Column(db.String(100), nullable=False)
    product_sub_code = db.Column(db.String(100), nullable=True)
    product_name = db.Column(db.String(200), nullable=True)
    quantity     = db.Column(db.Integer, nullable=False, default=1)
    shop_name    = db.Column(db.String(100), nullable=True)
    ordered_at   = db.Column(db.Date, nullable=True)
    status       = db.Column(db.String(30), default='ordered', index=True)
    agent        = db.Column(db.String(20), default='daniel',  index=True)  # 'daniel' / 'tegu'
    memo         = db.Column(db.Text, nullable=True)
    created_at   = db.Column(db.DateTime, default=datetime.now)
    updated_at   = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)

    STATUS_LABELS = {
        'ordered': '発注済',
        'arrived': '入荷済',
    }

    @property
    def status_label(self):
        return self.STATUS_LABELS.get(self.status, self.status)

    def to_dict(self):
        return {
            'id': self.id,
            'purchase_no': self.purchase_no or '',
            'order_id': self.order_id or '',
            'order_item_id': self.order_item_id,
            'product_code': self.product_code,
            'product_sub_code': self.product_sub_code or '',
            'product_name': self.product_name or '',
            'quantity': self.quantity,
            'shop_name': self.shop_name or '',
            'ordered_at': self.ordered_at.strftime('%Y/%m/%d') if self.ordered_at else '',
            'status': self.status,
            'status_label': self.status_label,
            'memo': self.memo or '',
        }
