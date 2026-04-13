from datetime import datetime
from models import db


class Purchase(db.Model):
    __tablename__ = 'purchases'

    id = db.Column(db.Integer, primary_key=True)
    order_item_id = db.Column(db.Integer, db.ForeignKey('order_items.id'), nullable=False)
    product_code = db.Column(db.String(100), nullable=False)
    product_sub_code = db.Column(db.String(100), nullable=True)
    product_name = db.Column(db.String(200), nullable=True)
    quantity = db.Column(db.Integer, nullable=False)
    shop_name = db.Column(db.String(100), nullable=True)
    ordered_at = db.Column(db.Date, nullable=False)
    status = db.Column(db.String(30), default='ordered')
    memo = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.now)
    updated_at = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)

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
