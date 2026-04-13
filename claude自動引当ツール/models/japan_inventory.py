from datetime import datetime
from models import db


class JapanInventoryStaging(db.Model):
    __tablename__ = 'japan_inventory_staging'

    id = db.Column(db.Integer, primary_key=True)
    ems_item_id = db.Column(db.Integer, db.ForeignKey('ems_items.id'), nullable=False)
    product_code = db.Column(db.String(100), nullable=False)
    product_sub_code = db.Column(db.String(100), nullable=True)
    quantity = db.Column(db.Integer, nullable=False)
    status = db.Column(db.String(30), default='waiting')
    assigned_order_item_id = db.Column(db.Integer, db.ForeignKey('order_items.id'), nullable=True)
    reflected_at = db.Column(db.DateTime, nullable=True)
    excluded_reason = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.now)
    updated_at = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)

    ems_item = db.relationship('EmsItem', backref='japan_staging')
    assigned_order_item = db.relationship('OrderItem', backref='japan_staging_assigned')

    STATUS_LABELS = {
        'waiting': '仕分け待ち',
        'assigned_to_order': '受注割当済',
        'to_japan_stock': '日本在庫へ',
        'excluded': '除外',
        'returned_to_ems': 'EMS差戻',
        'reflected': '反映完了',
    }

    @property
    def status_label(self):
        return self.STATUS_LABELS.get(self.status, self.status)

    def to_dict(self):
        return {
            'id': self.id,
            'ems_item_id': self.ems_item_id,
            'product_code': self.product_code,
            'product_sub_code': self.product_sub_code or '',
            'quantity': self.quantity,
            'status': self.status,
            'status_label': self.status_label,
            'assigned_order_item_id': self.assigned_order_item_id,
            'reflected_at': self.reflected_at.strftime('%Y/%m/%d %H:%M') if self.reflected_at else '',
            'excluded_reason': self.excluded_reason or '',
        }
