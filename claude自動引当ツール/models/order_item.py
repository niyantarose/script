from datetime import datetime
from models import db


class OrderItem(db.Model):
    __tablename__ = 'order_items'

    id = db.Column(db.Integer, primary_key=True)
    order_id = db.Column(db.Integer, db.ForeignKey('orders.id'), nullable=False)
    product_code = db.Column(db.String(100), nullable=False)
    product_sub_code = db.Column(db.String(100), nullable=True)
    product_name = db.Column(db.String(200), nullable=True)
    quantity = db.Column(db.Integer, nullable=False)
    inventory_type = db.Column(db.String(20), nullable=False)
    status = db.Column(db.String(30), default='pending')
    allocated_qty = db.Column(db.Integer, default=0)
    shipped_flag = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.now)
    updated_at = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)

    purchases = db.relationship('Purchase', backref='order_item', lazy='dynamic')
    allocations = db.relationship('Allocation', backref='order_item', lazy='dynamic')

    STATUS_LABELS = {
        'pending': '引当待ち',
        'provisional_allocated': '仮引当済',
        'allocated_sokunou': '即納引当済',
        'partial_waiting': '部分引当中（EMS待ち）',
        'priority_hold': '先送り待機',
        'shortage': '在庫不足',
        'fully_allocated': '引当完了',
        'shipped': '発送完了',
    }

    @property
    def status_label(self):
        return self.STATUS_LABELS.get(self.status, self.status)

    def to_dict(self):
        return {
            'id': self.id,
            'order_id': self.order_id,
            'product_code': self.product_code,
            'product_sub_code': self.product_sub_code or '',
            'product_name': self.product_name or '',
            'quantity': self.quantity,
            'inventory_type': self.inventory_type,
            'status': self.status,
            'status_label': self.status_label,
            'allocated_qty': self.allocated_qty,
            'shipped_flag': self.shipped_flag,
        }
