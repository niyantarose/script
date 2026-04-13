from datetime import datetime
from models import db


class EmsItem(db.Model):
    __tablename__ = 'ems_items'

    id = db.Column(db.Integer, primary_key=True)
    ems_id = db.Column(db.Integer, db.ForeignKey('ems.id'), nullable=False)
    order_item_id = db.Column(db.Integer, db.ForeignKey('order_items.id'), nullable=False)
    product_code = db.Column(db.String(100), nullable=False)
    product_sub_code = db.Column(db.String(100), nullable=True)
    quantity = db.Column(db.Integer, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.now)

    order_item = db.relationship('OrderItem', backref='ems_items')

    def to_dict(self):
        return {
            'id': self.id,
            'ems_id': self.ems_id,
            'order_item_id': self.order_item_id,
            'product_code': self.product_code,
            'product_sub_code': self.product_sub_code or '',
            'quantity': self.quantity,
        }
