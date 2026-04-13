from datetime import datetime
from models import db


class Inventory(db.Model):
    __tablename__ = 'inventory'

    id = db.Column(db.Integer, primary_key=True)
    product_code = db.Column(db.String(100), nullable=False)
    product_sub_code = db.Column(db.String(100), nullable=True)
    inventory_type = db.Column(db.String(20), nullable=False)
    quantity = db.Column(db.Integer, default=0)
    reserved_qty = db.Column(db.Integer, default=0)
    available_qty = db.Column(db.Integer, default=0)
    updated_at = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)

    def to_dict(self):
        return {
            'id': self.id,
            'product_code': self.product_code,
            'product_sub_code': self.product_sub_code or '',
            'inventory_type': self.inventory_type,
            'quantity': self.quantity,
            'reserved_qty': self.reserved_qty,
            'available_qty': self.available_qty,
        }
