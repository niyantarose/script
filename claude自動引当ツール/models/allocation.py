from datetime import datetime
from models import db


class Allocation(db.Model):
    __tablename__ = 'allocations'

    id = db.Column(db.Integer, primary_key=True)
    order_item_id = db.Column(db.Integer, db.ForeignKey('order_items.id'), nullable=False)
    inventory_type = db.Column(db.String(20), nullable=False)
    allocation_type = db.Column(db.String(20), nullable=False)
    quantity = db.Column(db.Integer, nullable=False)
    ems_item_id = db.Column(db.Integer, db.ForeignKey('ems_items.id'), nullable=True)
    allocated_by = db.Column(db.String(50), nullable=True)
    allocated_at = db.Column(db.DateTime, default=datetime.now)

    ems_item = db.relationship('EmsItem', backref='allocations')

    ALLOCATION_TYPE_LABELS = {
        '仮引当': '仮引当',
        '本引当': '本引当',
        '手動': '手動引当',
    }

    def to_dict(self):
        return {
            'id': self.id,
            'order_item_id': self.order_item_id,
            'inventory_type': self.inventory_type,
            'allocation_type': self.allocation_type,
            'quantity': self.quantity,
            'ems_item_id': self.ems_item_id,
            'allocated_by': self.allocated_by or '',
            'allocated_at': self.allocated_at.strftime('%Y/%m/%d %H:%M') if self.allocated_at else '',
        }
