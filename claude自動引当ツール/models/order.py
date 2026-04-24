from datetime import datetime
from models import db


class Order(db.Model):
    __tablename__ = 'orders'

    id = db.Column(db.Integer, primary_key=True)
    yahoo_order_id = db.Column(db.String(50), nullable=False, unique=True)
    ordered_at = db.Column(db.DateTime, nullable=False)
    desired_delivery_date = db.Column(db.Date, nullable=True)
    customer_name = db.Column(db.String(100), nullable=True)
    priority_ship_flag = db.Column(db.Boolean, default=False)
    yahoo_ship_status  = db.Column(db.String(30), nullable=True)
    yahoo_order_status = db.Column(db.String(5),  nullable=True)
    status = db.Column(db.String(30), default='pending')
    delay_memo = db.Column(db.Text, nullable=True)
    customer_contacted_flag = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.now)
    updated_at = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)

    items = db.relationship('OrderItem', backref='order', lazy='dynamic')

    @property
    def days_elapsed(self):
        return (datetime.now() - self.ordered_at).days

    def to_dict(self):
        return {
            'id': self.id,
            'yahoo_order_id': self.yahoo_order_id,
            'ordered_at': self.ordered_at.strftime('%Y/%m/%d %H:%M'),
            'desired_delivery_date': self.desired_delivery_date.strftime('%Y/%m/%d') if self.desired_delivery_date else '',
            'customer_name': self.customer_name or '',
            'priority_ship_flag': self.priority_ship_flag,
            'yahoo_ship_status': self.yahoo_ship_status or '',
            'status': self.status,
            'delay_memo': self.delay_memo or '',
            'customer_contacted_flag': self.customer_contacted_flag,
            'days_elapsed': self.days_elapsed,
        }
