from datetime import datetime
from models import db


class Alert(db.Model):
    __tablename__ = 'alerts'

    id = db.Column(db.Integer, primary_key=True)
    alert_type = db.Column(db.String(50), nullable=False)
    order_id = db.Column(db.Integer, db.ForeignKey('orders.id'), nullable=True)
    order_item_id = db.Column(db.Integer, db.ForeignKey('order_items.id'), nullable=True)
    product_code = db.Column(db.String(100), nullable=True)
    message = db.Column(db.Text, nullable=False)
    resolved_flag = db.Column(db.Boolean, default=False)
    resolved_at = db.Column(db.DateTime, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.now)

    order = db.relationship('Order', backref='alerts')
    order_item = db.relationship('OrderItem', backref='alerts')

    ALERT_TYPE_LABELS = {
        'purchase_missing': '発注漏れ',
        'korea_ship_missing': '韓国発送漏れ',
        'japan_arrival_missing': '日本入荷漏れ',
        'japan_ship_missing': '発送漏れ',
        'stock_shortage': '在庫不足',
        'delay_warning': '遅延警告',
    }

    @property
    def alert_type_label(self):
        return self.ALERT_TYPE_LABELS.get(self.alert_type, self.alert_type)

    def to_dict(self):
        return {
            'id': self.id,
            'alert_type': self.alert_type,
            'alert_type_label': self.alert_type_label,
            'order_id': self.order_id,
            'order_item_id': self.order_item_id,
            'product_code': self.product_code or '',
            'message': self.message,
            'resolved_flag': self.resolved_flag,
            'created_at': self.created_at.strftime('%Y/%m/%d %H:%M') if self.created_at else '',
        }
