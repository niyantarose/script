from datetime import datetime
from models import db


class EmsItem(db.Model):
    __tablename__ = 'ems_items'

    id           = db.Column(db.Integer, primary_key=True)
    ems_id       = db.Column(db.Integer, db.ForeignKey('ems.id'), nullable=False, index=True)
    # OrderItemへの紐付け（nullable: 受注なしでも保存可能）
    order_item_id = db.Column(db.Integer, db.ForeignKey('order_items.id'), nullable=True, index=True)
    purchase_date = db.Column(db.String(50), nullable=True)   # 구매日 作成日（Wata形式 B列）
    purchase_no   = db.Column(db.String(100), nullable=True)  # 구매番号（Wata+_NN形式 F列）
    product_code = db.Column(db.String(100), nullable=False)
    product_sub_code = db.Column(db.String(100), nullable=True)
    product_name = db.Column(db.String(200), nullable=True)
    quantity     = db.Column(db.Integer, nullable=False, default=1)
    created_at   = db.Column(db.DateTime, default=datetime.now)

    order_item = db.relationship('OrderItem', backref='ems_items')

    def to_dict(self):
        return {
            'id':              self.id,
            'ems_id':          self.ems_id,
            'order_item_id':   self.order_item_id,
            'product_code':    self.product_code,
            'product_sub_code':self.product_sub_code or '',
            'product_name':    self.product_name or '',
            'quantity':        self.quantity,
        }
