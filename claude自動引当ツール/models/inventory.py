from datetime import datetime
from models import db


class Inventory(db.Model):
    __tablename__ = 'inventory'

    id               = db.Column(db.Integer, primary_key=True)
    product_code     = db.Column(db.String(100), nullable=False, index=True)
    product_sub_code = db.Column(db.String(100), nullable=True)
    product_name     = db.Column(db.String(200), nullable=True)   # 商品名
    inventory_type   = db.Column(db.String(20),  nullable=False)  # 'yahoo' / 'local'
    quantity         = db.Column(db.Integer, default=0)           # 在庫数（Yahoo同期値）
    reserved_qty     = db.Column(db.Integer, default=0)           # 引当済み数
    available_qty    = db.Column(db.Integer, default=0)           # 引当可能数
    yahoo_stock      = db.Column(db.Integer, default=0)           # Yahoo上の在庫数（生値）
    price            = db.Column(db.Integer, default=0)           # 販売価格
    # ── ロケーション管理（即納在庫） ──────────────────────
    location         = db.Column(db.String(100), nullable=True)   # 棚番号（例: A-3-左）
    is_immediate     = db.Column(db.Boolean, default=False)       # 即納在庫フラグ
    # ── タイムスタンプ ──────────────────────────────────
    last_synced_at   = db.Column(db.DateTime, nullable=True)      # Yahoo最終同期日時
    updated_at       = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)

    def to_dict(self):
        return {
            'id':              self.id,
            'product_code':    self.product_code,
            'product_sub_code':self.product_sub_code or '',
            'product_name':    self.product_name or '',
            'inventory_type':  self.inventory_type,
            'quantity':        self.quantity,
            'reserved_qty':    self.reserved_qty,
            'available_qty':   self.available_qty,
            'yahoo_stock':     self.yahoo_stock,
            'price':           self.price,
            'location':        self.location or '',
            'is_immediate':    self.is_immediate,
            'last_synced_at':  self.last_synced_at.strftime('%Y/%m/%d %H:%M') if self.last_synced_at else '',
        }
