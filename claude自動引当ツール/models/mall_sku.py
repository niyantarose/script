from datetime import datetime
from models import db


class MallSku(db.Model):
    """モール側商品コード → 内部商品コードのマッピング（Phase 4 の土台）。"""
    __tablename__ = 'mall_skus'
    __table_args__ = (
        db.UniqueConstraint('mall', 'external_code', name='uq_mall_external'),
    )

    id                = db.Column(db.Integer, primary_key=True)
    mall              = db.Column(db.String(20), nullable=False)  # yahoo/amazon/qoo10/mercari/tiktok
    external_code     = db.Column(db.String(100), nullable=False)
    external_sub_code = db.Column(db.String(100), nullable=True)
    product_code      = db.Column(db.String(100), nullable=False, index=True)
    product_sub_code  = db.Column(db.String(100), nullable=True)
    created_at        = db.Column(db.DateTime, default=datetime.now)
