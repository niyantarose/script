from datetime import datetime
from models import db


class Ems(db.Model):
    __tablename__ = 'ems'

    id = db.Column(db.Integer, primary_key=True)
    ems_number = db.Column(db.String(50), nullable=False, unique=True)
    shipped_at = db.Column(db.Date, nullable=False)
    estimated_arrival = db.Column(db.Date, nullable=False)
    arrived_at = db.Column(db.Date, nullable=True)
    status = db.Column(db.String(30), default='in_transit')
    memo = db.Column(db.Text, nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.now)
    updated_at = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)

    items = db.relationship('EmsItem', backref='ems', lazy='dynamic')

    STATUS_LABELS = {
        'in_transit': '輸送中',
        'arrived': '入荷済',
    }

    @property
    def status_label(self):
        return self.STATUS_LABELS.get(self.status, self.status)

    def to_dict(self):
        return {
            'id': self.id,
            'ems_number': self.ems_number,
            'shipped_at': self.shipped_at.strftime('%Y/%m/%d') if self.shipped_at else '',
            'estimated_arrival': self.estimated_arrival.strftime('%Y/%m/%d') if self.estimated_arrival else '',
            'arrived_at': self.arrived_at.strftime('%Y/%m/%d') if self.arrived_at else '',
            'status': self.status,
            'status_label': self.status_label,
            'memo': self.memo or '',
        }
