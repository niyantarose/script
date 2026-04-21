from datetime import datetime
from models import db


class ImportLog(db.Model):
    __tablename__ = 'import_logs'

    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(200), nullable=False, unique=True)
    file_type = db.Column(db.String(50), nullable=False)  # 'purchase', 'ems'
    imported_at = db.Column(db.DateTime, default=datetime.now)
    record_count = db.Column(db.Integer, default=0)
