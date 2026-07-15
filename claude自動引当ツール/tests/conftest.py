import os
import sys

import pytest
from flask import Flask

# プロジェクトルートを import パスに追加
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from models import db as _db  # noqa: E402


@pytest.fixture()
def app():
    """in-memory SQLite で全テーブルを作った素の Flask アプリ。

    app.py の create_app() は .env や APScheduler に依存するため使わず、
    モデル層のテストに必要な最小構成だけ組む。
    """
    flask_app = Flask(
        __name__,
        template_folder=os.path.join(ROOT, 'templates'),
        static_folder=os.path.join(ROOT, 'static'),
    )
    flask_app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///:memory:'
    flask_app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
    flask_app.config['TESTING'] = True
    _db.init_app(flask_app)
    with flask_app.app_context():
        _db.create_all()
        yield flask_app
        _db.session.rollback()
        _db.drop_all()


@pytest.fixture()
def client(app):
    return app.test_client()
