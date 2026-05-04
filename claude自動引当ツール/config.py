import os
from urllib.parse import quote_plus

from dotenv import load_dotenv

load_dotenv()


def _mysql_uri() -> str:
    """明示的に DB_ENGINE=mysql のときのみ使用（後方互換）。"""
    db_user = os.getenv('DB_USER', 'inventory_user')
    db_pass = os.getenv('DB_PASSWORD', '')
    db_host = os.getenv('DB_HOST', 'localhost')
    db_name = os.getenv('DB_NAME', 'inventory_db')
    user_q = quote_plus(db_user)
    pass_q = quote_plus(db_pass) if db_pass else ''
    auth = f'{user_q}:{pass_q}@' if pass_q else f'{user_q}@'
    return f'mysql+pymysql://{auth}{db_host}/{db_name}?charset=utf8mb4'


def _postgres_uri_from_parts() -> str:
    """DATABASE_URL 未設定時に PostgreSQL URI を組み立てる。"""
    db_user = os.getenv('DB_USER', 'postgres')
    db_pass = os.getenv('DB_PASSWORD', '')
    db_host = os.getenv('DB_HOST', '127.0.0.1')
    db_port = os.getenv('DB_PORT', '5432')
    db_name = os.getenv('DB_NAME', 'postgres')
    user_q = quote_plus(db_user)
    name_q = quote_plus(db_name)
    if db_pass:
        return (
            f'postgresql+psycopg://{user_q}:{quote_plus(db_pass)}'
            f'@{db_host}:{db_port}/{name_q}'
        )
    return f'postgresql+psycopg://{user_q}@{db_host}:{db_port}/{name_q}'


class Config:
    SECRET_KEY = os.getenv('SECRET_KEY', 'dev-secret-key')
    APSCHEDULER_ENABLED = (
        os.getenv('ENABLE_APSCHEDULER', os.getenv('APSCHEDULER_ENABLED', 'false')).lower() == 'true'
    )

    if os.getenv('USE_SQLITE', 'true').lower() == 'true':
        SQLALCHEMY_DATABASE_URI = 'sqlite:///inventory.db'
    elif os.getenv('DB_ENGINE', '').lower() == 'mysql':
        SQLALCHEMY_DATABASE_URI = _mysql_uri()
    else:
        db_url = (os.getenv('DATABASE_URL') or '').strip()
        if db_url:
            SQLALCHEMY_DATABASE_URI = db_url
        else:
            SQLALCHEMY_DATABASE_URI = _postgres_uri_from_parts()

    SQLALCHEMY_TRACK_MODIFICATIONS = False
    TEMPLATES_AUTO_RELOAD = True
