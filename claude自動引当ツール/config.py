import os
from dotenv import load_dotenv

load_dotenv()


class Config:
    SECRET_KEY = os.getenv('SECRET_KEY', 'dev-secret-key')

    if os.getenv('USE_SQLITE', 'true').lower() == 'true':
        SQLALCHEMY_DATABASE_URI = 'sqlite:///inventory.db'
    else:
        db_user = os.getenv('DB_USER', 'inventory_user')
        db_pass = os.getenv('DB_PASSWORD', '')
        db_host = os.getenv('DB_HOST', 'localhost')
        db_name = os.getenv('DB_NAME', 'inventory_db')
        SQLALCHEMY_DATABASE_URI = f'mysql+pymysql://{db_user}:{db_pass}@{db_host}/{db_name}?charset=utf8mb4'

    SQLALCHEMY_TRACK_MODIFICATIONS = False
    TEMPLATES_AUTO_RELOAD = True
