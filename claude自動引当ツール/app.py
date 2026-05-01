from flask import Flask
from datetime import datetime
import os
import time
from config import Config
from models import db

try:
    from apscheduler.schedulers.background import BackgroundScheduler
    _APScheduler = BackgroundScheduler
except ImportError:
    _APScheduler = None


def create_app():
    app = Flask(__name__)
    app.config.from_object(Config)

    db.init_app(app)

    # Blueprint登録
    from routes.dashboard import bp as dashboard_bp
    from routes.orders import bp as orders_bp
    from routes.purchases import bp as purchases_bp
    from routes.ems import bp as ems_bp
    from routes.japan_inventory import bp as japan_bp
    from routes.import_data import bp as import_bp
    from routes.order_search import bp as order_search_bp
    from routes.oauth import bp as oauth_bp
    from routes.stock import bp as stock_bp

    app.register_blueprint(dashboard_bp)
    app.register_blueprint(orders_bp)
    app.register_blueprint(purchases_bp)
    app.register_blueprint(ems_bp)
    app.register_blueprint(japan_bp)
    app.register_blueprint(import_bp)
    app.register_blueprint(order_search_bp)
    app.register_blueprint(oauth_bp)
    app.register_blueprint(stock_bp)

    # テンプレートにnowを渡す
    @app.context_processor
    def inject_now():
        now_ts = int(time.time())
        env_path = os.path.join(os.path.dirname(__file__), '.env')
        env_map = {}
        try:
            with open(env_path, 'r', encoding='utf-8') as f:
                for line in f:
                    if '=' in line and not line.lstrip().startswith('#'):
                        key, value = line.strip().split('=', 1)
                        env_map[key] = value
        except Exception:
            env_map = {}

        expires_at = int(env_map.get('YAHOO_REFRESH_EXPECTED_EXPIRES_AT', '0') or 0)
        refresh_days_left = None
        refresh_expires_at_text = ''
        refresh_status = 'unknown'
        if expires_at > 0:
            refresh_days_left = (expires_at - now_ts) / 86400
            refresh_expires_at_text = datetime.fromtimestamp(expires_at).strftime('%Y/%m/%d %H:%M')
            if refresh_days_left < 0:
                refresh_status = 'expired'
            elif refresh_days_left <= 1:
                refresh_status = 'danger'
            elif refresh_days_left <= 3:
                refresh_status = 'warn'
            else:
                refresh_status = 'ok'

        return {
            'now': datetime.now().strftime('%Y/%m/%d'),
            'yahoo_refresh_days_left': refresh_days_left,
            'yahoo_refresh_expires_at': refresh_expires_at_text,
            'yahoo_refresh_status': refresh_status,
        }

    # DB初期化
    with app.app_context():
        db.create_all()

    # Flask内APSchedulerは明示有効化時のみ起動（Gunicorn worker重複を防ぐ）
    if _APScheduler and app.config.get('APSCHEDULER_ENABLED', False):
        from routes.import_data import run_all_imports_job
        scheduler = _APScheduler(daemon=True)
        scheduler.add_job(
            func=run_all_imports_job,
            args=[app],
            trigger='interval',
            hours=1,
            id='auto_import',
            misfire_grace_time=300,
        )
        try:
            scheduler.start()
            app.logger.info('APScheduler started: auto-import every 1 hour')
        except Exception as e:
            app.logger.warning(f'Scheduler start failed: {e}')
    else:
        app.logger.info('APScheduler disabled in Flask app process')

    return app


app = create_app()

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
