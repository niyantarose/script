from flask import Flask
from datetime import datetime
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
        return {'now': datetime.now().strftime('%Y/%m/%d')}

    # DB初期化
    with app.app_context():
        db.create_all()

    # 1時間ごとの自動取込スケジューラ
    if _APScheduler:
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

    return app


app = create_app()

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
