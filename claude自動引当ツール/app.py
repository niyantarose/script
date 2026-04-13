from flask import Flask
from datetime import datetime
from config import Config
from models import db


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

    app.register_blueprint(dashboard_bp)
    app.register_blueprint(orders_bp)
    app.register_blueprint(purchases_bp)
    app.register_blueprint(ems_bp)
    app.register_blueprint(japan_bp)

    # テンプレートにnowを渡す
    @app.context_processor
    def inject_now():
        return {'now': datetime.now().strftime('%Y/%m/%d')}

    # DB初期化
    with app.app_context():
        db.create_all()

    return app


app = create_app()

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
