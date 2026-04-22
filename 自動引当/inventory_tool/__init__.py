from flask import Flask
from sqlalchemy import inspect, text

from .commands import register_commands
from .config import Config
from .extensions import db
from .views import main_bp, register_template_helpers


def ensure_schema_compatibility() -> None:
    inspector = inspect(db.engine)
    table_names = set(inspector.get_table_names())
    if "orders" not in table_names:
        return

    order_columns = {column["name"] for column in inspector.get_columns("orders")}
    if "customer_code" not in order_columns:
        with db.engine.begin() as connection:
            connection.execute(text("ALTER TABLE orders ADD COLUMN customer_code VARCHAR(100)"))

    purchase_columns = {column["name"] for column in inspector.get_columns("purchases")} if "purchases" in table_names else set()
    if "purchases" in table_names and "source_type" not in purchase_columns:
        with db.engine.begin() as connection:
            connection.execute(text("ALTER TABLE purchases ADD COLUMN source_type VARCHAR(20) DEFAULT 'daniel' NOT NULL"))

    ems_columns = {column["name"] for column in inspector.get_columns("ems")} if "ems" in table_names else set()
    if "ems" in table_names and "source_type" not in ems_columns:
        with db.engine.begin() as connection:
            connection.execute(text("ALTER TABLE ems ADD COLUMN source_type VARCHAR(20) DEFAULT 'daniel' NOT NULL"))


def create_app(config_object: type[Config] = Config) -> Flask:
    app = Flask(__name__, template_folder="../templates", static_folder="../static")
    app.config.from_object(config_object)

    db.init_app(app)
    app.register_blueprint(main_bp)
    register_template_helpers(app)
    register_commands(app)

    with app.app_context():
        db.create_all()
        ensure_schema_compatibility()

    return app
