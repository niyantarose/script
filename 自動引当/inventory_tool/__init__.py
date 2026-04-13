from flask import Flask

from .commands import register_commands
from .config import Config
from .extensions import db
from .views import main_bp, register_template_helpers


def create_app(config_object: type[Config] = Config) -> Flask:
    app = Flask(__name__, template_folder="../templates", static_folder="../static")
    app.config.from_object(config_object)

    db.init_app(app)
    app.register_blueprint(main_bp)
    register_template_helpers(app)
    register_commands(app)

    with app.app_context():
        db.create_all()

    return app
