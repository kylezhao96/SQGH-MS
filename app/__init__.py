from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from flask_migrate import Migrate
from app.config import Config
from flask_cors import CORS


db = SQLAlchemy()
migrate = Migrate()
cors = CORS()


def create_app():
    app = Flask(__name__)
    app.config.from_object(Config)
    db.init_app(app)
    migrate.init_app(app, db)
    # cors.init_app(app, supports_credentials=True)
    from app.api import bp as api_bp
    app.register_blueprint(api_bp, url_prefix='/api')
    return app


from app import models
