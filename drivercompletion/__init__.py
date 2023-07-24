from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from flask_mail import Mail
from drivercompletion.config import config
from werkzeug.middleware.dispatcher import DispatcherMiddleware

db = SQLAlchemy()
mail = Mail()


def create_app(config_class=config):
    app = Flask(__name__)
    app.config.from_object(config)
    app.wsgi_app = DispatcherMiddleware(app.wsgi_app, {
    "/drivercompletion": app
    })

    db.init_app(app)
    mail.init_app(app)

    from drivercompletion.errors.handlers import errors
    from drivercompletion.main.routes import main
    app.register_blueprint(main)
    app.register_blueprint(errors)

    return app