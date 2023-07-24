import os
from datetime import date, datetime, timedelta
from dotenv import load_dotenv

os.environ["WERKZEUG_RUN_MAIN"] = "true"
load_dotenv()

support = []
for r in  os.getenv('SUPPORT').split(','):
    support.append(str(r))

recipients = []
for r in  os.getenv('ADMINS').split(','):
    recipients.append(str(r))

class Config:
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    MAIL_SERVER = 'cdldelivers-com.mail.protection.outlook.com'
    MAIL_PORT = 25
    EMAIL = os.getenv("EMAIL")
    MAIL_USERNAME = os.getenv("EMAIL")
    MAIL_DEFAULT_SENDER = os.getenv("EMAIL")
    MAIL_PASSWORD = os.getenv("MAIL_PASS")
    SECURITY_EMAIL_SENDER = os.getenv("EMAIL")
    MAIL_USE_TLS = True
    MAIL_USE_SSL = False
    RECIPIENTS = recipients
    SUPPORT = support

    @staticmethod
    def init_app(app):
        pass


class DevelopmentConfig(Config):
    DEBUG = True
    SQLALCHEMY_DATABASE_URI = os.getenv("DEV_DATABASE_URL")


class TestingConfig(Config):
    TESTING = True
    SQLALCHEMY_DATABASE_URI = os.getenv("TEST_DATABASE_URL")


class ProductionConfig(Config):
    SQLALCHEMY_DATABASE_URI = os.getenv("DATABASE_URL") 
   


config = {
    'development': DevelopmentConfig,
    'testing': TestingConfig,
    'production': ProductionConfig,
    'default': DevelopmentConfig
}

config = config[os.getenv('FLASK_ENV')]