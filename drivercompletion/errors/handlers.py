from flask import Blueprint, render_template
from drivercompletion.utils import send_error_email

errors = Blueprint('errors', __name__)


@errors.app_errorhandler(404)
def error_404(error):
    return render_template('404.html'), 404


@errors.app_errorhandler(403)
def error_403(error):
    return render_template('403.html'), 403


@errors.app_errorhandler(500)
def error_500(error):
    send_error_email()
    return render_template('500.html'), 500
