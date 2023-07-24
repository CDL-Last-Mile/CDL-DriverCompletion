from flask import render_template, request, Blueprint
from drivercompletion.api_func.report import get_driver_report
from drivercompletion.utils import json_return


main = Blueprint('main', __name__)


@main.route("/")
@main.route("/home")
def home():
    return render_template('home.html')

@main.route("/report", methods=["GET", "POST"])
def get_report():
    report, success, msg = get_driver_report()
    return json_return(report, success, msg)


    

