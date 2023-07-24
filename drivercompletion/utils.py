from drivercompletion import mail 
from flask_mail import Message
from drivercompletion.config import config

import os

def send_error_email():
    file_name = 'error.log'
    subject = 'Driver Completion Error'
    msg = Message(
                    sender=str(config.MAIL_DEFAULT_SENDER),
                    subject=subject,
                    recipients = config.SUPPORT
                )
    msg.body = 'There was a server error when trying to perform the driver completion report. Please check app log to see error'
    if file_name in os.listdir():
        file = open(file_name, 'rb')
        msg.attach(file_name, 'text/plain', file.read())
    mail.send(msg)
    return 'Administrator has been contacted.'

def json_return(data, success, msg):
    if isinstance(data, list):
        data = data
    else:
        data = [data]
    return {
        'data': data, 
        'success': success, 
        'msg': msg,
        'count': len(data)
    }
    