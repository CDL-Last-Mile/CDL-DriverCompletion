from flask import render_template
from flask_mail import Message
from sqlalchemy import Date
from drivercompletion import db, mail
from drivercompletion.models import( 
    Orders, 
    Employees,
    ClientMaster,
    Terminals, 
    OrderDrivers)
from drivercompletion.config import config
from datetime import datetime, date

import os
import xlsxwriter
import pandas as pd
import openpyxl


def get_non_complete_count(): 
    non_complete_count = db.session.query(Orders.OrderTrackingID, ClientMaster.ClientID, OrderDrivers.DriverID)
    non_complete_count = non_complete_count.join(OrderDrivers, Orders.OrderTrackingID == OrderDrivers.OrderTrackingID)
    non_complete_count = non_complete_count.join(ClientMaster, ClientMaster.ClientID == Orders.ClientID)
    return non_complete_count

def get_completed_count(): 
    complete_count = db.session.query(OrderDrivers.OrderTrackingID, Orders.ClientID, ClientMaster.ClientID)
    complete_count = complete_count.join(Orders, OrderDrivers.OrderTrackingID == Orders.OrderTrackingID)
    complete_count = complete_count.join(ClientMaster, ClientMaster.ClientID == Orders.ClientID)
    return complete_count

def get_uncomplete_count(employee_id):
    today = datetime.today()
    today = today.date()
    non_complete_count = get_non_complete_count()
    response = non_complete_count.filter(
        OrderDrivers.DriverID == employee_id, 
        Orders.Status == 'N',
        Orders.DeliveryTargetTo.cast(Date) == today)
    response = len(response.all())
    return response


def get_complete_count(employee_id):
    today = datetime.today()
    today = today.date()
    status_list = ['N', 'D', 'L']
    complete_count = get_completed_count()
    response = complete_count.filter(
        OrderDrivers.DriverID == employee_id, 
        ~Orders.Status.in_(status_list),
        Orders.DeliveryTargetTo.cast(Date) == today)
    response = len(response.all())
    return response

def get_driver_report(driver_type=None, target_date=None, driver_center=None, driver_numbers=None):
    drivers = []
    success = False
    try:
        today = datetime.today()
        today = today.date()
        date_filter = today
        if target_date is not None:
            date_filter = target_date
        status_list = ['N', 'D', 'L']
        driver_filter = 'C'
        if driver_type is not None:
            driver_filter = driver_type
        dbquery = db.session.query(
            Terminals.TerminalID.label('terminal_id'), 
            Terminals.TerminalName.label('terminal_name'), 
            Employees.ID.label('driver_id'),
            Employees.DriverNo.label('driver_no'), 
            Employees.LastName.label('last_name'), 
            Employees.FirstName.label('first_name'),
            (db.session.query(db.func.count(OrderDrivers.OrderTrackingID)).join(Orders, Orders.OrderTrackingID == OrderDrivers.OrderTrackingID).join(ClientMaster, ClientMaster.ClientID == Orders.ClientID).filter(OrderDrivers.DriverID == Employees.ID, Orders.Status == 'N', Orders.DeliveryTargetTo.cast(Date) == date_filter)
            ).label('noncomplete_count'),
            (db.session.query(db.func.count(OrderDrivers.OrderTrackingID)).join(Orders, Orders.OrderTrackingID == OrderDrivers.OrderTrackingID).join(ClientMaster, ClientMaster.ClientID == Orders.ClientID).filter(OrderDrivers.DriverID == Employees.ID, ~Orders.Status.in_(status_list), Orders.DeliveryTargetTo.cast(Date) == date_filter)
            ).label('complete_count')
        )
        dbquery = dbquery.join(Terminals, Terminals.TerminalID == Employees.TerminalID)
        dbquery = dbquery.filter(Employees.Status == 'A', Employees.Driver == 'Y', Employees.DriverType == driver_filter)
        if driver_center is not None: 
            dbquery = dbquery.filter(Terminals.TerminalID.in_(driver_center))
        if driver_numbers is not None and len(driver_numbers) > 0: 
            dbquery = dbquery.filter(Employees.DriverNo.in_(driver_numbers))
        dbquery = dbquery.group_by(
            Terminals.TerminalID,
            Terminals.TerminalName,
            Employees.ID,
            Employees.DriverNo,
            Employees.LastName,
            Employees.FirstName)
        dbquery = dbquery.order_by(
            Terminals.TerminalID,
            Terminals.TerminalName,
            Employees.ID,
            Employees.DriverNo,
            Employees.LastName,
            Employees.FirstName)

        drivers = [r._asdict() for r in dbquery.all()]
        total_summary ={
            'active': 0, 
            'complete': 0, 
            'total': 0,
            'percent_complete': 0, 
            'name': 'Total'
        }
        summary = {}
        for driver in drivers:
            divisor =  (int(driver['complete_count']) + int(driver['noncomplete_count']))
            if divisor != 0:
                driver['completion_percentage'] = str(round((driver['complete_count']/divisor) * 100, 2)) + '%'
            else:
                driver['completion_percentage'] = str(0) + "%"
            driver.pop('terminal_id')
            driver.pop('driver_id')
            terminal = driver['terminal_name']
            if terminal in summary:
                summary[terminal]['active'] += int(driver['noncomplete_count'])
                summary[terminal]['complete'] += int(driver['complete_count'])
                summary[terminal]['total'] += int(driver['complete_count']) + int(driver['noncomplete_count'])
            else:
                summary[terminal] = {}
                summary[terminal]['active'] = int(driver['noncomplete_count'])
                summary[terminal]['complete'] = int(driver['complete_count'])
                summary[terminal]['name'] = terminal
                summary[terminal]['total'] = int(driver['complete_count']) + int(driver['noncomplete_count'])
            total_summary['active'] += int(driver['noncomplete_count'])
            total_summary['complete'] += int(driver['complete_count'])
            total_summary['total'] += int(driver['complete_count']) + int(driver['noncomplete_count'])
            if summary[terminal]['total'] > 0:
                summary[terminal]['percent_complete'] = round((summary[terminal]['complete']/summary[terminal]['total']) * 100, 2) 
            if total_summary['total'] > 0:
                total_summary['percent_complete'] = round((total_summary['complete']/total_summary['total']) * 100, 2) 
            

        # Convert to dataframe 
        df = pd.DataFrame(drivers)
        now = datetime.now()
        time = now.strftime("%H_%M_%S")
        file_name = 'Driver_Completion_Report_' + time +'.xlsx'
        df.to_excel(file_name, sheet_name='Driver_Completion')
        report_time = datetime.now()
        date_time = report_time.strftime("%m/%d/%Y, %H:%M:%S")
        subject = 'Driver Completion Report'
        msg = Message(
                        sender=str(config.EMAIL),
                        subject=subject,
                        recipients = config.RECIPIENTS
                    )

        msg.html = render_template('driver_report.html', day_of_report=date_time, total_summary=total_summary, summary=list(summary.values()))
        file = open(file_name, 'rb')

        
        msg.attach(file_name, '	application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', file.read())
        mail.send(msg)
        success = True
        msg = 'Driver report generated succesfully'
    except Exception as e:
        print(e)
        msg = str(e)
        
    return drivers, success, msg