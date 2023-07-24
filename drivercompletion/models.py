from drivercompletion import db


class Orders(db.Model):
    __tablename__ = "Orders"
    OrderTrackingID = db.Column(db.Integer, primary_key=True)
    Status = db.Column(db.String(1))
    ClientID = db.Column(db.Integer)
    DeliveryTargetTo = db.Column(db.DateTime)

class OrderDrivers(db.Model):
    __tablename__ = "OrderDrivers"
    DriverID = db.Column(db.Integer, primary_key=True)
    OrderTrackingID = db.Column(db.Integer)

class ClientMaster(db.Model):
    __tablename__ = "ClientMaster"
    ClientID = db.Column(db.Integer, primary_key=True)

class Employees(db.Model):
    __tablename__ = "Employees"
    ID = db.Column(db.Integer, primary_key=True)
    DriverNo = db.Column(db.Integer)
    LastName = db.Column(db.String(35))
    FirstName = db.Column(db.String(35))
    TerminalID = db.Column(db.Integer)
    Driver = db.Column(db.String(1))
    Status = db.Column(db.String(1))
    DriverType = db.Column(db.String(1))
 

class Terminals(db.Model):
    __tablename__ = "Terminals"
    TerminalID = db.Column(db.Integer, primary_key=True)
    TerminalName = db.Column(db.String(50))

