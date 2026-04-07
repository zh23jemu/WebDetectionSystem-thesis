from extensions import db
from flask_login import UserMixin
from datetime import datetime

class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(150), unique=True, nullable=False)
    password = db.Column(db.String(150), nullable=False)
    role = db.Column(db.String(20), default='user')
    logs = db.relationship('WorkLog', backref='user', lazy=True)

class DetectionHistory(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(300))
    result_image = db.Column(db.String(300))
    upload_date = db.Column(db.DateTime, default=datetime.now)
    defect_type = db.Column(db.String(100))
    confidence = db.Column(db.Float)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'))
    # [新增] 记录该次检测判定的等级，方便溯源
    grade = db.Column(db.String(10), default='B级')

class WorkLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'))
    login_time = db.Column(db.DateTime, default=datetime.now)
    logout_time = db.Column(db.DateTime, nullable=True)

class SalesRecord(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    customer_name = db.Column(db.String(100))
    product_grade = db.Column(db.String(50))
    quantity = db.Column(db.Integer)
    total_price = db.Column(db.Float)
    sale_date = db.Column(db.DateTime, default=datetime.now)

class SystemSettings(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    key_name = db.Column(db.String(50), unique=True)
    value = db.Column(db.String(200))

#库存表
class Inventory(db.Model):
    grade = db.Column(db.String(10), primary_key=True) # 'A级', 'B级', 'C级'
    count = db.Column(db.Integer, default=0)