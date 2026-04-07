from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, session, send_file
from flask_login import login_user, login_required, logout_user, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
import os
from datetime import datetime, date, timedelta
from sqlalchemy import func
from collections import Counter
import pandas as pd
from io import BytesIO

from config import Config
from extensions import db, login_manager
from models import User, DetectionHistory, WorkLog, SalesRecord, SystemSettings, Inventory
from utils.detection import process_image

app = Flask(__name__)
app.config.from_object(Config)

db.init_app(app)
login_manager.init_app(app)
login_manager.login_view = 'login'
login_manager.login_message = "请登录后操作"

#类别映射与分级规则 ---
DEFECT_MAPPING = {
    'Crack': '裂纹',
    'Dead Knot': '死节',
    'Live Knot': '活节'
    #可在这里添加更多类别
}


def get_setting(key, default):
    with app.app_context():
        try:
            setting = SystemSettings.query.filter_by(key_name=key).first()
            if setting: return setting.value
        except:
            pass
        return default


@login_manager.user_loader
def load_user(user_id):
    return db.session.get(User, int(user_id))


#基础路由
@app.route('/')
def index():
    return redirect(url_for('dashboard')) if current_user.is_authenticated else redirect(url_for('login'))


@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated: return redirect(url_for('dashboard'))

    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        # [新增] 获取前端传来的登录意图 (user 或 admin)
        login_role = request.form.get('login_role')

        user = User.query.filter_by(username=username).first()

        if user and check_password_hash(user.password, password):
            # [核心逻辑] 校验账号角色与当前登录标签页是否一致
            if user.role != login_role:
                # 如果用户想在"管理员"标签登录普通账号，或者反之，则拒绝
                expected = "管理员" if user.role == 'admin' else "普通用户"
                flash(f'身份不匹配：该账号是“{expected}”，请切换到对应的标签页登录。', 'warning')
                return redirect(url_for('login'))

            # 校验通过，允许登录
            login_user(user)
            log = WorkLog(user_id=user.id, login_time=datetime.now())
            db.session.add(log)
            db.session.commit()
            session['current_log_id'] = log.id
            return redirect(url_for('dashboard'))

        flash('账号或密码错误', 'danger')
    return render_template('login.html')


@app.route('/logout')
@login_required
def logout():
    log_id = session.get('current_log_id')
    if log_id:
        log = db.session.get(WorkLog, log_id)
        if log:
            log.logout_time = datetime.now()
            db.session.commit()
    logout_user()
    session.clear()
    flash('已安全退出', 'success')
    return redirect(url_for('login'))


@app.route('/dashboard')
@login_required
def dashboard():
    return render_template('dashboard.html')

# [新增] 帮助中心/缺陷图谱路由
@app.route('/help')
@login_required
def help_center():
    return render_template('help.html')

#个人中心
@app.route('/profile', methods=['GET', 'POST'])
@login_required
def profile():
    if request.method == 'POST':
        old_pass = request.form.get('old_password')
        new_pass = request.form.get('new_password')
        confirm_pass = request.form.get('confirm_password')

        if not check_password_hash(current_user.password, old_pass):
            flash('旧密码错误', 'danger')
        elif new_pass != confirm_pass:
            flash('两次输入的新密码不一致', 'warning')
        else:
            current_user.password = generate_password_hash(new_pass)
            db.session.commit()
            flash('密码修改成功，请重新登录', 'success')
            return redirect(url_for('logout'))
    return render_template('profile.html')


#检测功能与库存联动
@app.route('/detect', methods=['GET', 'POST'])
@login_required
def detect():
    if request.method == 'POST':
        files = request.files.getlist('images')
        results = []
        if not files or files[0].filename == '':
            return jsonify({'error': '未选择文件'}), 400

        conf = float(get_setting('conf_threshold', 0.25))
        model_name = get_setting('current_model', 'best.pt')
        model_path = os.path.join(app.root_path, model_name)
        if not os.path.exists(model_path): model_path = app.config['MODEL_PATH']

        for file in files:
            if file:
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(filepath)

                # 1. 原始检测 (返回英文类别)
                res = process_image(filepath, app.config['RESULT_FOLDER'], model_path, conf)
                print(res)
                if 'error' in res: return jsonify([res])

                # 2. 类别映射与分级逻辑
                detected_grade = 'A级'  # 默认
                final_defect_str = "无缺陷"

                if res['is_clean']:
                    detected_grade = 'A级'
                    final_defect_str = "无缺陷"
                else:
                    # 拆分原始英文类别
                    raw_defects = res['defects'].split(',')
                    zh_defects = []

                    has_severe = False  # 是否有 C级缺陷 (裂纹/死节)
                    has_mild = False  # 是否有 B级缺陷 (活节)

                    for d in raw_defects:
                        d = d.strip()
                        # 映射为中文，找不到则保留原文
                        zh_name = DEFECT_MAPPING.get(d, d)
                        zh_defects.append(zh_name)

                        # 判定逻辑
                        if zh_name in ['裂纹', '死节']:
                            has_severe = True
                        elif zh_name == '活节':
                            has_mild = True

                    # 重新组合为中文显示的字符串
                    final_defect_str = ",".join(zh_defects)

                    # 优先级判别：只要有C级缺陷就是C级，否则如果有B级缺陷就是B级
                    if has_severe:
                        detected_grade = 'C级'
                    elif has_mild:
                        detected_grade = 'B级'
                    else:
                        detected_grade = 'B级'  # 未知缺陷默认归为B级

                # 3. 更新回结果字典，供前端展示
                res['defects'] = final_defect_str
                res['grade'] = detected_grade

                # 4. 库存联动
                inv = db.session.get(Inventory, detected_grade)
                if inv:
                    inv.count += 1
                else:
                    # 防止首次运行没有库存记录
                    new_inv = Inventory(grade=detected_grade, count=1)
                    db.session.add(new_inv)

                # 5. 存入历史记录 (存中文类别和判定等级)
                history = DetectionHistory(
                    filename=filename,
                    result_image=res['filepath'],
                    defect_type=final_defect_str,  # 这里展示的是“裂纹,活节”等原始信息
                    confidence=res.get('confidence', 0.0),
                    user_id=current_user.id,
                    grade=detected_grade  # 这里展示的是 A/B/C
                )
                db.session.add(history)
                results.append(res)

        db.session.commit()
        return jsonify(results)
    return render_template('detect.html')


#历史记录
@app.route('/history', methods=['GET', 'POST'])
@login_required
def history():
    if request.method == 'POST':
        r_id = request.form.get('id')
        rec = db.session.get(DetectionHistory, r_id)
        if rec and (current_user.role == 'admin' or rec.user_id == current_user.id):
            db.session.delete(rec)
            db.session.commit()
            return jsonify({'status': 'success'})
        return jsonify({'status': 'error', 'msg': '无权操作'}), 403

    page = request.args.get('page', 1, type=int)
    search = request.args.get('search', '')
    query = DetectionHistory.query
    if current_user.role != 'admin': query = query.filter_by(user_id=current_user.id)
    if search: query = query.filter(DetectionHistory.defect_type.contains(search))

    records = query.order_by(DetectionHistory.upload_date.desc()).paginate(page=page, per_page=12)
    return render_template('history.html', records=records)


#日志与报表
@app.route('/logs')
@login_required
def logs():
    log_query = WorkLog.query
    hist_query = DetectionHistory.query
    if current_user.role != 'admin':
        log_query = log_query.filter_by(user_id=current_user.id)
        hist_query = hist_query.filter_by(user_id=current_user.id)

    today_count = hist_query.filter(func.date(DetectionHistory.upload_date) == date.today()).count()

    total_seconds = 0
    for l in log_query.filter(WorkLog.logout_time != None).all():
        total_seconds += (l.logout_time - l.login_time).total_seconds()
    if session.get('current_log_id'):
        curr = db.session.get(WorkLog, session['current_log_id'])
        if curr: total_seconds += (datetime.now() - curr.login_time).total_seconds()
    work_time_str = f"{int(total_seconds // 3600)}小时 {int((total_seconds % 3600) // 60)}分钟"

    all_defects = db.session.query(DetectionHistory.defect_type).filter(
        DetectionHistory.user_id == current_user.id if current_user.role != 'admin' else True
    ).all()
    counter = Counter()
    for row in all_defects:
        if row.defect_type:
            for tag in row.defect_type.split(','):
                if tag.strip(): counter[tag.strip()] += 1
    stats_data = [[k, v] for k, v in counter.items()]

    trend_data = []
    trend_labels = []
    for i in range(6, -1, -1):
        target_date = date.today() - timedelta(days=i)
        count = hist_query.filter(func.date(DetectionHistory.upload_date) == target_date).count()
        trend_labels.append(target_date.strftime('%m-%d'))
        trend_data.append(count)

    return render_template('logs.html',
                           today_count=today_count, work_time=work_time_str, stats=stats_data,
                           trend_labels=trend_labels, trend_data=trend_data)


#数据导出
@app.route('/export_data')
@login_required
def export_data():
    if current_user.role == 'admin':
        records = DetectionHistory.query.all()
    else:
        records = DetectionHistory.query.filter_by(user_id=current_user.id).all()

    data = []
    for r in records:
        data.append({
            "ID": r.id, "文件名": r.filename, "时间": r.upload_date,
            "缺陷类型": r.defect_type, "判定等级": r.grade, "操作员ID": r.user_id
        })

    df = pd.DataFrame(data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    output.seek(0)
    return send_file(output, download_name=f'report_{date.today()}.xlsx', as_attachment=True)


#销售功能
@app.route('/sales', methods=['GET', 'POST'])
@login_required
def sales():
    if request.method == 'POST':
        grade = request.form.get('grade')
        qty = int(request.form.get('quantity'))

        inv = db.session.get(Inventory, grade)
        if not inv or inv.count < qty:
            flash(f'{grade} 库存不足！当前仅剩 {inv.count if inv else 0} 张', 'danger')
        else:
            inv.count -= qty
            sale = SalesRecord(
                customer_name=request.form.get('customer'),
                product_grade=grade, quantity=qty,
                total_price=float(request.form.get('price'))
            )
            db.session.add(sale)
            db.session.commit()
            flash("订单创建成功", "success")
        return redirect(url_for('sales'))

    sales_history = SalesRecord.query.order_by(SalesRecord.sale_date.desc()).all()
    inventory = Inventory.query.all()
    return render_template('sales.html', sales=sales_history, inventory=inventory)


#管理员功能
@app.route('/admin/users', methods=['GET', 'POST'])
@login_required
def admin_users():
    if current_user.role != 'admin': return redirect(url_for('dashboard'))
    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'add':
            if not User.query.filter_by(username=request.form.get('username')).first():
                new_user = User(username=request.form.get('username'),
                                password=generate_password_hash(request.form.get('password')),
                                role=request.form.get('role'))
                db.session.add(new_user)
                db.session.commit()
                flash('用户添加成功', 'success')
        elif action == 'delete':
            if int(request.form.get('user_id')) != current_user.id:
                User.query.filter_by(id=request.form.get('user_id')).delete()
                db.session.commit()
    return render_template('admin.html', users=User.query.all())


@app.route('/settings', methods=['GET', 'POST'])
@login_required
def settings():
    if current_user.role != 'admin': return redirect(url_for('dashboard'))
    if request.method == 'POST':
        val = request.form.get('conf_threshold')
        s = SystemSettings.query.filter_by(key_name='conf_threshold').first()
        if not s:
            s = SystemSettings(key_name='conf_threshold')
            db.session.add(s)
        s.value = val

        f = request.files.get('model_file')
        if f and f.filename.endswith('.pt'):
            fname = secure_filename(f.filename)
            f.save(os.path.join(app.root_path, fname))
            m = SystemSettings.query.filter_by(key_name='current_model').first()
            if not m:
                m = SystemSettings(key_name='current_model')
                db.session.add(m)
            m.value = fname
            flash(f'模型切换为 {fname}', 'success')
        db.session.commit()
        flash('设置已保存', 'success')
    return render_template('settings.html', conf=get_setting('conf_threshold', 0.25),
                           model=get_setting('current_model', 'best.pt'))


if __name__ == '__main__':
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    os.makedirs(app.config['RESULT_FOLDER'], exist_ok=True)
    with app.app_context():
        db.create_all()
        if not User.query.filter_by(username='admin').first():
            db.session.add(User(username='admin', password=generate_password_hash('123456'), role='admin'))
        if not Inventory.query.first():
            db.session.add(Inventory(grade='A级', count=0))
            db.session.add(Inventory(grade='B级', count=0))
            db.session.add(Inventory(grade='C级', count=0))
        db.session.commit()
    app.run(debug=True, host='127.0.0.1', port=5000)