import os
import sys
from datetime import datetime
from flask import Flask, render_template, request, jsonify
import excel_manager
import admin_guard

if getattr(sys, 'frozen', False):
    BUNDLE_DIR = sys._MEIPASS
else:
    BUNDLE_DIR = os.path.dirname(os.path.abspath(__file__))

app = Flask(
    __name__,
    template_folder=os.path.join(BUNDLE_DIR, 'templates'),
    static_folder=os.path.join(BUNDLE_DIR, 'static')
)

RECORDS_LOCKED_MESSAGE = (
    "偵測到 records.xlsx 已被打開，請先存檔並關閉該檔案，再重新執行一次動作。"
)


def api_error_response(error, default_status=400):
    if isinstance(error, PermissionError):
        return jsonify({"status": "error", "message": RECORDS_LOCKED_MESSAGE}), 423
    return jsonify({"status": "error", "message": str(error)}), default_status


excel_manager.init_storage()

def get_current_cycle_info():
    now = datetime.now()
    month_str = now.strftime('%Y-%m')
    week_num = (now.day - 1) // 7 + 1
    cycle_label = f"第 {week_num} 週"
    return month_str, cycle_label

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/system_info', methods=['GET'])
def system_info():
    month_str, cycle_label = get_current_cycle_info()
    cfg = admin_guard.get_system_config()
    return jsonify({
        "initialized": cfg["initialized"],
        "org_type": cfg["org_type"],
        "id_label": "工號" if cfg["org_type"] == "company" else "學號",
        "current_month": month_str,
        "current_cycle": cycle_label,
        "date": datetime.now().strftime('%Y-%m-%d')
    })

@app.route('/api/initial_setup', methods=['POST'])
def initial_setup():
    """第一次開啟程式時，儲存管理員權限與機構類型（學校/公司）"""
    data = request.get_json(silent=True) or {}
    org_type = data.get('org_type', 'school')
    enable_admin = bool(data.get('enable_admin', False))
    account = data.get('account', '')
    password = data.get('password', '')

    admin_guard.save_initial_setup(org_type, enable_admin, account, password)
    id_label = "工號" if org_type == "company" else "學號"
    try:
        excel_manager.init_storage(id_label)
    except PermissionError as exc:
        return api_error_response(exc)

    return jsonify({"status": "success", "message": "初始化設定成功！"})

@app.route('/api/smart_match', methods=['POST'])
def smart_match():
    data = request.get_json(silent=True) or {}
    month_str = data.get('month')
    if not month_str:
        month_str, _ = get_current_cycle_info()

    try:
        matched = excel_manager.match_existing_user(
            month_str,
            data.get('name'),
            data.get('nick'),
            data.get('sid'),
            data.get('act_no')
        )
    except PermissionError as exc:
        return api_error_response(exc)
    if matched:
        return jsonify({"status": "matched", "data": matched})
    try:
        next_act_no = excel_manager.get_next_activity_no(month_str)
    except PermissionError as exc:
        return api_error_response(exc)
    return jsonify({"status": "none", "next_act_no": next_act_no})

@app.route('/api/admin_login', methods=['POST'])
def admin_login():
    data = request.get_json(silent=True) or {}
    name = data.get('name', '')
    nick = data.get('nick', '')  # 暱稱作為密碼
    if admin_guard.check_admin_credentials(name, nick):
        return jsonify({"status": "success", "message": "管理員特權身分驗證通過！已啟用 1 分鐘特權模式。"})
    return jsonify({"status": "error", "message": "非管理員帳號或密碼錯誤！"}), 401

@app.route('/api/add', methods=['POST'])
def add():
    data = request.get_json(silent=True) or {}
    current_month, current_cycle = get_current_cycle_info()
    cfg = admin_guard.get_system_config()
    id_label = "工號" if cfg["org_type"] == "company" else "學號"

    is_admin = bool(data.get('is_admin', False))
    if is_admin and data.get('custom_month') and data.get('custom_cycle'):
        target_month = data.get('custom_month')
        target_cycle = data.get('custom_cycle')
    else:
        target_month = current_month
        target_cycle = current_cycle

    try:
        success, message = excel_manager.add_user_record(
            target_month,
            target_cycle,
            data.get('name'),
            data.get('gender', '男'),
            data.get('nick'),
            data.get('sid'),
            data.get('act_no'),
            data.get('height'),
            data.get('weight'),
            data.get('body_fat'),
            data.get('waist'),
            data.get('sys_bp'),
            data.get('dia_bp'),
            id_field_name=id_label
        )
        if not success:
            return jsonify({"status": "error", "message": message}), 400
        return jsonify({"status": "success", "message": f"成功寫入 Excel 的 [{target_month}] 試算表 ({target_cycle})！"})
    except Exception as exc:
        return api_error_response(exc)

@app.route('/api/search', methods=['GET'])
def search():
    month = request.args.get('month', '')
    query = request.args.get('query', '')
    current_month, current_cycle = get_current_cycle_info()
    if not month:
        month = current_month

    try:
        user_data = excel_manager.search_user_all_cycles(month, query)
    except PermissionError as exc:
        return api_error_response(exc)
    if user_data:
        user_data["system_month"] = current_month
        user_data["system_cycle"] = current_cycle
        return jsonify({"status": "success", "data": user_data})
    return jsonify({"status": "error", "message": f"在 [{month}] 試算表中查無此使用者！"}), 404

@app.route('/api/update', methods=['POST'])
def update():
    data = request.get_json(silent=True) or {}
    sheet_name = data.get('sheet_name')
    cycle_label = data.get('cycle_label', '')
    is_admin = bool(data.get('is_admin', False))

    current_month, current_cycle = get_current_cycle_info()

    if not is_admin:
        if sheet_name != current_month:
            return jsonify({"status": "error", "message": "該月份已結束，已過期鎖定！"}), 403
        if current_cycle not in cycle_label:
            return jsonify({"status": "error", "message": "非本週數據已過期鎖定！"}), 403

    try:
        success, message = excel_manager.update_user_record(
            sheet_name,
            int(data['row_idx']),
            int(data['start_col']),
            data.get('name'),
            data.get('gender', '男'),
            data.get('nick'),
            data.get('sid'),
            data.get('act_no'),
            data.get('height'),
            data.get('weight'),
            data.get('body_fat'),
            data.get('waist'),
            data.get('sys_bp'),
            data.get('dia_bp'),
            is_admin=is_admin
        )
        if not success:
            return jsonify({"status": "error", "message": message}), 400
        return jsonify({"status": "success", "message": "資料修改成功！"})
    except Exception as exc:
        return api_error_response(exc)

if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000, debug=False)
