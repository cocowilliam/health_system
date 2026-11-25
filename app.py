import os
import sys
import time
import threading
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

excel_manager.init_storage()

# ★★★ 核心：心跳與自動關閉看門狗 ★★★
last_heartbeat = time.time()
server_started = False

def watchdog():
    """看門狗：當伺服器啟動超過 8 秒，且已超過 5 秒沒收到任何網頁心跳，自動徹底結束處理程序"""
    global last_heartbeat
    while True:
        time.sleep(1)
        if server_started and (time.time() - last_heartbeat > 5):
            # 所有分頁皆已關閉，退出程式並釋放 Port
            os._exit(0)

threading.Thread(target=watchdog, daemon=True).start()

@app.route('/api/heartbeat', methods=['POST'])
def heartbeat():
    global last_heartbeat, server_started
    last_heartbeat = time.time()
    server_started = True
    return jsonify({"status": "alive"})

@app.route('/api/shutdown', methods=['POST'])
def shutdown():
    """瀏覽器分頁關閉時觸發"""
    def do_exit():
        time.sleep(0.5)
        os._exit(0)
    threading.Thread(target=do_exit).start()
    return jsonify({"status": "shutting_down"})

# 以下原有路由保持不變：
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
    return jsonify({
        "current_month": month_str,
        "current_cycle": cycle_label,
        "date": datetime.now().strftime('%Y-%m-%d')
    })

@app.route('/api/smart_match', methods=['POST'])
def smart_match():
    data = request.get_json(silent=True) or {}
    month_str = data.get('month')
    if not month_str:
        month_str, _ = get_current_cycle_info()

    matched = excel_manager.match_existing_user(
        month_str,
        data.get('name'),
        data.get('nick'),
        data.get('sid'),
        data.get('act_no')
    )
    if matched:
        return jsonify({"status": "matched", "data": matched})
    return jsonify({"status": "none"})

@app.route('/api/admin_login', methods=['POST'])
def admin_login():
    data = request.get_json(silent=True) or {}
    name = data.get('name', '')
    sid = data.get('student_id', '')
    if admin_guard.check_admin_credentials(name, sid):
        return jsonify({"status": "success", "message": "特權帳號密碼驗證通過！已啟用管理員特權模式。"})
    return jsonify({"status": "error", "message": "非管理員帳號或密碼錯誤！"}), 401

@app.route('/api/add', methods=['POST'])
def add():
    data = request.get_json(silent=True) or {}
    current_month, current_cycle = get_current_cycle_info()

    is_admin = bool(data.get('is_admin', False))
    if is_admin and data.get('custom_month') and data.get('custom_cycle'):
        target_month = data.get('custom_month')
        target_cycle = data.get('custom_cycle')
    else:
        target_month = current_month
        target_cycle = current_cycle

    try:
        excel_manager.add_user_record(
            target_month,
            target_cycle,
            data.get('name'),
            data.get('nick'),
            data.get('sid'),
            data.get('act_no'),
            data.get('height'),
            data.get('weight'),
            data.get('waist'),
            data.get('sys_bp'),
            data.get('dia_bp')
        )
        return jsonify({"status": "success", "message": f"成功寫入 Excel 的 [{target_month}] 試算表 ({target_cycle})！"})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 400

@app.route('/api/search', methods=['GET'])
def search():
    month = request.args.get('month', '')
    query = request.args.get('query', '')
    current_month, current_cycle = get_current_cycle_info()
    if not month:
        month = current_month

    user_data = excel_manager.search_user_all_cycles(month, query)
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
        excel_manager.update_user_record(
            sheet_name,
            int(data['row_idx']),
            int(data['start_col']),
            data.get('name'),
            data.get('nick'),
            data.get('sid'),
            data.get('act_no'),
            data.get('height'),
            data.get('weight'),
            data.get('waist'),
            data.get('sys_bp'),
            data.get('dia_bp'),
            is_admin=is_admin
        )
        return jsonify({"status": "success", "message": "資料修改成功！"})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 400

if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000, debug=False)
