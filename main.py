import os
import sys
import time
import socket
import threading
import webview
from app import app
import excel_manager
import admin_guard

# 無終端機時將輸出重定向，防止打包閃退
if sys.stdout is None or sys.stderr is None:
    class NullWriter:
        def write(self, data): pass
        def flush(self): pass
    sys.stdout = NullWriter()
    sys.stderr = NullWriter()

PORT = 5000

# ★★★ 關鍵一：單一實例鎖定（Single Instance Lock）★★★
# 透過全域本機 Socket 鎖，若發現已在運行，絕不開第二個 Port，直接退出！
lock_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
try:
    lock_socket.bind(('127.0.0.1', 49151))  # 鎖定專屬旗標 port
except OSError:
    # 已經有程式在執行了，直接結束本次啟動，避免重複霸佔 port
    sys.exit(0)

def start_flask():
    """在背景啟動 Flask"""
    app.run(host='127.0.0.1', port=PORT, debug=False, use_reloader=False)

def on_closed():
    """★★★ 關鍵二：使用者按右上角 [X] 關閉視窗時，整個程式 100% 徹底結束退出 ★★★"""
    try:
        lock_socket.close()
    except Exception:
        pass
    os._exit(0)

if __name__ == '__main__':
    # 1. 確保設定檔與 Excel 存在
    admin_guard.ensure_config_exists()
    excel_manager.init_storage()

    # 2. 啟動後端伺服器執行緒
    flask_thread = threading.Thread(target=start_flask, daemon=True)
    flask_thread.start()

    # 等待 Flask 就緒
    time.sleep(1.0)

    # 3. 建立獨立原生桌面視窗（無黑窗，像正式軟體一樣）
    window = webview.create_window(
        title="健康資料管理系統",
        url=f"http://127.0.0.1:{PORT}",
        width=700,
        height=850,
        resizable=True,
        min_size=(500, 600)
    )

    # 綁定視窗關閉事件：視窗關掉，程式立刻完全終止
    window.events.closed += on_closed

    # 啟動視窗主迴圈
    webview.start()

    # 視窗結束後強制清理退出
    on_closed()
