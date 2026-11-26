import os
import sys
import time
import socket
import threading
import webview

# 無終端機時將輸出導向 Null，防止打包後閃退
if sys.stdout is None or sys.stderr is None:
    class NullWriter:
        def write(self, data): pass
        def flush(self): pass
    sys.stdout = NullWriter()
    sys.stderr = NullWriter()

from app import app
import excel_manager
import admin_guard

PORT = 5000

# 單一實例鎖定（防止多開霸佔 port）
lock_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
try:
    lock_socket.bind(('127.0.0.1', 49151))
except OSError:
    sys.exit(0)

def start_flask():
    app.run(host='127.0.0.1', port=PORT, debug=False, use_reloader=False)

def on_closed():
    try:
        lock_socket.close()
    except Exception:
        pass
    os._exit(0)

if __name__ == '__main__':
    # 確保資料夾與基礎儲存就緒
    excel_manager.init_storage()

    # 啟動 Flask 背景服務
    flask_thread = threading.Thread(target=start_flask, daemon=True)
    flask_thread.start()

    time.sleep(1.0)

    # 啟動原生獨立視窗
    window = webview.create_window(
        title="健康資料管理系統",
        url=f"http://127.0.0.1:{PORT}",
        width=1100,      # 調寬以容納左右雙欄
        height=880,
        resizable=True,
        min_size=(680, 640)
    )


    window.events.closed += on_closed
    webview.start()
    on_closed()
