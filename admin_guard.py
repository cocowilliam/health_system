import os
import sys

# 取得 .exe 所在的外部資料夾路徑
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_PATH = os.path.join(BASE_DIR, 'admin_config.txt')

ENABLED = True
DEFAULT_ACCOUNT = "health"
DEFAULT_PASSWORD = "lth0612"

def ensure_config_exists():
    """每次啟動檢查：若偵測不到 admin_config.txt 則自動生成"""
    if not os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
                f.write("# 管理員特權帳號密碼設定檔 (可隨時用記事本修改)\n")
                f.write(f"ACCOUNT={DEFAULT_ACCOUNT}\n")
                f.write(f"PASSWORD={DEFAULT_PASSWORD}\n")
        except Exception:
            pass

def get_admin_credentials():
    """讀取帳密設定，若檔案遺失則重新生成"""
    ensure_config_exists()
    account = DEFAULT_ACCOUNT
    password = DEFAULT_PASSWORD
    try:
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line.startswith('ACCOUNT='):
                    account = line.split('ACCOUNT=', 1)[1].strip()
                elif line.startswith('PASSWORD='):
                    password = line.split('PASSWORD=', 1)[1].strip()
    except Exception:
        pass
    return account, password

def check_admin_credentials(name_input, sid_input):
    if not ENABLED:
        return False
    account, password = get_admin_credentials()
    return (str(sid_input).strip() == account) and (str(name_input).strip() == password)
