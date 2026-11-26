import os
import sys

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_PATH = os.path.join(BASE_DIR, 'admin_config.txt')

def ascii_encrypt(text: str, shift: int = 5) -> str:
    """對字串進行簡單 ASCII 位移加密"""
    return "".join(chr((ord(c) + shift) % 65536) for c in text)

def ascii_decrypt(cipher: str, shift: int = 5) -> str:
    """ASCII 位移解密"""
    return "".join(chr((ord(c) - shift) % 65536) for c in cipher)

def is_system_initialized():
    """檢查系統是否已完成第一次設定"""
    return os.path.exists(CONFIG_PATH)

def save_initial_setup(org_type, enable_admin, account="", password=""):
    """儲存第一次精靈設定"""
    enc_acc = ascii_encrypt(account) if enable_admin else ""
    enc_pwd = ascii_encrypt(password) if enable_admin else ""
    enabled_str = "true" if enable_admin else "false"

    with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
        f.write("# 系統基礎組態設定檔\n")
        f.write(f"ORG_TYPE={org_type}\n")
        f.write(f"ADMIN_ENABLED={enabled_str}\n")
        f.write(f"ENC_ACCOUNT={enc_acc}\n")
        f.write(f"ENC_PASSWORD={enc_pwd}\n")

def get_system_config():
    """讀取當前系統設定（場域類型、特權是否啟用）"""
    if not os.path.exists(CONFIG_PATH):
        return {"initialized": False, "org_type": "school", "admin_enabled": False}

    org_type = "school"
    admin_enabled = False
    acc = ""
    pwd = ""

    try:
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line.startswith('ORG_TYPE='):
                    org_type = line.split('=', 1)[1].strip().lower()
                elif line.startswith('ADMIN_ENABLED='):
                    admin_enabled = line.split('=', 1)[1].strip().lower() == 'true'
                elif line.startswith('ENC_ACCOUNT='):
                    acc = ascii_decrypt(line.split('=', 1)[1].strip())
                elif line.startswith('ENC_PASSWORD='):
                    pwd = ascii_decrypt(line.split('=', 1)[1].strip())
    except Exception:
        pass

    return {
        "initialized": True,
        "org_type": org_type,
        "admin_enabled": admin_enabled,
        "account": acc,
        "password": pwd
    }

def check_admin_credentials(name_input, nick_input):
    cfg = get_system_config()
    if not cfg["admin_enabled"]:
        return False
    return (str(name_input).strip() == cfg["account"]) and (str(nick_input).strip() == cfg["password"])
