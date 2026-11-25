import os
import sys
import re
from datetime import datetime
import openpyxl

# 取得 .exe 所在的外部資料夾路徑
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_DIR = os.path.join(BASE_DIR, 'data')
EXCEL_PATH = os.path.join(DATA_DIR, 'records.xlsx')
LOG_PATH = os.path.join(DATA_DIR, 'audit_log.txt')

BASE_HEADERS = ['活動編號', '名字', '暱稱', '學號']
CYCLE_FIELDS = ['週期/日期', '身高(cm)', '體重(kg)', 'BMI(公式)', '腰圍(cm)', '血壓(高壓/低壓)']

def clean_number(val, field_name="數值"):
    if val is None or str(val).strip() == "":
        raise ValueError(f"{field_name}不可為空")
    cleaned = re.sub(r'[^\d.]', '', str(val).strip())
    if not cleaned:
        raise ValueError(f"{field_name}數值格式錯誤: '{val}'")
    return float(cleaned)

def extract_week_number(cycle_label):
    match = re.search(r'第\s*(\d+)\s*週', str(cycle_label))
    if match:
        return int(match.group(1))
    return 999

def extract_week_title(cycle_label):
    match = re.search(r'第\s*(\d+)\s*週', str(cycle_label))
    if match:
        return f"第{match.group(1)}週"
    return "量測"

def init_storage():
    """每次啟動檢查：若偵測不到 data 資料夾、Excel 檔或 Log 檔時自動生成"""
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)

    if not os.path.exists(EXCEL_PATH):
        wb = openpyxl.Workbook()
        ws = wb.active
        current_sheet_name = datetime.now().strftime('%Y-%m')
        ws.title = current_sheet_name
        initial_headers = BASE_HEADERS + [f"第1次 - {f}" for f in CYCLE_FIELDS]
        ws.append(initial_headers)
        wb.save(EXCEL_PATH)

    if not os.path.exists(LOG_PATH):
        with open(LOG_PATH, 'w', encoding='utf-8') as f:
            f.write(f"=== 健康資料管理系統 修改歷史日誌 (建立於 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}) ===\n")

def get_or_create_sheet(wb, sheet_name):
    if sheet_name in wb.sheetnames:
        return wb[sheet_name]
    ws = wb.create_sheet(title=sheet_name)
    initial_headers = BASE_HEADERS + [f"第1次 - {f}" for f in CYCLE_FIELDS]
    ws.append(initial_headers)
    return ws

def log_modification(identifier, changes, is_admin=False):
    init_storage()
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    prefix = "[管理員特權]" if is_admin else "[一般修改]"
    log_entry = f"{prefix} [{timestamp}] 對象: {identifier} | 變更: {changes}\n"
    with open(LOG_PATH, 'a', encoding='utf-8') as f:
        f.write(log_entry)

def find_user_by_groups(ws, name, nick, sid, act_no):
    name = str(name or '').strip()
    nick = str(nick or '').strip()
    sid = str(sid or '').strip()
    act_no = str(act_no or '').strip()

    for r in range(2, ws.max_row + 1):
        rec_act = str(ws.cell(row=r, column=1).value or '').strip()
        rec_name = str(ws.cell(row=r, column=2).value or '').strip()
        rec_nick = str(ws.cell(row=r, column=3).value or '').strip()
        rec_sid = str(ws.cell(row=r, column=4).value or '').strip()

        match_group1 = (name and name == rec_name) or (nick and nick == rec_nick)
        match_group2 = (sid and sid == rec_sid) or (act_no and act_no == rec_act)

        if match_group1 and match_group2:
            return r, {
                'act_no': rec_act,
                'name': rec_name,
                'nick': rec_nick,
                'sid': rec_sid
            }
    return None, None

def match_existing_user(sheet_name, name, nick, sid, act_no):
    init_storage()
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    if sheet_name not in wb.sheetnames:
        return None
    ws = wb[sheet_name]
    _, record = find_user_by_groups(ws, name, nick, sid, act_no)
    return record

def sort_and_rewrite_user_cycles(ws, row_idx, cycles_data):
    cycles_data.sort(key=lambda c: extract_week_number(c['cycle_label']))

    max_c = max(ws.max_column, 5 + len(cycles_data) * 6)
    for c in range(5, max_c + 1):
        ws.cell(row=row_idx, column=c, value=None)

    col = 5
    for c in cycles_data:
        ws.cell(row=row_idx, column=col, value=c['cycle_label'])
        ws.cell(row=row_idx, column=col + 1, value=c['height'])
        ws.cell(row=row_idx, column=col + 2, value=c['weight'])

        h_letter = openpyxl.utils.get_column_letter(col + 1)
        w_letter = openpyxl.utils.get_column_letter(col + 2)
        bmi_formula = f'=IFERROR(ROUND({w_letter}{row_idx}/(({h_letter}{row_idx}/100)^2), 1), "")'
        ws.cell(row=row_idx, column=col + 3, value=bmi_formula)

        ws.cell(row=row_idx, column=col + 4, value=c['waist'])
        ws.cell(row=row_idx, column=col + 5, value=c['bp'])

        week_tag = extract_week_title(c['cycle_label'])
        for f_idx, f_name in enumerate(CYCLE_FIELDS):
            target_col = col + f_idx
            ws.cell(row=1, column=target_col, value=f"{week_tag} - {f_name}")

        col += 6

def add_user_record(sheet_name, cycle_label, name, nick, sid, act_no, height, weight, waist, sys_bp, dia_bp):
    init_storage()
    h_val = clean_number(height, "身高")
    w_val = clean_number(weight, "體重")
    waist_val = clean_number(waist, "腰圍")
    sys_val = int(clean_number(sys_bp, "收縮壓(高壓)"))
    dia_val = int(clean_number(dia_bp, "舒張壓(低壓)"))
    bp_str = f"{sys_val}/{dia_val}"

    try:
        wb = openpyxl.load_workbook(EXCEL_PATH)
        ws = get_or_create_sheet(wb, sheet_name)

        name = str(name or '').strip()
        nick = str(nick or '').strip()
        sid = str(sid or '').strip()
        act_no = str(act_no or '').strip()

        target_row, existing_info = find_user_by_groups(ws, name, nick, sid, act_no)

        if target_row is None:
            if not (name and nick and sid and act_no):
                raise ValueError("首次填寫請完整填妥【名字、暱稱、學號、活動編號】以建立資料！")

            for r in range(2, ws.max_row + 1):
                cur_act = str(ws.cell(row=r, column=1).value or '').strip()
                cur_nick = str(ws.cell(row=r, column=3).value or '').strip()

                if nick and nick == cur_nick:
                    raise ValueError("⚠️ 這個暱稱已經有人使用囉，請再想想別的！")

                if act_no and act_no == cur_act:
                    raise ValueError("⚠️ 活動編號出現重複！請重新確認您的活動編號。")

            target_row = ws.max_row + 1
            ws.cell(row=target_row, column=1, value=act_no)
            ws.cell(row=target_row, column=2, value=name)
            ws.cell(row=target_row, column=3, value=nick)
            ws.cell(row=target_row, column=4, value=sid)
            existing_cycles = []
        else:
            existing_cycles = []
            col = 5
            while ws.cell(row=target_row, column=col).value is not None:
                existing_cycles.append({
                    'cycle_label': ws.cell(row=target_row, column=col).value,
                    'height': ws.cell(row=target_row, column=col + 1).value,
                    'weight': ws.cell(row=target_row, column=col + 2).value,
                    'waist': ws.cell(row=target_row, column=col + 4).value,
                    'bp': ws.cell(row=target_row, column=col + 5).value
                })
                col += 6

        formatted_cycle_label = f"{cycle_label} ({datetime.now().strftime('%m/%d')})"
        
        updated = False
        target_week_num = extract_week_number(formatted_cycle_label)
        for c in existing_cycles:
            if extract_week_number(c['cycle_label']) == target_week_num:
                c['cycle_label'] = formatted_cycle_label
                c['height'] = h_val
                c['weight'] = w_val
                c['waist'] = waist_val
                c['bp'] = bp_str
                updated = True
                break

        if not updated:
            existing_cycles.append({
                'cycle_label': formatted_cycle_label,
                'height': h_val,
                'weight': w_val,
                'waist': waist_val,
                'bp': bp_str
            })

        sort_and_rewrite_user_cycles(ws, target_row, existing_cycles)
        wb.save(EXCEL_PATH)
        return True
    except PermissionError:
        raise PermissionError("無法寫入 Excel！請確認 records.xlsx 未被其他軟體開啟。")

def search_user_all_cycles(sheet_name, query):
    init_storage()
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=False)
    if sheet_name not in wb.sheetnames:
        return None

    ws = wb[sheet_name]
    q = str(query).strip()

    target_row = None
    for r in range(2, ws.max_row + 1):
        vals = [
            str(ws.cell(row=r, column=1).value or '').strip(),
            str(ws.cell(row=r, column=2).value or '').strip(),
            str(ws.cell(row=r, column=3).value or '').strip(),
            str(ws.cell(row=r, column=4).value or '').strip(),
        ]
        if q in vals:
            target_row = r
            break

    if not target_row:
        return None

    r = target_row
    cycles = []
    col = 5
    while ws.cell(row=r, column=col).value is not None:
        cycle_label = str(ws.cell(row=r, column=col).value)
        bp_val = str(ws.cell(row=r, column=col + 5).value or "0/0")
        sys_bp, dia_bp = bp_val.split('/') if '/' in bp_val else (bp_val, "")

        cycles.append({
            "start_col": col,
            "cycle_label": cycle_label,
            "height": ws.cell(row=r, column=col + 1).value,
            "weight": ws.cell(row=r, column=col + 2).value,
            "waist": ws.cell(row=r, column=col + 4).value,
            "sys_bp": sys_bp,
            "dia_bp": dia_bp
        })
        col += 6

    return {
        "sheet_name": sheet_name,
        "row_idx": r,
        "activity_no": ws.cell(row=r, column=1).value,
        "name": ws.cell(row=r, column=2).value,
        "nickname": ws.cell(row=r, column=3).value,
        "student_id": ws.cell(row=r, column=4).value,
        "cycles": cycles
    }

def update_user_record(sheet_name, row_idx, start_col, name, nick, sid, act_no, height, weight, waist, sys_bp, dia_bp, is_admin=False):
    init_storage()
    h_val = clean_number(height, "身高")
    w_val = clean_number(weight, "體重")
    waist_val = clean_number(waist, "腰圍")
    sys_val = int(clean_number(sys_bp, "收縮壓(高壓)"))
    dia_val = int(clean_number(dia_bp, "舒張壓(低壓)"))
    bp_str = f"{sys_val}/{dia_val}"

    try:
        wb = openpyxl.load_workbook(EXCEL_PATH)
        ws = wb[sheet_name]

        existing_cycles = []
        col = 5
        while ws.cell(row=row_idx, column=col).value is not None:
            if col == start_col:
                existing_cycles.append({
                    'cycle_label': ws.cell(row=row_idx, column=col).value,
                    'height': h_val,
                    'weight': w_val,
                    'waist': waist_val,
                    'bp': bp_str
                })
            else:
                existing_cycles.append({
                    'cycle_label': ws.cell(row=row_idx, column=col).value,
                    'height': ws.cell(row=row_idx, column=col + 1).value,
                    'weight': ws.cell(row=row_idx, column=col + 2).value,
                    'waist': ws.cell(row=row_idx, column=col + 4).value,
                    'bp': ws.cell(row=row_idx, column=col + 5).value
                })
            col += 6

        ws.cell(row=row_idx, column=1, value=str(act_no).strip())
        ws.cell(row=row_idx, column=2, value=str(name).strip())
        ws.cell(row=row_idx, column=3, value=str(nick).strip())
        ws.cell(row=row_idx, column=4, value=str(sid).strip())

        sort_and_rewrite_user_cycles(ws, row_idx, existing_cycles)
        wb.save(EXCEL_PATH)
        log_modification(f"{sid} (活動編號:{act_no})", f"更新第{(start_col-5)//6+1}次量測並重新排序表頭", is_admin)
        return True
    except PermissionError:
        raise PermissionError("無法更新 Excel！請確認 records.xlsx 未被其他軟體開啟。")
