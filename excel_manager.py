import os
import sys
import re
from datetime import datetime
import openpyxl
from openpyxl.styles import PatternFill

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_DIR = os.path.join(BASE_DIR, 'data')
EXCEL_PATH = os.path.join(DATA_DIR, 'records.xlsx')
LOG_PATH = os.path.join(DATA_DIR, 'audit_log.txt')

BASE_HEADERS = ['活動編號', '名字', '性別', '暱稱', '學號/工號']
CYCLE_FIELDS = ['週期/日期', '身高(cm)', '體重(kg)', 'BMI(公式)', '體脂率(%)', '腰圍(cm)', '血壓(高壓/低壓)']

# ★★★ 關鍵修復：使用 Microsoft Excel 100% 相容的 8 碼 aRGB 亮黃色 ('FFFFFF00') ★★★
YELLOW_FILL = PatternFill(fill_type="solid", start_color="FFFFFF00", end_color="FFFFFF00")
NO_FILL = PatternFill(fill_type=None)

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

def is_body_fat_suspicious(gender, height, weight, body_fat):
    """
    明確的體脂率檢核機制：
    男性正常約 10%~20%，超過 30% 或小於 5% 屬於高度可疑或需特別關注。
    女性正常約 18%~28%，超過 38% 或小於 10% 屬於高度可疑或需特別關注。
    BMI 與體脂嚴重背離者（如極消瘦但報極高體脂，或重度肥胖報極低體脂）亦標黃。
    """
    try:
        bf = float(body_fat)
        h_m = float(height) / 100.0
        w_kg = float(weight)
        bmi = w_kg / (h_m * h_m)

        g = str(gender).strip()
        if g == "男":
            if bf < 5.0 or bf > 32.0:
                return True
            if bmi >= 27.0 and bf < 12.0:
                return True
            if bmi <= 18.5 and bf > 25.0:
                return True
        else: # 女
            if bf < 10.0 or bf > 40.0:
                return True
            if bmi >= 27.0 and bf < 16.0:
                return True
            if bmi <= 18.5 and bf > 30.0:
                return True
        return False
    except Exception:
        return True

def init_storage(id_field_name="學號"):
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)

    if not os.path.exists(EXCEL_PATH):
        wb = openpyxl.Workbook()
        ws = wb.active
        current_sheet_name = datetime.now().strftime('%Y-%m')
        ws.title = current_sheet_name
        headers = ['活動編號', '名字', '性別', '暱稱', id_field_name] + [f"第1週 - {f}" for f in CYCLE_FIELDS]
        ws.append(headers)
        wb.save(EXCEL_PATH)

    if not os.path.exists(LOG_PATH):
        with open(LOG_PATH, 'w', encoding='utf-8') as f:
            f.write(f"=== 健康資料管理系統 修改歷史日誌 ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')}) ===\n")

def get_or_create_sheet(wb, sheet_name, id_field_name="學號"):
    if sheet_name in wb.sheetnames:
        return wb[sheet_name]
    ws = wb.create_sheet(title=sheet_name)
    headers = ['活動編號', '名字', '性別', '暱稱', id_field_name] + [f"第1週 - {f}" for f in CYCLE_FIELDS]
    ws.append(headers)
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
        rec_gender = str(ws.cell(row=r, column=3).value or '').strip()
        rec_nick = str(ws.cell(row=r, column=4).value or '').strip()
        rec_sid = str(ws.cell(row=r, column=5).value or '').strip()

        match_group1 = (name and name == rec_name) or (nick and nick == rec_nick)
        match_group2 = (sid and sid == rec_sid) or (act_no and act_no == rec_act)

        if match_group1 and match_group2:
            return r, {
                'act_no': rec_act,
                'name': rec_name,
                'gender': rec_gender,
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

def sort_and_rewrite_user_cycles(ws, row_idx, gender, cycles_data):
    """
    排序並重組，且確實寫入黃色樣式
    """
    cycles_data.sort(key=lambda c: extract_week_number(c['cycle_label']))

    start_c = 6
    max_c = max(ws.max_column, start_c + len(cycles_data) * 7)
    for c in range(start_c, max_c + 1):
        cell = ws.cell(row=row_idx, column=c)
        cell.value = None
        cell.fill = NO_FILL

    col = start_c
    for c in cycles_data:
        # col+0: 週期/日期
        ws.cell(row=row_idx, column=col, value=c['cycle_label'])
        # col+1: 身高
        ws.cell(row=row_idx, column=col + 1, value=c['height'])
        # col+2: 體重
        ws.cell(row=row_idx, column=col + 2, value=c['weight'])

        # col+3: 原生 BMI 公式
        h_let = openpyxl.utils.get_column_letter(col + 1)
        w_let = openpyxl.utils.get_column_letter(col + 2)
        bmi_formula = f'=IFERROR(ROUND({w_let}{row_idx}/(({h_let}{row_idx}/100)^2), 1), "")'
        ws.cell(row=row_idx, column=col + 3, value=bmi_formula)

        # col+4: 體脂率 (檢核不合理則填寫黃色背景)
        bf_cell = ws.cell(row=row_idx, column=col + 4, value=c['body_fat'])
        if is_body_fat_suspicious(gender, c['height'], c['weight'], c['body_fat']):
            bf_cell.fill = YELLOW_FILL
        else:
            bf_cell.fill = NO_FILL

        # col+5: 腰圍
        ws.cell(row=row_idx, column=col + 5, value=c['waist'])
        # col+6: 血壓
        ws.cell(row=row_idx, column=col + 6, value=c['bp'])

        # 自動擴充補齊表頭
        week_tag = extract_week_title(c['cycle_label'])
        for f_idx, f_name in enumerate(CYCLE_FIELDS):
            ws.cell(row=1, column=col + f_idx, value=f"{week_tag} - {f_name}")

        col += 7

def add_user_record(sheet_name, cycle_label, name, gender, nick, sid, act_no, height, weight, body_fat, waist, sys_bp, dia_bp, id_field_name="學號"):
    init_storage(id_field_name)
    h_val = clean_number(height, "身高")
    w_val = clean_number(weight, "體重")
    bf_val = clean_number(body_fat, "體脂率")
    waist_val = clean_number(waist, "腰圍")
    sys_val = int(clean_number(sys_bp, "收縮壓(高壓)"))
    dia_val = int(clean_number(dia_bp, "舒張壓(低壓)"))
    bp_str = f"{sys_val}/{dia_val}"

    try:
        wb = openpyxl.load_workbook(EXCEL_PATH)
        ws = get_or_create_sheet(wb, sheet_name, id_field_name)

        name = str(name or '').strip()
        gender = str(gender or '男').strip()
        nick = str(nick or '').strip()
        sid = str(sid or '').strip()
        act_no = str(act_no or '').strip()

        target_row, existing_info = find_user_by_groups(ws, name, nick, sid, act_no)

        if target_row is None:
            if not (name and nick and sid and act_no):
                raise ValueError("首次填寫請完整填妥【名字、性別、暱稱、" + id_field_name + "、活動編號】以建立資料！")

            for r in range(2, ws.max_row + 1):
                cur_act = str(ws.cell(row=r, column=1).value or '').strip()
                cur_nick = str(ws.cell(row=r, column=4).value or '').strip()

                if nick and nick == cur_nick:
                    raise ValueError("⚠️ 這個暱稱已經有人使用囉，請再想想別的！")
                if act_no and act_no == cur_act:
                    raise ValueError("⚠️ 活動編號出現重複！請重新確認您的活動編號。")

            target_row = ws.max_row + 1
            ws.cell(row=target_row, column=1, value=act_no)
            ws.cell(row=target_row, column=2, value=name)
            ws.cell(row=target_row, column=3, value=gender)
            ws.cell(row=target_row, column=4, value=nick)
            ws.cell(row=target_row, column=5, value=sid)
            existing_cycles = []
        else:
            gender = existing_info['gender'] or gender
            existing_cycles = []
            col = 6
            while ws.cell(row=target_row, column=col).value is not None:
                existing_cycles.append({
                    'cycle_label': ws.cell(row=target_row, column=col).value,
                    'height': ws.cell(row=target_row, column=col + 1).value,
                    'weight': ws.cell(row=target_row, column=col + 2).value,
                    'body_fat': ws.cell(row=target_row, column=col + 4).value,
                    'waist': ws.cell(row=target_row, column=col + 5).value,
                    'bp': ws.cell(row=target_row, column=col + 6).value
                })
                col += 7

        formatted_cycle_label = f"{cycle_label} ({datetime.now().strftime('%m/%d')})"

        updated = False
        target_week_num = extract_week_number(formatted_cycle_label)
        for c in existing_cycles:
            if extract_week_number(c['cycle_label']) == target_week_num:
                c['cycle_label'] = formatted_cycle_label
                c['height'] = h_val
                c['weight'] = w_val
                c['body_fat'] = bf_val
                c['waist'] = waist_val
                c['bp'] = bp_str
                updated = True
                break

        if not updated:
            existing_cycles.append({
                'cycle_label': formatted_cycle_label,
                'height': h_val,
                'weight': w_val,
                'body_fat': bf_val,
                'waist': waist_val,
                'bp': bp_str
            })

        sort_and_rewrite_user_cycles(ws, target_row, gender, existing_cycles)
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
            str(ws.cell(row=r, column=4).value or '').strip(),
            str(ws.cell(row=r, column=5).value or '').strip(),
        ]
        if q in vals:
            target_row = r
            break

    if not target_row:
        return None

    r = target_row
    cycles = []
    col = 6
    while ws.cell(row=r, column=col).value is not None:
        cycle_label = str(ws.cell(row=r, column=col).value)
        bp_val = str(ws.cell(row=r, column=col + 6).value or "0/0")
        sys_bp, dia_bp = bp_val.split('/') if '/' in bp_val else (bp_val, "")

        cycles.append({
            "start_col": col,
            "cycle_label": cycle_label,
            "height": ws.cell(row=r, column=col + 1).value,
            "weight": ws.cell(row=r, column=col + 2).value,
            "body_fat": ws.cell(row=r, column=col + 4).value,
            "waist": ws.cell(row=r, column=col + 5).value,
            "sys_bp": sys_bp,
            "dia_bp": dia_bp
        })
        col += 7

    return {
        "sheet_name": sheet_name,
        "row_idx": r,
        "activity_no": ws.cell(row=r, column=1).value,
        "name": ws.cell(row=r, column=2).value,
        "gender": ws.cell(row=r, column=3).value,
        "nickname": ws.cell(row=r, column=4).value,
        "student_id": ws.cell(row=r, column=5).value,
        "cycles": cycles
    }

def update_user_record(sheet_name, row_idx, start_col, name, gender, nick, sid, act_no, height, weight, body_fat, waist, sys_bp, dia_bp, is_admin=False):
    init_storage()
    h_val = clean_number(height, "身高")
    w_val = clean_number(weight, "體重")
    bf_val = clean_number(body_fat, "體脂率")
    waist_val = clean_number(waist, "腰圍")
    sys_val = int(clean_number(sys_bp, "收縮壓(高壓)"))
    dia_val = int(clean_number(dia_bp, "舒張壓(低壓)"))
    bp_str = f"{sys_val}/{dia_val}"

    try:
        wb = openpyxl.load_workbook(EXCEL_PATH)
        ws = wb[sheet_name]

        existing_cycles = []
        col = 6
        while ws.cell(row=row_idx, column=col).value is not None:
            if col == start_col:
                existing_cycles.append({
                    'cycle_label': ws.cell(row=row_idx, column=col).value,
                    'height': h_val,
                    'weight': w_val,
                    'body_fat': bf_val,
                    'waist': waist_val,
                    'bp': bp_str
                })
            else:
                existing_cycles.append({
                    'cycle_label': ws.cell(row=row_idx, column=col).value,
                    'height': ws.cell(row=row_idx, column=col + 1).value,
                    'weight': ws.cell(row=row_idx, column=col + 2).value,
                    'body_fat': ws.cell(row=row_idx, column=col + 4).value,
                    'waist': ws.cell(row=row_idx, column=col + 5).value,
                    'bp': ws.cell(row=row_idx, column=col + 6).value
                })
            col += 7

        ws.cell(row=row_idx, column=1, value=str(act_no).strip())
        ws.cell(row=row_idx, column=2, value=str(name).strip())
        ws.cell(row=row_idx, column=3, value=str(gender).strip())
        ws.cell(row=row_idx, column=4, value=str(nick).strip())
        ws.cell(row=row_idx, column=5, value=str(sid).strip())

        sort_and_rewrite_user_cycles(ws, row_idx, gender, existing_cycles)
        wb.save(EXCEL_PATH)
        log_modification(f"{sid} (活動編號:{act_no})", f"更新第{(start_col-6)//7+1}次量測並重新檢核體脂率", is_admin)
        return True
    except PermissionError:
        raise PermissionError("無法更新 Excel！請確認 records.xlsx 未被其他軟體開啟。")
