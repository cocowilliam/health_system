import os
import re
import threading
from datetime import datetime

import openpyxl
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter

DATA_DIR = os.path.join(os.getcwd(), "data")
FILE_PATH = os.path.join(DATA_DIR, "records.xlsx")
FILE_LOCK = threading.RLock()

YELLOW_FILL = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
HEADER_FILL = PatternFill(start_color="EFEFEF", end_color="EFEFEF", fill_type="solid")
WEEK_HEADER_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

THIN_BORDER = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)

WEEK_SUB_FIELDS = [
    "身高(cm)",
    "體重(kg)",
    "體脂率(%)",
    "腰圍(cm)",
    "收縮壓(mmHg)",
    "舒張壓(mmHg)",
    "BMI",
]


def _text(value):
    if value is None:
        return ""
    return str(value).strip()


def _sheet_name(month):
    value = _text(month)
    if not value:
        value = datetime.now().strftime("%Y-%m")
    value = value.replace("/", "-")
    if re.fullmatch(r"\d{4}-\d{1,2}", value):
        year, month_num = value.split("-")
        return f"{int(year):04d}-{int(month_num):02d}"
    if re.fullmatch(r"\d{1,2}月", value):
        return f"{datetime.now().year:04d}-{int(value[:-1]):02d}"
    if re.fullmatch(r"\d{1,2}", value):
        return f"{datetime.now().year:04d}-{int(value):02d}"
    if value.endswith("月"):
        value = value[:-1]
    if re.fullmatch(r"\d{4}-\d{1,2}", value):
        year, month_num = value.split("-")
        return f"{int(year):04d}-{int(month_num):02d}"
    return value[:31]


def _safe_text(value):
    value = _text(value)
    # Prevent spreadsheet formula/control expression injection.
    if value.startswith(("=", "+", "-", "@")):
        return "'" + value
    return value


def _number(value, decimal=True):
    if value in (None, ""):
        return ""
    try:
        number = float(str(value).strip())
        if not (-1_000_000 < number < 1_000_000):
            raise ValueError("numeric value out of range")
        return round(number, 1) if decimal else int(round(number))
    except (TypeError, ValueError, OverflowError):
        raise ValueError(f"數值格式不正確：{value}")


def init_storage(id_label=None):
    """Initialize storage and the Excel file. id_label is kept for API compatibility."""
    id_label = id_label or "學號/工號"
    os.makedirs(DATA_DIR, exist_ok=True)
    with FILE_LOCK:
        if not os.path.exists(FILE_PATH):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = _sheet_name(datetime.now().strftime("%Y-%m"))
            _setup_fixed_headers(ws, id_label)
            for week in range(1, 6):
                ensure_week_columns(ws, week)
            ws.freeze_panes = "F3"
            ws.auto_filter.ref = f"A2:{get_column_letter(ws.max_column)}{ws.max_row}"
            wb.save(FILE_PATH)
            wb.close()
            return

        wb = openpyxl.load_workbook(FILE_PATH)
        try:
            current_sheet = _sheet_name(datetime.now().strftime("%Y-%m"))
            needs_save = current_sheet not in wb.sheetnames
            ws = wb[current_sheet] if not needs_save else wb.create_sheet(title=current_sheet)

            if ws.max_row < 2:
                _setup_fixed_headers(ws, id_label)
                needs_save = True

            existing_weeks = get_existing_weeks(ws)
            for week in range(1, 6):
                if week not in existing_weeks:
                    ensure_week_columns(ws, week)
                    existing_weeks = get_existing_weeks(ws)
                    needs_save = True

            if needs_save:
                ws.freeze_panes = "F3"
                ws.auto_filter.ref = f"A2:{get_column_letter(ws.max_column)}{ws.max_row}"
                wb.save(FILE_PATH)
        finally:
            wb.close()


def _load_workbook(write=False):
    init_storage()
    return openpyxl.load_workbook(FILE_PATH, data_only=not write)


def auto_fit_columns(ws):
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            value = _text(cell.value)
            length = sum(2 if ord(char) > 127 else 1 for char in value)
            max_len = max(max_len, length)
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 24)


def _resolve_existing_sheet(wb, month_str):
    target = _sheet_name(month_str)
    if target in wb.sheetnames:
        return wb[target]
    candidates = {
        f"{target.replace('-', '')}",
        f"{target}月",
        f"{int(target.split('-')[1])}月",
        f"{target.split('-')[1]}月",
    }
    for name in wb.sheetnames:
        if name in candidates:
            return wb[name]
    return None


def get_or_create_sheet(wb, month_str, id_label="學號/工號"):
    sheet_name = _sheet_name(month_str)
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.create_sheet(title=sheet_name)
    if ws.max_row < 2:
        _setup_fixed_headers(ws, id_label)
    existing_weeks = get_existing_weeks(ws)
    for week in range(1, 6):
        if week not in existing_weeks:
            ensure_week_columns(ws, week)
            existing_weeks = get_existing_weeks(ws)
    ws.freeze_panes = "F3"
    ws.sheet_view.showGridLines = True
    ws.auto_filter.ref = f"A2:{get_column_letter(ws.max_column)}{ws.max_row}"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_layout_view = False
    return ws


def _setup_fixed_headers(ws, id_label="學號/工號"):
    fixed_headers = ["活動編號", "姓名", "性別", "暱稱", id_label]
    for col_idx, text in enumerate(fixed_headers, 1):
        ws.merge_cells(start_row=1, start_column=col_idx, end_row=2, end_column=col_idx)
        cell = ws.cell(row=1, column=col_idx, value=text)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        ws.cell(row=2, column=col_idx).border = THIN_BORDER
        cell.border = THIN_BORDER
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 24


def match_existing_user(month_str, name="", nickname="", student_id="", act_id="", gender=""):
    """Signature matches /api/smart_match: month, name, nick, sid, act_no."""
    init_storage()
    with FILE_LOCK:
        wb = _load_workbook(write=False)
        try:
            if month_str:
                ws = _resolve_existing_sheet(wb, month_str)
                sheets = [ws] if ws is not None else []
            else:
                sheets = wb.worksheets

            for ws in sheets:
                if ws is None or ws.max_row < 3:
                    continue
                for row in range(3, ws.max_row + 1):
                    values = {
                        "act_id": _text(ws.cell(row=row, column=1).value),
                        "name": _text(ws.cell(row=row, column=2).value),
                        "gender": _text(ws.cell(row=row, column=3).value),
                        "nickname": _text(ws.cell(row=row, column=4).value),
                        "student_id": _text(ws.cell(row=row, column=5).value),
                    }
                    hit = (
                        (_text(student_id) and values["student_id"] == _text(student_id))
                        or (_text(act_id) and values["act_id"] == _text(act_id))
                        or (
                            _text(name)
                            and _text(nickname)
                            and values["name"].lower() == _text(name).lower()
                            and values["nickname"].lower() == _text(nickname).lower()
                        )
                    )
                    if hit:
                        return values
            return None
        finally:
            wb.close()


def check_identity_conflict(ws, act_id, nickname, student_id, target_row=None):
    for row in range(3, ws.max_row + 1):
        if target_row is not None and row == target_row:
            continue
        cur_act = _text(ws.cell(row=row, column=1).value)
        cur_nick = _text(ws.cell(row=row, column=4).value)
        cur_sid = _text(ws.cell(row=row, column=5).value)
        if act_id and cur_act.lower() == _text(act_id).lower():
            return f"活動編號 '{act_id}' 已被使用！"
        if nickname and cur_nick.lower() == _text(nickname).lower():
            return f"暱稱 '{nickname}' 已被使用！"
        if student_id and cur_sid == _text(student_id):
            return f"學號/工號 '{student_id}' 已存在於此表中！"
    return None


def get_existing_weeks(ws):
    weeks = {}
    col = 6
    while col <= ws.max_column:
        match = re.search(r"第\s*(\d+)\s*週", _text(ws.cell(row=1, column=col).value))
        if match:
            weeks[int(match.group(1))] = col
            col += len(WEEK_SUB_FIELDS)
        else:
            col += 1
    return weeks


def ensure_week_columns(ws, week_num):
    week_num = int(week_num)
    weeks = get_existing_weeks(ws)
    if week_num in weeks:
        return weeks[week_num]

    all_weeks = sorted(set(list(weeks.keys()) + [week_num]))
    user_data = {}
    for row in range(3, ws.max_row + 1):
        basic = [_text(ws.cell(row=row, column=col).value) for col in range(1, 6)]
        records = {}
        for week, start_col in weeks.items():
            records[week] = [ws.cell(row=row, column=start_col + i).value for i in range(len(WEEK_SUB_FIELDS))]
        user_data[row] = (basic, records)

    for merged in list(ws.merged_cells.ranges):
        if merged.min_col >= 6:
            ws.unmerge_cells(str(merged))
    for row in ws.iter_rows(min_row=1, max_row=max(ws.max_row, 2), min_col=6, max_col=max(ws.max_column, 12)):
        for cell in row:
            cell.value = None
            cell.fill = PatternFill(fill_type=None)
            cell.border = Border()

    cur_col = 6
    for week in all_weeks:
        ws.merge_cells(start_row=1, start_column=cur_col, end_row=1,
                       end_column=cur_col + len(WEEK_SUB_FIELDS) - 1)
        title = ws.cell(row=1, column=cur_col, value=f"第 {week} 週")
        title.alignment = Alignment(horizontal="center", vertical="center")
        title.font = Font(bold=True)
        title.fill = WEEK_HEADER_FILL
        for i, sub_field in enumerate(WEEK_SUB_FIELDS):
            col = cur_col + i
            header = ws.cell(row=2, column=col, value=sub_field)
            header.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            header.font = Font(bold=True)
            header.fill = HEADER_FILL
            header.border = THIN_BORDER
            ws.cell(row=1, column=col).border = THIN_BORDER
        cur_col += len(WEEK_SUB_FIELDS)

    for row, (basic, records) in user_data.items():
        for col, value in enumerate(basic, 1):
            ws.cell(row=row, column=col, value=value)
        for week, start_col in get_existing_weeks(ws).items():
            values = records.get(week)
            if values:
                for i, value in enumerate(values):
                    ws.cell(row=row, column=start_col + i, value=value)

    ws.freeze_panes = "F3"
    auto_fit_columns(ws)
    return get_existing_weeks(ws)[week_num]


def _week_from_label(label):
    match = re.search(r"(\d+)", _text(label))
    if not match:
        raise ValueError("週次格式不正確")
    week = int(match.group(1))
    if week < 1 or week > 53:
        raise ValueError("週次必須介於 1 至 53")
    return week


def _bmi_formula(row, start_col):
    h = get_column_letter(start_col)
    w = get_column_letter(start_col + 1)
    return f'=IFERROR(ROUND({w}{row}/(({h}{row}/100)*({h}{row}/100)),1),"")'


def _write_measurements(ws, row, start_col, gender, height, weight, body_fat, waist, sbp, dbp):
    values = [
        _number(height),
        _number(weight),
        _number(body_fat),
        _number(waist),
        _number(sbp, decimal=False),
        _number(dbp, decimal=False),
        _bmi_formula(row, start_col),
    ]
    for i, value in enumerate(values):
        cell = ws.cell(row=row, column=start_col + i, value=value)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER
        if i == 2:
            try:
                bf = float(body_fat) if body_fat not in (None, "") else None
                high = (gender == "男" and bf is not None and bf > 25) or (
                    gender == "女" and bf is not None and bf > 30
                )
                cell.fill = YELLOW_FILL if high else PatternFill(fill_type=None)
            except (TypeError, ValueError):
                cell.fill = PatternFill(fill_type=None)


def add_user_record(month, week, name, gender, nickname, student_id, act_id,
                    height, weight, body_fat, waist, sbp, dbp, id_field_name="學號"):
    init_storage()
    week = _week_from_label(week)
    if not (_text(name) or _text(nickname)) or not (_text(student_id) or _text(act_id)):
        return False, "名字/暱稱至少填一項，學號/工號/活動編號至少填一項！"
    try:
        h = _number(height)
        w = _number(weight)
        bf = _number(body_fat)
        waist_num = _number(waist)
        sys_num = _number(sbp, decimal=False)
        dia_num = _number(dbp, decimal=False)
        if h == "" or w == "" or bf == "" or waist_num == "" or sys_num == "" or dia_num == "":
            return False, "身高、體重、體脂率、腰圍與高低血壓皆為必填！"
        if not (30 <= h <= 250) or not (10 <= w <= 400):
            return False, "身高或體重數值超出合理範圍！"
        if not (0 <= bf <= 100) or not (20 <= waist_num <= 300):
            return False, "體脂率或腰圍數值超出合理範圍！"
        if not (60 <= sys_num <= 300) or not (30 <= dia_num <= 200) or sys_num <= dia_num:
            return False, "血壓數值不合理；收縮壓需大於舒張壓！"
    except ValueError as exc:
        return False, str(exc)
    with FILE_LOCK:
        wb = openpyxl.load_workbook(FILE_PATH)
        try:
            ws = get_or_create_sheet(wb, month, id_field_name)
            target_row = None
            sid = _text(student_id)
            for row in range(3, ws.max_row + 1):
                if sid and _text(ws.cell(row=row, column=5).value) == sid:
                    target_row = row
                    break

            conflict = check_identity_conflict(
                ws, _text(act_id), _text(nickname), sid, target_row
            )
            if conflict:
                return False, conflict

            if target_row is None:
                target_row = ws.max_row + 1
                basics = [_safe_text(act_id), _safe_text(name), _safe_text(gender),
                          _safe_text(nickname), _safe_text(student_id)]
                for col, value in enumerate(basics, 1):
                    cell = ws.cell(row=target_row, column=col, value=value)
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = THIN_BORDER

            start_col = ensure_week_columns(ws, week)
            _write_measurements(
                ws, target_row, start_col, _text(gender), height, weight,
                body_fat, waist, sbp, dbp
            )
            ws.freeze_panes = "F3"
            auto_fit_columns(ws)
            wb.save(FILE_PATH)
            return True, "新增成功！"
        finally:
            wb.close()


def search_user_all_cycles(month, query):
    query = _text(query).lower()
    if not query:
        return None
    init_storage()
    with FILE_LOCK:
        wb = openpyxl.load_workbook(FILE_PATH, data_only=False)
        try:
            ws = _resolve_existing_sheet(wb, month)
            if ws is None:
                return None
            weeks = get_existing_weeks(ws)
            for row in range(3, ws.max_row + 1):
                identity = [
                    _text(ws.cell(row=row, column=col).value).lower()
                    for col in range(1, 6)
                ]
                if not any(query in value for value in identity):
                    continue
                cycles = []
                for week in sorted(weeks):
                    start_col = weeks[week]
                    height = ws.cell(row=row, column=start_col).value
                    weight = ws.cell(row=row, column=start_col + 1).value
                    body_fat = ws.cell(row=row, column=start_col + 2).value
                    waist = ws.cell(row=row, column=start_col + 3).value
                    sys_bp = ws.cell(row=row, column=start_col + 4).value
                    dia_bp = ws.cell(row=row, column=start_col + 5).value
                    bmi = ws.cell(row=row, column=start_col + 6).value
                    if isinstance(bmi, str) and bmi.startswith("="):
                        bmi = None
                        try:
                            if height not in (None, "") and weight not in (None, ""):
                                bmi = round(float(weight) / ((float(height) / 100) ** 2), 1)
                        except (TypeError, ValueError, ZeroDivisionError):
                            bmi = None
                    cycles.append({
                        "cycle_label": f"第 {week} 週",
                        "start_col": start_col,
                        "height": height,
                        "weight": weight,
                        "body_fat": body_fat,
                        "waist": waist,
                        "sys_bp": sys_bp,
                        "dia_bp": dia_bp,
                        "bmi": bmi,
                    })
                return {
                    "sheet_name": ws.title,
                    "row_idx": row,
                    "activity_no": ws.cell(row=row, column=1).value,
                    "name": ws.cell(row=row, column=2).value,
                    "gender": ws.cell(row=row, column=3).value,
                    "nickname": ws.cell(row=row, column=4).value,
                    "student_id": ws.cell(row=row, column=5).value,
                    "cycles": cycles,
                }
            return None
        finally:
            wb.close()


def update_user_record(sheet_name, row_idx, start_col, name, gender, nickname, student_id,
                       act_id, height, weight, body_fat, waist, sbp, dbp, is_admin=False):
    init_storage()
    with FILE_LOCK:
        wb = openpyxl.load_workbook(FILE_PATH)
        try:
            ws = _resolve_existing_sheet(wb, sheet_name)
            if ws is None:
                return False, "找不到指定月份的紀錄！"
            row_idx = int(row_idx)
            start_col = int(start_col)
            if row_idx < 3 or row_idx > ws.max_row:
                return False, "列位置無效！"

            if is_admin:
                basics = {
                    1: _safe_text(act_id),
                    2: _safe_text(name),
                    3: _safe_text(gender),
                    4: _safe_text(nickname),
                    5: _safe_text(student_id),
                }
                for col, value in basics.items():
                    ws.cell(row=row_idx, column=col, value=value)

                conflict = check_identity_conflict(
                    ws, _text(act_id), _text(nickname), _text(student_id), row_idx
                )
                if conflict:
                    return False, conflict
            else:
                gender = _text(ws.cell(row=row_idx, column=3).value) or "男"

            found = any(start_col == col for col in get_existing_weeks(ws).values())
            if not found:
                return False, "週次欄位無效，未修改資料！"

            _write_measurements(
                ws, row_idx, start_col, gender, height, weight,
                body_fat, waist, sbp, dbp
            )
            for col in range(1, 6):
                ws.cell(row=row_idx, column=col).alignment = Alignment(horizontal="center", vertical="center")
                ws.cell(row=row_idx, column=col).border = THIN_BORDER
            ws.freeze_panes = "F3"
            auto_fit_columns(ws)
            wb.save(FILE_PATH)
            return True, "更新成功！"
        finally:
            wb.close()
