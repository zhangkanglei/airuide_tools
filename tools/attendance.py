import pandas as pd
from openpyxl import load_workbook
from copy import copy
from datetime import date
import calendar
import re
import os

def process_attendance(old_path, template_path):
    # 1. 读取原始表，提取统计日期
    old_raw = pd.read_excel(old_path, header=None)
    title_row = old_raw.iloc[0, 0]
    date_match = re.search(r'统计日期：(\d{4})-(\d{2})-\d{2}', str(title_row))
    if date_match:
        year = int(date_match.group(1))
        month = int(date_match.group(2))
    else:
        year = 2025
        month = 11

    # 计算月份信息
    _, days_in_month = calendar.monthrange(year, month)
    week_days = []
    dates = []
    is_weekend = []
    for day in range(1, days_in_month +1):
        d = date(year, month, day)
        weekday = d.weekday()
        if weekday == 0:
            w = "一"
        elif weekday ==1:
            w="二"
        elif weekday ==2:
            w="三"
        elif weekday ==3:
            w="四"
        elif weekday ==4:
            w="五"
        elif weekday ==5:
            w="六"
        else:
            w="日"
        week_days.append(w)
        dates.append(day)
        is_weekend.append(w in ["六", "日"])

    # 处理原始表人员数据
    data_rows_old = old_raw.iloc[4:]
    name_col = 0
    daily_cols_old = list(range(31, 62))
    old_people = {}
    for idx, row in data_rows_old.iterrows():
        name = row[name_col]
        if pd.isna(name):
            continue
        name = str(name).strip()
        daily_status = []
        work_days = 0
        overtime_days =0
        leave_days=0
        for i, col in enumerate(daily_cols_old):
            if i >= days_in_month:
                break
            status = row[col]
            if pd.isna(status):
                daily_status.append(None)
                continue
            status_str = str(status).strip()
            current_is_weekend = is_weekend[i]

            if "产假" in status_str:
                sym = "产"
                daily_status.append(sym)
                continue
            elif "0.5天" in status_str or "半天" in status_str:
                sym = "半"
                daily_status.append(sym)
                if "事假" in status_str or "病假" in status_str:
                    work_days +=0.5
                    leave_days +=0.5
                elif "加班" in status_str:
                    work_days +=0
                    overtime_days +=0.5
                else:
                    work_days +=0.5
                    leave_days +=0.5
                continue
            else:
                if "外出" in status_str:
                    sym = "√"
                    daily_status.append(sym)
                    work_days +=1
                    continue
                if "病假" in status_str:
                    sym = "△"
                    daily_status.append(sym)
                    leave_days +=1
                    continue
                elif "事假" in status_str:
                    sym = "○"
                    daily_status.append(sym)
                    leave_days +=1
                    continue
                elif "加班" in status_str:
                    sym = "□"
                    daily_status.append(sym)
                    overtime_days +=1
                    work_days +=0
                    continue
                elif current_is_weekend and "休息" in status_str:
                    sym = None
                    daily_status.append(sym)
                    continue
                elif "休息" in status_str:
                    sym = None
                    daily_status.append(sym)
                    continue
                else:
                    sym = "√"
                    daily_status.append(sym)
                    work_days +=1
                    continue

        old_people[name] = {
            "daily": daily_status,
            "work_days": work_days,
            "overtime_days": overtime_days,
            "leave_days": leave_days
        }

    # 加载底板
    wb = load_workbook(template_path)
    ws = wb.active

    # 从底板的第二行提取部门和日期，用来生成文件名
    title_cell = ws.cell(row=2, column=1).value
    dept_name = ""
    file_month = month  # 兜底用原始表的月份
    if title_cell:
        title_str = str(title_cell).strip()
        # 先提取月份，支持 11月、01月 这种格式
        month_match = re.search(r'(\d{1,2})月', title_str)
        if month_match:
            file_month = int(month_match.group(1))
        # 提取干净的部门名，去掉年份、月份、考勤相关的后缀
        dept_name = title_str
        dept_name = re.sub(r'\d{4}年', '', dept_name)  # 去掉年份
        dept_name = re.sub(r'\d{1,2}月', '', dept_name)  # 去掉月份
        dept_name = dept_name.replace("考勤表", "").replace("考勤", "").replace("部门：","").strip()

    # 删除多余列
    if days_in_month <31:
        ws.delete_cols(33)

    # 更新表头
    date_row_idx =4
    for i, d in enumerate(dates):
        col = 3 + i
        ws.cell(row=date_row_idx, column=col, value=d)

    week_row_idx=5
    for i, w in enumerate(week_days):
        col=3+i
        ws.cell(row=week_row_idx, column=col, value=w)

    # 读取底板人员
    template_names = []
    start_row =6
    max_row = ws.max_row
    for row_idx in range(start_row, max_row+1):
        name_cell = ws.cell(row=row_idx, column=2)
        name = name_cell.value
        if pd.isna(name):
            continue
        name = str(name).strip()
        template_names.append((row_idx, name))

    # 填充数据
    for row_idx, name in template_names:
        if name in old_people:
            person = old_people[name]
            for i, sym in enumerate(person["daily"]):
                col=3+i
                cell = ws.cell(row=row_idx, column=col, value=sym)
                template_cell = ws.cell(row=start_row, column=col)
                if template_cell.has_style:
                    cell.font = copy(template_cell.font)
                    cell.border = copy(template_cell.border)
                    cell.fill = copy(template_cell.fill)
                    cell.alignment = copy(template_cell.alignment)
            # 统计列
            cell = ws.cell(row=row_idx, column=33, value=person["work_days"])
            cell.font = copy(ws.cell(row=start_row, column=33).font)
            cell.border = copy(ws.cell(row=start_row, column=33).border)
            cell.alignment = copy(ws.cell(row=start_row, column=33).alignment)

            cell = ws.cell(row=row_idx, column=34, value=person["overtime_days"])
            cell.font = copy(ws.cell(row=start_row, column=34).font)
            cell.border = copy(ws.cell(row=start_row, column=34).border)
            cell.alignment = copy(ws.cell(row=start_row, column=34).alignment)

            cell = ws.cell(row=row_idx, column=35, value=person["leave_days"])
            cell.font = copy(ws.cell(row=start_row, column=35).font)
            cell.border = copy(ws.cell(row=start_row, column=35).border)
            cell.alignment = copy(ws.cell(row=start_row, column=35).alignment)

    # 保存输出
    output_name = f"{year}年{month}月考勤表_完成.xlsx"
    output_path = os.path.join(os.path.dirname(old_path), output_name)
    # 自动调整统计列的列宽，避免19.5显示成20
    ws.column_dimensions['AG'].width = 8  # 出勤天数列
    ws.column_dimensions['AH'].width = 8  # 加班天数列
    ws.column_dimensions['AI'].width = 8  # 请假天数列

    # 保存输出，根据底板第二行的部门和月份生成文件名
    if dept_name:
        output_name = f"{file_month}月{dept_name}考勤.xlsx"
    else:
        # 兜底，如果没提取到就用默认名字
        output_name = f"{year}年{month}月考勤表_完成.xlsx"
    output_path = os.path.join(os.path.dirname(old_path), output_name)
    wb.save(output_path)

    return output_path, output_name