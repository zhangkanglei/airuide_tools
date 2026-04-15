import pandas as pd

from openpyxl.styles import Font, Alignment
from openpyxl import load_workbook
from copy import copy
from datetime import date
import calendar
import re
import os

def process_attendance(old_path, template_path):
    # 1. 读取原始表，提取统计日期（核心，文件名/填充日期均用此年月）
    old_raw = pd.read_excel(old_path, header=None)
    title_row = old_raw.iloc[0, 0]
    date_match = re.search(r'统计日期：(\d{4})-(\d{2})-\d{2}', str(title_row))
    if date_match:
        year = int(date_match.group(1))
        month = int(date_match.group(2))
    else:
        # 兜底默认年月
        year = 2025
        month = 11

    # 计算月份信息（天数/星期/是否周末）
    _, days_in_month = calendar.monthrange(year, month)
    week_days = []
    dates = []
    is_weekend = []
    for day in range(1, days_in_month + 1):
        d = date(year, month, day)
        weekday = d.weekday()
        if weekday == 0:
            w = "一"
        elif weekday == 1:
            w = "二"
        elif weekday == 2:
            w = "三"
        elif weekday == 3:
            w = "四"
        elif weekday == 4:
            w = "五"
        elif weekday == 5:
            w = "六"
        else:
            w = "日"
        week_days.append(w)
        dates.append(day)
        is_weekend.append(w in ["六", "日"])

    # 处理原始表人员数据（核心规则优化）
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
        overtime_days = 0
        leave_days = 0
        for i, col in enumerate(daily_cols_old):
            if i >= days_in_month:
                break
            status = row[col]
            if pd.isna(status):
                daily_status.append(None)
                continue
            status_str = str(status).strip()

            # 产假处理
            if "产假" in status_str:
                sym = "产"
                daily_status.append(sym)
                continue

            # 仅"半天"判定为0.5天（排除0.5小时/0.5加班等干扰）
            elif "0.5天" in status_str or "半天" in status_str:
                sym = "半"
                daily_status.append(sym)
                if "事假" in status_str or "病假" in status_str:
                    work_days += 0.5
                    leave_days += 0.5
                elif "加班" in status_str:
                    work_days += 0
                    overtime_days += 0.5
                else:
                    work_days += 0.5
                    leave_days += 0.5
                continue

            else:
                # 替换为 新规则（无工作日/休息日，只看文字）
                # 外出 + 无休息 = 正常出勤 | 外出 + 休息 = 休息
                if "外出" in status_str:
                    if "休息" not in status_str:
                        sym = "√"
                        daily_status.append(sym)
                        work_days += 1
                    else:
                        sym = None
                        daily_status.append(sym)
                    continue

                # 病假
                elif "病假" in status_str:
                    sym = "△"
                    daily_status.append(sym)
                    leave_days += 1
                    continue

                # 事假
                elif "事假" in status_str:
                    sym = "○"
                    daily_status.append(sym)
                    leave_days += 1
                    continue

                # 加班（仅算加班天数，不算出勤）
                elif "加班" in status_str:
                    sym = "□"
                    daily_status.append(sym)
                    overtime_days += 1
                    work_days += 0
                    continue

                # 休息（周末/工作日休息均不算出勤）
                elif "休息" in status_str:
                    sym = None
                    daily_status.append(sym)
                    continue

                # 正常出勤
                else:
                    sym = "√"
                    daily_status.append(sym)
                    work_days += 1
                    continue

        # 存入人员考勤数据
        old_people[name] = {
            "daily": daily_status,
            "work_days": work_days,
            "overtime_days": overtime_days,
            "leave_days": leave_days
        }

    # 加载底板并处理核心需求：提取部门+填充日期到红框位置
    wb = load_workbook(template_path)
    ws = wb.active
    # ========== 全局统一格式（字体+居中）==========
    uni_font = Font(name="宋体", size=10)  # 统一宋体10号
    uni_align = Alignment(horizontal="center", vertical="center")  # 统一水平+垂直居中
    dept_name = ""
    # 从底板第二行第一列（A2）提取部门名（剔除无关后缀）
    dept_cell = ws.cell(row=2, column=1).value
    if dept_cell:
        dept_name = str(dept_cell).strip()
        # 剔除部门名中无关字符，适配各种底板写法
        dept_name = re.sub(r'\d{4}年|\d{1,2}月', '', dept_name)
        dept_name = dept_name.replace("考勤表", "").replace("考勤", "").replace("部门：", "").strip()
    target_row = 2  # 第2行
    target_col = 35 # AI列（对应数字35）
    # 处理合并单元格兼容（如果AI2是合并单元格，自动定位左上角）
    from openpyxl.worksheet.cell_range import CellRange
    target_range = None
    for merged_range in ws.merged_cells.ranges:
        if merged_range.min_row <= target_row <= merged_range.max_row and merged_range.min_col <= target_col <= merged_range.max_col:
            target_range = merged_range
            break

    if target_range:
        # AI2是合并单元格，给左上角赋值
        ws.cell(row=target_range.min_row, column=target_range.min_col, value=f"{year}年{month}月")
    else:
        # AI2不是合并单元格，直接赋值
        ws.cell(row=target_row, column=target_col, value=f"{year}年{month}月")


    # 更新底板表头：日期+星期
    date_row_idx = 4
    for i, d in enumerate(dates):
        col = 3 + i
        ws.cell(row=date_row_idx, column=col, value=d)
    week_row_idx = 5
    for i, w in enumerate(week_days):
        col = 3 + i
        ws.cell(row=week_row_idx, column=col, value=w)

    # 读取底板中的人员名单
    template_names = []
    start_row = 6
    max_row = ws.max_row
    for row_idx in range(start_row, max_row + 1):
        name_cell = ws.cell(row=row_idx, column=2)
        name = name_cell.value
        if pd.isna(name):
            continue
        name = str(name).strip()
        template_names.append((row_idx, name))

    # 填充人员考勤数据+保留底板样式
    for row_idx, name in template_names:
        if name in old_people:
            person = old_people[name]
            # 填充每日考勤状态
            for i, sym in enumerate(person["daily"]):
                col = 3 + i
                cell = ws.cell(row=row_idx, column=col, value=sym)
                template_cell = ws.cell(row=start_row, column=col)
                if template_cell.has_style:
                    cell.font = uni_font
                    cell.alignment = uni_align
                    cell.border = copy(template_cell.border)
            # 填充统计列：出勤/加班/请假（保留样式+自动调宽）
            # 核心修改：0值显示为空
            # 出勤天数（原33→现34列）
            work_days_val = person["work_days"] if person["work_days"] != 0 else ""
            cell = ws.cell(row=row_idx, column=34, value=work_days_val)
            cell.font = uni_font
            cell.alignment = uni_align
            cell.border = copy(ws.cell(row=start_row, column=34).border)
            # 加班天数（原34→现35列）
            overtime_days_val = person["overtime_days"] if person["overtime_days"] != 0 else ""
            cell = ws.cell(row=row_idx, column=35, value=overtime_days_val)
            cell.font = copy(ws.cell(row=start_row, column=35).font)
            cell.border = copy(ws.cell(row=start_row, column=35).border)
            cell.alignment = copy(ws.cell(row=start_row, column=35).alignment)
            # 请假天数（原35→现36列）
            leave_days_val = person["leave_days"] if person["leave_days"] != 0 else ""
            cell = ws.cell(row=row_idx, column=36, value=leave_days_val)
            cell.font = copy(ws.cell(row=start_row, column=36).font)
            cell.border = copy(ws.cell(row=start_row, column=36).border)
            cell.alignment = copy(ws.cell(row=start_row, column=36).alignment)

    # 自动调整统计列宽（原AG/AH/AI→现AH/AI/AJ）
    ws.column_dimensions['AH'].width = 8  # 出勤天数列（34列）
    ws.column_dimensions['AI'].width = 8  # 加班天数列（35列）
    ws.column_dimensions['AJ'].width = 8  # 请假天数列（36列）

    # 生成文件名：原始表年月+底板部门（如2025年11月一六级部考勤.xlsx）
    if dept_name:
        output_name = f"{year}年{month}月{dept_name}考勤.xlsx"
    else:
        output_name = f"{year}年{month}月考勤表_完成.xlsx"
    output_path = os.path.join(os.path.dirname(old_path), output_name)

    # 保存文件并返回路径和名称
    wb.save(output_path)
    return output_path, output_name