# -*- coding: utf-8 -*-
"""
生日名单导出工具
输入：学生信息Excel（含「幼儿园」「小学总表」两个sheet）+ 老师信息Excel（含「全部名单」sheet）+ 月份
输出：指定月份的师生生日名单 Word 文档（.docx）

格式规则（参照 2026.8月份师生生日名单.doc）：
- 标题：X年X月份教职工及学生生日名单（居中加粗16pt）
- 每天一个日期段：X月X日（居中加粗12pt），其下为当天过生日的人
- 排序：幼儿园在最前（小→中→大→大大、班级号升序），小学按年级班级顺序（一~六、班级号升序），老师最后并另起一行
- 同一天内：同班级多人合并为"班级+姓名1、姓名2"，不同班级/组之间用"，"连接；老师输出为"姓名老师，姓名老师"
"""
import os
import re
import logging
from datetime import datetime

import pandas as pd
import xlrd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# 中文数字 → 阿拉伯数字（一~十）
CN_NUM = {
    '一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
    '六': 6, '七': 7, '八': 8, '九': 9, '十': 10,
}

# 幼儿园班级类型排序（小班 → 中班 → 大班 → 大大班，由低到高）
KG_TYPE_ORDER = {'小': 0, '中': 1, '大': 2, '大大': 3}


def cn_to_int(s):
    """中文数字转阿拉伯数字，如 '三' -> 3"""
    s = str(s).strip()
    if s in CN_NUM:
        return CN_NUM[s]
    # 十以内的个位数（'十' = 10，'一十'不常见）
    if s == '十':
        return 10
    return 0


def parse_birthday(raw):
    """
    把各种格式的生日值统一成 4 位字符串 MMDD（如 '0101'）
    支持：数字101（补零）、文本'0101'、'1月5日'等
    无效/空值返回 ''
    """
    if raw is None:
        return ''
    if isinstance(raw, float) and pd.isna(raw):
        return ''
    s = str(raw).strip()
    if s == '' or s.lower() == 'nan' or s.lower() == 'none':
        return ''
    # 去掉浮点尾巴：101.0 -> 101
    if s.endswith('.0'):
        s = s[:-2]
    if s.isdigit():
        return s.zfill(4)
    # 中文日期格式：1月5日 -> 0105
    m = re.match(r'(\d{1,2})月(\d{1,2})日', s)
    if m:
        return f"{int(m.group(1)):02d}{int(m.group(2)):02d}"
    return s


def convert_kg_class(cls):
    """幼儿园班级格式转换：'大大一班' -> '大大（1）班'；非预期格式原样返回"""
    m = re.match(r'^(小|中|大|大大)([一二三四五六七八九十]+)班$', cls)
    if m:
        num = cn_to_int(m.group(2))
        if num:
            return f"{m.group(1)}（{num}）班"
    return cls


def parse_xx_class(cls):
    """小学班级解析：'一（2）班' -> (年级, 班级号)，用于排序；解析失败返回 (99, 99)"""
    m = re.match(r'^([一二三四五六七八九十]+)（(\d+)）班$', cls)
    if m:
        return (cn_to_int(m.group(1)), int(m.group(2)))
    return (99, 99)


def infer_year(student_filename):
    """从学生表文件名推断年份（学年结束年份），如 '2025~2026年...' -> 2026；失败用当前年"""
    matches = re.findall(r'20\d{2}', student_filename)
    if matches:
        return max(int(x) for x in matches)
    return datetime.now().year


def read_students(student_path):
    """
    读取学生信息Excel（sheet：幼儿园、小学总表）
    返回列表，元素为 dict:
      bd: MMDD字符串, sort_key: 排序用元组, cls: 显示班级名, name: 姓名
    """
    xl = pd.ExcelFile(student_path)
    students = []

    # ---- 幼儿园 ----
    kg = xl.parse('幼儿园')
    for idx, row in kg.iterrows():
        bd = parse_birthday(row.get('生日'))
        if not bd:
            continue
        name = str(row.get('学生姓名')).strip() if pd.notna(row.get('学生姓名')) else ''
        cls_raw = str(row.get('班级')).strip() if pd.notna(row.get('班级')) else ''
        if not name or not cls_raw:
            continue
        cls = convert_kg_class(cls_raw)
        # 班级类型序号 + 班级号；班级号解析失败放最后
        m = re.match(r'^(小|中|大|大大)（(\d+)）班$', cls)
        type_no = KG_TYPE_ORDER.get(m.group(1), 99) if m else 99
        class_no = int(m.group(2)) if m else 99
        students.append({
            'bd': bd,
            'sort_key': (0, type_no, class_no, idx),  # 0=幼儿园层级
            'cls': cls,
            'name': name,
        })

    # ---- 小学总表 ----
    xx = xl.parse('小学总表')
    for idx, row in xx.iterrows():
        bd = parse_birthday(row.get('生日'))
        if not bd:
            continue
        name = str(row.get('学生姓名')).strip() if pd.notna(row.get('学生姓名')) else ''
        cls_raw = str(row.get('班级')).strip() if pd.notna(row.get('班级')) else ''
        if not name or not cls_raw:
            continue
        grade, class_no = parse_xx_class(cls_raw)
        students.append({
            'bd': bd,
            'sort_key': (1, grade, class_no, idx),  # 1=小学层级
            'cls': cls_raw,
            'name': name,
        })
    return students


def read_teachers(teacher_path):
    """
    读取老师信息Excel（sheet：全部名单，列：姓名、出生日期）
    出生日期取最后一列'出生日期'（MMDD文本，如'0102'）
    支持 .xls（xlrd）和 .xlsx（pandas）
    返回列表，元素为 dict: bd(MMDD), name, seq(原顺序)
    """
    teachers = []
    if teacher_path.lower().endswith('.xlsx'):
        df = pd.read_excel(teacher_path, sheet_name='全部名单')
        name_col, bd_col = None, None
        for col in df.columns:
            h = str(col).strip()
            if h == '姓名':
                name_col = col
            elif '出生日期' in h:
                bd_col = col  # 同名多列时取最后一列（MMDD文本列）
        for seq, row in df.iterrows():
            name = str(row[name_col]).strip() if name_col is not None and pd.notna(row[name_col]) else ''
            bd = parse_birthday(row[bd_col]) if bd_col is not None else ''
            if name and bd:
                teachers.append({'bd': bd, 'name': name, 'seq': seq})
    else:
        wb = xlrd.open_workbook(teacher_path)
        sh = wb.sheet_by_name('全部名单')
        hdr = [str(sh.cell_value(0, c)).strip() for c in range(sh.ncols)]
        name_col = None
        bd_col = None
        for c, h in enumerate(hdr):
            if h == '姓名':
                name_col = c
            elif '出生日期' in h:
                bd_col = c  # 同名多列时取最后一列（MMDD文本列）
        for r in range(1, sh.nrows):
            name = str(sh.cell_value(r, name_col)).strip() if name_col is not None else ''
            bd = parse_birthday(sh.cell_value(r, bd_col)) if bd_col is not None else ''
            if name and bd:
                teachers.append({'bd': bd, 'name': name, 'seq': r})
    return teachers


def build_docx(year, month, students, teachers, output_path):
    """生成生日名单 Word 文档"""
    doc = Document()

    # 默认字体：宋体 12pt
    normal = doc.styles['Normal']
    normal.font.name = 'Times New Roman'
    normal.font.size = Pt(12)
    normal.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    def _set_font(run, size, bold):
        run.font.name = 'Times New Roman'
        run.font.size = Pt(size)
        run.font.bold = bold
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    # 标题（居中加粗16pt）
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_para.add_run(f"{year}年{month}月份教职工及学生生日名单")
    _set_font(run, 16, True)

    # 按日期 1~31 分组输出
    for day in range(1, 32):
        target_bd = f"{month:02d}{day:02d}"

        day_students = sorted(
            (s for s in students if s['bd'] == target_bd),
            key=lambda s: s['sort_key']
        )
        day_teachers = [t for t in teachers if t['bd'] == target_bd]
        if not day_students and not day_teachers:
            continue

        # 日期行（居中加粗12pt）
        date_para = doc.add_paragraph()
        date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = date_para.add_run(f"{month}月{day}日")
        _set_font(run, 12, True)

        # 学生行：同班级合并为"班级+姓名1、姓名2"，不同组用"，"连接
        if day_students:
            groups = []
            for s in day_students:
                if groups and groups[-1]['cls'] == s['cls']:
                    groups[-1]['names'].append(s['name'])
                else:
                    groups.append({'cls': s['cls'], 'names': [s['name']]})
            student_text = '，'.join(f"{g['cls']}{'、'.join(g['names'])}" for g in groups)
            stu_para = doc.add_paragraph()
            stu_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = stu_para.add_run(student_text)
            _set_font(run, 12, False)

        # 老师行：另起一行，"姓名老师，姓名老师"
        if day_teachers:
            teacher_text = '，'.join(f"{t['name']}老师" for t in day_teachers)
            tea_para = doc.add_paragraph()
            tea_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = tea_para.add_run(teacher_text)
            _set_font(run, 12, False)

        # 日期之间空一行
        doc.add_paragraph()

    doc.save(output_path)


def process_birthday_export(student_path, teacher_path, month):
    """
    主入口：读取数据 -> 按月份筛选 -> 生成Word
    :param student_path: 学生信息Excel路径（.xlsx）
    :param teacher_path: 老师信息Excel路径（.xls / .xlsx）
    :param month: 月份 1~12
    :return: (输出文件路径, 输出文件名)
    """
    logging.info(f"========== 开始处理生日名单导出，月份：{month} ==========")

    year = infer_year(os.path.basename(student_path))
    students = read_students(student_path)
    teachers = read_teachers(teacher_path)
    logging.info(f"读取学生 {len(students)} 条，老师 {len(teachers)} 条，年份：{year}")

    output_name = f"{year}年{month}月份教职工及学生生日名单.docx"
    output_path = os.path.join(os.path.dirname(student_path), output_name)
    build_docx(year, month, students, teachers, output_path)

    logging.info(f"========== 生日名单导出完成：{output_name} ==========")
    return output_path, output_name
