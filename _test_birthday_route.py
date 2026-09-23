# -*- coding: utf-8 -*-
"""生日名单路由的 Flask 测试客户端脚本"""
import sys
import os

sys.path.insert(0, r'E:\PycharmProjects\airuide_tools')
os.chdir(r'E:\PycharmProjects\airuide_tools')
from app import app

client = app.test_client()

# GET 页面
r = client.get('/birthday_export')
html = r.get_data(as_text=True)
print('GET /birthday_export:', r.status_code, '页面含8月选项:', 'value="8"' in html)

# 首页包含新工具
r2 = client.get('/')
print('首页含生日名单导出:', '生日名单导出' in r2.get_data(as_text=True))

# POST 正常流程（8月），并验证清理
with open(r'E:\艾瑞德\综合部\老师学生生日导出\2025~2026年学生生日最新.xlsx', 'rb') as f1, \
     open(r'E:\艾瑞德\综合部\老师学生生日导出\全体教职工生日名单.xls', 'rb') as f2:
    data = {
        'student_file': (f1, '2025~2026年学生生日最新.xlsx'),
        'teacher_file': (f2, '全体教职工生日名单.xls'),
        'month': '8',
    }
    r3 = client.post('/birthday_export', data=data, content_type='multipart/form-data')
    cd = r3.headers.get('Content-Disposition', '')
    print('POST 8月:', r3.status_code, '下载文件名:', cd[:120])
    r3.close()  # 关闭 response，触发 call_on_close 清理

# POST 缺文件
r4 = client.post('/birthday_export', data={'month': '8'}, content_type='multipart/form-data')
print('POST 缺文件:', r4.status_code, r4.get_data(as_text=True)[:40])

# POST 月份非法
with open(r'E:\艾瑞德\综合部\老师学生生日导出\2025~2026年学生生日最新.xlsx', 'rb') as f1, \
     open(r'E:\艾瑞德\综合部\老师学生生日导出\全体教职工生日名单.xls', 'rb') as f2:
    data = {'student_file': (f1, 's.xlsx'), 'teacher_file': (f2, 't.xls'), 'month': '13'}
    r5 = client.post('/birthday_export', data=data, content_type='multipart/form-data')
    print('POST 月份13:', r5.status_code, r5.get_data(as_text=True)[:40])

# 检查 uploads 目录无残留
leftovers = [f for f in os.listdir('uploads') if f.endswith('.docx')]
print('uploads 残留docx:', leftovers)
