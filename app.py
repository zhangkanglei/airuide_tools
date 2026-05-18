from flask import Flask, request, send_file, render_template
import os
import zipfile
import tempfile
import shutil  # 新增：用于清理临时目录
from datetime import datetime  # 新增：时间戳相关

app = Flask(__name__)
# 新增：设置最大上传大小为1GB（足够处理几百张照片）
app.config['MAX_CONTENT_LENGTH'] = 1024 * 1024 * 1024
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# 工具的配置列表，后续加新工具，只需要在这里加一项就行
TOOLS = [
    {
        "id": "attendance",
        "name": "考勤表自动转换",
        "desc": "自动处理考勤表，适配月份、自动转换状态、填充底板，支持半天/产假/旷工规则",
        "route": "/attendance"
    },
    {
        "id": "photo_rename",
        "name": "照片名称转换",
        "desc": "批量重命名照片文件，支持按序号/时间戳规则，自定义前缀，自动过滤非图片文件",
        "route": "/photo_rename"
    }
]

# 首页：工具列表+搜索
@app.route('/')
def index():
    search_key = request.args.get('search', '').strip()
    filtered_tools = []
    for tool in TOOLS:
        if search_key == '' or search_key in tool['name'] or search_key in tool['desc']:
            filtered_tools.append(tool)
    return render_template('index.html', tools=filtered_tools, search_key=search_key)

# 考勤表转换工具的路由
from tools.attendance import process_attendance
@app.route('/attendance', methods=['GET', 'POST'])
def attendance():
    if request.method == 'POST':
        old_file = request.files['old_file']
        template_files = request.files.getlist('template_files')

        if old_file.filename == '':
            return "请上传原始表！", 400
        if not template_files:
            return "请至少上传一个部门底板表！", 400

        if old_file.filename.endswith('.xls'):
            return "原始表请上传xlsx格式的文件，不支持xls格式，你可以用WPS/Office把文件另存为xlsx之后再上传。", 400
        for template_file in template_files:
            if template_file.filename.endswith('.xls'):
                return f"底板表【{template_file.filename}】请上传xlsx格式的文件，不支持xls格式！", 400

        old_filename = old_file.filename
        old_path = os.path.join(UPLOAD_FOLDER, old_filename)
        old_file.save(old_path)

        result_files = []
        try:
            for template_file in template_files:
                template_filename = template_file.filename
                template_path = os.path.join(UPLOAD_FOLDER, template_filename)
                template_file.save(template_path)

                output_path, output_name = process_attendance(old_path, template_path)
                result_files.append((output_path, output_name))

                os.remove(template_path)

            zip_temp = tempfile.NamedTemporaryFile(delete=False, suffix='.zip', dir=UPLOAD_FOLDER)
            zip_path = zip_temp.name
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for file_path, file_name in result_files:
                    zipf.write(file_path, arcname=file_name)
                    os.remove(file_path)

            os.remove(old_path)

            return send_file(
                zip_path,
                as_attachment=True,
                download_name=f"各部门考勤表_{old_filename.replace('.xlsx', '')}.zip"
            )

        except Exception as e:
            err_msg = str(e)
            if "File contains no valid workbook part" in err_msg:
                return "文件格式不对，你不能直接修改文件的后缀，需要用WPS/Office打开文件，然后另存为xlsx格式之后再上传。", 400
            else:
                for file_path, _ in result_files:
                    if os.path.exists(file_path):
                        os.remove(file_path)
                if os.path.exists(old_path):
                    os.remove(old_path)
                return f"处理出错了：{err_msg}", 500

    return render_template('attendance.html')

from tools.photo_rename import process_photo_rename
@app.route('/photo_rename', methods=['GET', 'POST'])
def photo_rename():
    if request.method == 'POST':
        # 获取上传的文件
        excel_file = request.files['excel_file']
        photo_files = request.files.getlist('photo_files')

        # 校验文件
        if excel_file.filename == '':
            return "请上传包含学生信息的Excel文件！", 400
        if not photo_files:
            return "请至少上传一张学生照片！", 400

        # 校验文件格式
        if excel_file.filename.endswith('.xls'):
            return "Excel文件请上传xlsx格式，不支持xls格式，请用WPS/Office另存为xlsx后再上传。", 400

        for photo_file in photo_files:
            if not photo_file.filename.lower().endswith('.jpg'):
                return f"照片【{photo_file.filename}】格式不对，只支持jpg格式的照片！", 400

        # 保存Excel临时文件
        excel_filename = excel_file.filename
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
        excel_file.save(excel_path)

        try:
            # 调用处理函数
            zip_path, zip_filename = process_photo_rename(excel_path, photo_files)

            # 删除Excel临时文件
            os.remove(excel_path)

            # 返回ZIP包下载
            return send_file(
                zip_path,
                as_attachment=True,
                download_name=zip_filename
            )

        except Exception as e:
            err_msg = str(e)
            # 异常时清理临时文件
            if os.path.exists(excel_path):
                os.remove(excel_path)
            if "File contains no valid workbook part" in err_msg:
                return "Excel文件格式损坏，不能直接修改后缀名，请用WPS/Office打开后另存为xlsx格式。", 400
            else:
                return f"处理出错了：{err_msg}", 500

    # GET请求返回前端页面
    return render_template('photo_rename.html')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)