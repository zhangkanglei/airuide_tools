from flask import Flask, request, send_file, render_template
import os
import zipfile  # 新增：用于打包多个文件
import tempfile  # 新增：临时文件处理

app = Flask(__name__)
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
    }
    # 后续你加新工具，只需要在这里加一个新的字典就行，比如：
    # {
    #     "id": "score",
    #     "name": "成绩统计工具",
    #     "desc": "自动统计班级成绩，排名、平均分等",
    #     "route": "/score"
    # }
]

# 首页：工具列表+搜索
@app.route('/')
def index():
    # 搜索功能
    search_key = request.args.get('search', '').strip()
    # 过滤工具
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
        # 处理上传的文件
        old_file = request.files['old_file']
        # 核心修改：接收多个底板文件
        template_files = request.files.getlist('template_files')

        # 校验原始文件
        if old_file.filename == '':
            return "请上传原始表！", 400
        if not template_files:
            return "请至少上传一个部门底板表！", 400

        # 校验文件格式（只支持xlsx）
        if old_file.filename.endswith('.xls'):
            return "原始表请上传xlsx格式的文件，不支持xls格式，你可以用WPS/Office把文件另存为xlsx之后再上传。", 400
        for template_file in template_files:
            if template_file.filename.endswith('.xls'):
                return f"底板表【{template_file.filename}】请上传xlsx格式的文件，不支持xls格式！", 400

        # 保存原始表临时文件
        old_filename = old_file.filename
        old_path = os.path.join(UPLOAD_FOLDER, old_filename)
        old_file.save(old_path)

        # 临时存储所有生成的结果文件路径
        result_files = []
        try:
            # 循环处理每个底板表
            for template_file in template_files:
                template_filename = template_file.filename
                # 保存当前底板表临时文件
                template_path = os.path.join(UPLOAD_FOLDER, template_filename)
                template_file.save(template_path)

                # 调用处理函数生成对应部门的考勤表
                output_path, output_name = process_attendance(old_path, template_path)
                result_files.append((output_path, output_name))

                # 删除临时底板文件（可选，清理空间）
                os.remove(template_path)

            # 打包所有结果文件为ZIP
            # 创建临时ZIP文件
            zip_temp = tempfile.NamedTemporaryFile(delete=False, suffix='.zip', dir=UPLOAD_FOLDER)
            zip_path = zip_temp.name
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for file_path, file_name in result_files:
                    # 向ZIP中添加文件，使用自定义文件名
                    zipf.write(file_path, arcname=file_name)
                    # 删除单个结果文件（可选，清理空间）
                    os.remove(file_path)

            # 删除原始表临时文件（可选）
            os.remove(old_path)

            # 返回ZIP包供下载
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
                # 异常时清理临时文件
                for file_path, _ in result_files:
                    if os.path.exists(file_path):
                        os.remove(file_path)
                if os.path.exists(old_path):
                    os.remove(old_path)
                return f"处理出错了：{err_msg}", 500

    return render_template('attendance.html')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)