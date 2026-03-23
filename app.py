from flask import Flask, request, send_file, render_template
import os

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
        template_file = request.files['template_file']

        if old_file.filename == '' or template_file.filename == '':
            return "请上传原始表和底板表！", 400
        if old_file.filename.endswith('.xls') or template_file.filename.endswith('.xls'):
            return "请上传xlsx格式的文件，不支持xls格式，你可以用WPS/Office把文件另存为xlsx之后再上传。", 400

        # 保存临时文件
        old_path = os.path.join(UPLOAD_FOLDER, old_file.filename)
        old_file.save(old_path)

        template_path = os.path.join(UPLOAD_FOLDER, template_file.filename)
        template_file.save(template_path)

        # 如果是xls，转成xlsx


        try:
            # 调用考勤工具的处理函数
            output_path, output_name = process_attendance(old_path, template_path)
            # 返回文件
            return send_file(output_path, as_attachment=True, download_name=output_name)
        except Exception as e:
            err_msg = str(e)
            if "File contains no valid workbook part" in err_msg:
                # 针对改后缀的错误，给出提示
                return "文件格式不对，你不能直接修改文件的后缀，需要用WPS/Office打开文件，然后另存为xlsx格式之后再上传。", 400
            else:
                return f"处理出错了：{err_msg}", 500

    return render_template('attendance.html')

if __name__ == '__main__':
    app.run(debug=True)