# tools/photo_rename.py
import os
import zipfile
import tempfile
from openpyxl import load_workbook

def process_photo_rename(excel_path, photo_files):
    """
    根据Excel中的身份证件号匹配照片，重命名为"姓名+学籍号.jpg"
    自动通过表头名称识别列，不依赖固定列位置
    :param excel_path: 上传的Excel文件路径
    :param photo_files: 上传的照片文件列表
    :return: 打包后的zip文件路径、zip文件名
    """
    # 1. 读取Excel，自动识别表头列
    wb = load_workbook(excel_path)
    ws = wb.active
    id_map = {}

    # 第一步：遍历第一行，找到表头对应的列索引
    header_row = 1
    name_col = None
    id_card_col = None
    student_id_col = None

    for col in range(1, ws.max_column + 1):
        header_value = str(ws.cell(row=header_row, column=col).value).strip() if ws.cell(row=header_row, column=col).value else ""
        if header_value == "姓名":
            name_col = col
        elif header_value == "身份证件号":
            id_card_col = col
        elif header_value == "学籍号":
            student_id_col = col

    # 检查是否找到所有必要的表头
    missing_headers = []
    if not name_col:
        missing_headers.append("姓名")
    if not id_card_col:
        missing_headers.append("身份证件号")
    if not student_id_col:
        missing_headers.append("学籍号")

    if missing_headers:
        raise Exception(f"Excel表头缺少必要字段：{', '.join(missing_headers)}。请确保第一行包含'姓名'、'身份证件号'、'学籍号'这三个表头。")

    # 第二步：从第二行开始读取数据，建立身份证件号->(姓名, 学籍号)的映射
    for row in range(header_row + 1, ws.max_row + 1):
        name = ws.cell(row=row, column=name_col).value
        id_card = str(ws.cell(row=row, column=id_card_col).value).strip() if ws.cell(row=row, column=id_card_col).value else ""
        student_id = str(ws.cell(row=row, column=student_id_col).value).strip() if ws.cell(row=row, column=student_id_col).value else ""

        if id_card and name and student_id:
            # 统一处理身份证件号中的X为大写
            id_map[id_card.upper()] = (name, student_id)

    if not id_map:
        raise Exception("Excel中没有读取到有效的学生数据，请检查数据是否从第二行开始，且身份证件号、姓名、学籍号都不为空。")

    # 2. 创建临时目录存储重命名后的照片
    temp_dir = tempfile.mkdtemp(dir='uploads')
    result_files = []
    fail_list = []

    # 3. 处理每张照片
    for photo_file in photo_files:
        filename = photo_file.filename
        # 只处理jpg文件
        if filename.lower().endswith(".jpg"):
            # 提取身份证件号（去掉.jpg后缀）
            id_card = os.path.splitext(filename)[0].strip().upper()

            if id_card in id_map:
                name, student_id = id_map[id_card]
                # 构造新文件名：姓名+学籍号.jpg
                new_filename = f"{name}+{student_id}.jpg"
                new_path = os.path.join(temp_dir, new_filename)

                # 防止重名
                counter = 1
                while os.path.exists(new_path):
                    new_filename = f"{name}+{student_id}_{counter}.jpg"
                    new_path = os.path.join(temp_dir, new_filename)
                    counter += 1

                # 保存重命名后的文件
                photo_file.save(new_path)
                result_files.append((new_path, new_filename))
            else:
                fail_list.append(filename)

    # 4. 打包所有成功重命名的文件为ZIP
    zip_temp = tempfile.NamedTemporaryFile(delete=False, suffix='.zip', dir='uploads')
    zip_path = zip_temp.name

    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for file_path, file_name in result_files:
            zipf.write(file_path, arcname=file_name)
            os.remove(file_path)

        # 如果有失败的文件，生成一个失败列表txt
        if fail_list:
            fail_txt_path = os.path.join(temp_dir, "匹配失败列表.txt")
            with open(fail_txt_path, 'w', encoding='utf-8') as f:
                f.write("以下照片未在Excel中找到对应的身份证件号：\n\n")
                for filename in fail_list:
                    f.write(f"- {filename}\n")
                f.write(f"\n\n共成功匹配：{len(result_files)} 张照片\n")
                f.write(f"匹配失败：{len(fail_list)} 张照片\n")
            zipf.write(fail_txt_path, arcname="匹配失败列表.txt")
            os.remove(fail_txt_path)

    # 5. 清理临时目录
    os.rmdir(temp_dir)

    return zip_path, f"学生照片重命名结果_{len(result_files)}张.zip"