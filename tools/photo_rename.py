import os
import zipfile
import tempfile
import shutil
import logging
import time
from openpyxl import load_workbook

def process_photo_rename(excel_path, photo_files, upload_folder='uploads'):
    """
    根据Excel中的身份证件号匹配照片，重命名为"姓名+学籍号.jpg"
    自动通过表头名称识别列，不依赖固定列位置
    :param excel_path: 上传的Excel文件路径
    :param photo_files: 上传的照片文件列表（Flask文件对象）
    :param upload_folder: 临时目录基路径
    :return: 打包后的zip文件路径、zip文件名
    """
    logging.info("========== 开始处理照片重命名 ==========")
    logging.info(f"Excel文件: {excel_path}")
    logging.info(f"照片数量: {len(photo_files)}")

    # 1. 读取Excel，自动识别表头列
    t0 = time.time()
    wb = load_workbook(excel_path)
    ws = wb.active
    logging.info(f"Excel加载耗时: {time.time()-t0:.2f}s")

    id_map = {}  # 身份证件号 -> (姓名, 学籍号)

    # 第一步：遍历第一行，找到表头对应的列索引
    header_row = 1
    name_col = None
    id_card_col = None
    student_id_col = None

    for col in range(1, ws.max_column + 1):
        cell_value = ws.cell(row=header_row, column=col).value
        header_value = str(cell_value).strip() if cell_value else ""
        if header_value == "姓名":
            name_col = col
        elif header_value == "身份证件号":
            id_card_col = col
        elif header_value == "学籍号":
            student_id_col = col

    missing_headers = []
    if not name_col:
        missing_headers.append("姓名")
    if not id_card_col:
        missing_headers.append("身份证件号")
    if not student_id_col:
        missing_headers.append("学籍号")

    if missing_headers:
        raise Exception(f"Excel表头缺少必要字段：{', '.join(missing_headers)}。请确保第一行包含'姓名'、'身份证件号'、'学籍号'这三个表头。")

    # 第二步：从第二行开始读取数据，建立映射
    for row in range(header_row + 1, ws.max_row + 1):
        name = ws.cell(row=row, column=name_col).value
        id_card_raw = ws.cell(row=row, column=id_card_col).value
        student_id_raw = ws.cell(row=row, column=student_id_col).value

        # 处理空值和空白字符
        name = str(name).strip() if name else ""
        id_card = str(id_card_raw).strip().upper() if id_card_raw else ""
        student_id = str(student_id_raw).strip() if student_id_raw else ""

        if id_card and name and student_id:
            id_map[id_card] = (name, student_id)

    if not id_map:
        raise Exception("Excel中没有读取到有效的学生数据，请检查数据是否从第二行开始，且身份证件号、姓名、学籍号都不为空。")

    logging.info(f"从Excel读取到 {len(id_map)} 条有效学生记录")

    # 2. 创建临时目录存储重命名后的照片
    temp_dir = tempfile.mkdtemp(dir=upload_folder)
    logging.info(f"创建临时目录: {temp_dir}")

    result_files = []  # (完整路径, 文件名)
    fail_list = []

    # 3. 处理每张照片
    for idx, photo_file in enumerate(photo_files, 1):
        filename = photo_file.filename
        logging.info(f"处理第 {idx}/{len(photo_files)} 张: {filename}")

        if not filename.lower().endswith(".jpg"):
            logging.warning(f"跳过非jpg文件: {filename}")
            continue

        # 提取身份证件号（去掉.jpg后缀）
        id_card = os.path.splitext(filename)[0].strip().upper()
        logging.debug(f"提取身份证号: {id_card}")

        if id_card in id_map:
            name, student_id = id_map[id_card]
            new_filename = f"{name}+{student_id}.jpg"
            new_path = os.path.join(temp_dir, new_filename)

            # 防止重名（极少发生）
            counter = 1
            while os.path.exists(new_path):
                new_filename = f"{name}+{student_id}_{counter}.jpg"
                new_path = os.path.join(temp_dir, new_filename)
                counter += 1
                if counter > 100:  # 安全保护
                    logging.error(f"文件名冲突过多，放弃: {new_filename}")
                    break

            # 保存文件（可能很耗时，记录耗时）
            t_save = time.time()
            photo_file.save(new_path)
            save_time = time.time() - t_save
            logging.info(f"保存 {new_filename} 耗时 {save_time:.2f}s")

            result_files.append((new_path, new_filename))
        else:
            logging.warning(f"未匹配到身份证号: {id_card} (来自文件 {filename})")
            fail_list.append(filename)

    # 4. 打包所有成功重命名的文件为ZIP
    zip_temp = tempfile.NamedTemporaryFile(delete=False, suffix='.zip', dir=upload_folder)
    zip_path = zip_temp.name
    logging.info(f"创建ZIP文件: {zip_path}")

    t_zip = time.time()
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for file_path, file_name in result_files:
            zipf.write(file_path, arcname=file_name)
            os.remove(file_path)  # 加入ZIP后立即删除原文件

        if fail_list:
            fail_txt_path = os.path.join(temp_dir, "匹配失败列表.txt")
            with open(fail_txt_path, 'w', encoding='utf-8') as f:
                f.write("以下照片未在Excel中找到对应的身份证件号：\n\n")
                for fname in fail_list:
                    f.write(f"- {fname}\n")
                f.write(f"\n\n共成功匹配：{len(result_files)} 张照片\n")
                f.write(f"匹配失败：{len(fail_list)} 张照片\n")
            zipf.write(fail_txt_path, arcname="匹配失败列表.txt")
            os.remove(fail_txt_path)
    logging.info(f"ZIP打包完成，耗时 {time.time()-t_zip:.2f}s")

    # 5. 清理临时目录
    try:
        shutil.rmtree(temp_dir, ignore_errors=True)
        logging.info(f"已删除临时目录: {temp_dir}")
    except Exception as e:
        logging.warning(f"删除临时目录失败: {e}")

    zip_filename = f"学生照片重命名结果_{len(result_files)}张.zip"
    logging.info(f"处理完成，成功 {len(result_files)} 张，失败 {len(fail_list)} 张")
    logging.info("========== 处理结束 ==========")

    return zip_path, zip_filename