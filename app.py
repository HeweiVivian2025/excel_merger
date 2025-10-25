# app.py
import os
from flask import Flask, request, render_template, send_file, jsonify, redirect, url_for
from openpyxl import load_workbook
import shutil
from werkzeug.utils import secure_filename
from datetime import datetime
import uuid
import re

app = Flask(__name__)

# 配置
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx'}
MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50MB
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'dev-secret-key')

# 确保上传目录存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# 允许的文件类型
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def try_parse_date(date_str):
    if not isinstance(date_str, str):
        return None
    date_str = date_str.strip()
    formats = ["%Y/%m/%d", "%Y-%m-%d", "%m/%d/%Y", "%Y年%m月%d日"]
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt).date()
        except ValueError:
            continue
    return None

def safe_cell_value(ws, row, col):
    try:
        cell = ws.cell(row=row, column=col)
        if ws.merged_cells:
            for merged_range in ws.merged_cells.ranges:
                if cell.coordinate in merged_range:
                    return ws.cell(merged_range.min_row, merged_range.min_col).value
        value = cell.value
        if value is None or value == "":
            return None
        if isinstance(value, str):
            value = value.strip()
            if value in ["", "NONE", "#N/A", "NULL", "#VALUE!", "#REF!"]:
                return None
            parsed_date = try_parse_date(value)
            if parsed_date:
                return parsed_date
        return value
    except Exception:
        return None

def has_row_data(ws, row, cols):
    return any(safe_cell_value(ws, row, col) not in [None, ""] for col in cols)

def find_sheet_by_name(workbook, target_name):
    clean_target = re.sub(r'\s+', '', target_name.lower())
    for name in workbook.sheetnames:
        if re.sub(r'\s+', '', name.lower()) == clean_target:
            return workbook[name]
    # 模糊匹配
    for name in workbook.sheetnames:
        if target_name.lower() in name.lower():
            return workbook[name]
    return None

def process_excel_files(template_path, source_files, output_path, config):
    """
    config = {
        "sheet_names": ["Sheet1", "Data"],
        "read_cols": [5, 6, 7],       # 列索引列表（从1开始）
        "write_start_col": 5,
        "data_start_row": 6,
        "data_end_row": 50
    }
    """
    try:
        shutil.copy(template_path, output_path)
        wb_output = load_workbook(output_path)

        total_rows_written = 0

        for sheet_name in config["sheet_names"]:
            ws_template = find_sheet_by_name(wb_output, sheet_name)
            if not ws_template:
                print(f"未找到模板中的工作表: {sheet_name}")
                continue

            sheet_rows_written = 0
            for src_path in source_files:
                try:
                    wb_src = load_workbook(src_path, data_only=True)
                    ws_src = find_sheet_by_name(wb_src, sheet_name)
                    if not ws_src:
                        print(f"源文件中无此工作表: {sheet_name} in {src_path}")
                        wb_src.close()
                        continue

                    for row in range(config["data_start_row"], config["data_end_row"] + 1):
                        if not has_row_data(ws_src, row, config["read_cols"]):
                            continue

                        values = [safe_cell_value(ws_src, row, col) for col in config["read_cols"]]
                        for offset, val in enumerate(values):
                            if val in [None, ""]:
                                continue
                            write_col = config["write_start_col"] + offset
                            try:
                                ws_template.cell(row=row, column=write_col).value = val
                            except Exception as e:
                                print(f"写入失败 R{row}C{write_col}: {e}")

                        sheet_rows_written += 1
                    wb_src.close()
                except Exception as e:
                    print(f"处理源文件失败: {src_path}, 错误: {e}")

            total_rows_written += sheet_rows_written
            print(f"工作表 [{sheet_name}] 写入 {sheet_rows_written} 行")

        wb_output.save(output_path)
        return True, f"成功合并 {total_rows_written} 行数据"
    except Exception as e:
        import traceback
        traceback.print_exc()
        return False, f"处理失败: {str(e)}"

# ================== Web 路由 ==================

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    if 'template' not in request.files or 'sources' not in request.files:
        return jsonify({'success': False, 'message': '请上传模板和源文件'})

    template_file = request.files['template']
    source_files = request.files.getlist('sources')

    if template_file.filename == '' or not allowed_file(template_file.filename):
        return jsonify({'success': False, 'message': '请上传有效的模板文件 (.xlsx)'})

    if not source_files or all(f.filename == '' for f in source_files):
        return jsonify({'success': False, 'message': '请上传至少一个源文件'})

    valid_sources = [f for f in source_files if f.filename != '' and allowed_file(f.filename)]
    if not valid_sources:
        return jsonify({'success': False, 'message': '没有有效的源文件'})

    try:
        # 保存模板
        template_filename = secure_filename(template_file.filename)
        template_path = os.path.join(app.config['UPLOAD_FOLDER'], template_filename)
        template_file.save(template_path)

        # 保存源文件
        source_paths = []
        for f in valid_sources:
            src_filename = secure_filename(f.filename)
            src_path = os.path.join(app.config['UPLOAD_FOLDER'], src_filename)
            f.save(src_path)
            source_paths.append(src_path)

        # 读取用户配置（前端传入）
        config = {
            "sheet_names": [name.strip() for name in request.form.get('sheet_names', '').split(',') if name.strip()],
            "read_cols": [int(x) for x in request.form.get('read_cols', '5,6,7').split(',')],  # 默认 E,F,G
            "write_start_col": int(request.form.get('write_start_col', 5)),
            "data_start_row": int(request.form.get('data_start_row', 6)),
            "data_end_row": int(request.form.get('data_end_row', 50))
        }

        if not config["sheet_names"]:
            return jsonify({'success': False, 'message': '请填写工作表名称'})

        # 输出文件名
        base_name = os.path.splitext(template_filename)[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        random_suffix = str(uuid.uuid4())[:8]
        output_filename = f"{base_name}_merged_{timestamp}_{random_suffix}.xlsx"
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)

        # 处理合并
        success, message = process_excel_files(template_path, source_paths, output_path, config)

        if success:
            return jsonify({
                'success': True,
                'message': message,
                'download_url': url_for('download_file', filename=output_filename)
            })
        else:
            return jsonify({'success': False, 'message': message})

    except Exception as e:
        print(f"上传错误: {e}")
        return jsonify({'success': False, 'message': f'服务器错误: {str(e)}'})

@app.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    return "文件不存在", 404

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)