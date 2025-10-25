import os
import tempfile
import shutil
from flask import Flask, request, jsonify, render_template, send_file, abort
from openpyxl import load_workbook
from datetime import datetime

app = Flask(__name__)

# 配置上传目录
app.config['UPLOAD_FOLDER'] = '/tmp/excel_merger_uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def find_sheet_by_name(wb, sheet_name):
    for sheet in wb.sheetnames:
        if sheet == sheet_name:
            return wb[sheet]
    return None

def safe_cell_value(ws, row, col):
    cell = ws.cell(row=row, column=col)
    if cell.value is None:
        return ""
    return str(cell.value)

def has_row_data(ws, row, cols):
    for col in cols:
        cell = ws.cell(row=row, column=col)
        if cell.value not in [None, "", "#DIV/0!"]:
            return True
    return False

def process_excel_files(template_path, source_files, output_path, config):
    try:
        wb_output = load_workbook(template_path)
        total_rows_written = 0

        for sheet_name in config["sheet_names"]:
            ws_template = find_sheet_by_name(wb_output, sheet_name)
            if not ws_template:
                print(f"⚠️ 未找到模板中的工作表: {sheet_name}")
                continue

            # 解除保护和锁定
            if ws_template.protection.sheet:
                try:
                    ws_template.protection.disable()
                except:
                    pass

            for row in ws_template.iter_rows():
                for cell in row:
                    cell.protection = None

            for src_path in source_files:
                try:
                    wb_src = load_workbook(src_path, data_only=True)
                    ws_src = find_sheet_by_name(wb_src, sheet_name)
                    if not ws_src:
                        print(f"⚠️ 源文件中无此工作表: {sheet_name} in {src_path}")
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
                                target_cell = ws_template.cell(row=row, column=write_col)
                                target_cell.value = val
                                target_cell.protection = None
                                print(f"📝 写入 R{row}C{write_col} = {val}")
                            except Exception as e:
                                print(f"❌ 写入失败 R{row}C{write_col}: {e}")

                        total_rows_written += 1
                    wb_src.close()
                except Exception as e:
                    print(f"❌ 处理源文件失败: {src_path}, 错误: {e}")

        wb_output.save(output_path)
        print(f"✅ 成功保存合并文件: {output_path}")
        return True, f"成功合并 {total_rows_written} 行数据"

    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"❌ 处理失败: {str(e)}")
        return False, f"处理失败: {str(e)}"

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    try:
        template_file = request.files.get('template')
        source_files = request.files.getlist('sources')
        sheet_names = request.form.get('sheet_names').split(',')
        read_cols = list(map(int, request.form.get('read_cols').split(',')))
        write_start_col = int(request.form.get('write_start_col'))
        data_start_row = int(request.form.get('data_start_row'))
        data_end_row = int(request.form.get('data_end_row'))

        if not template_file or len(source_files) == 0:
            return jsonify({'success': False, 'message': '请上传模板和数据源文件'})

        # 保存模板
        template_path = os.path.join(app.config['UPLOAD_FOLDER'], template_file.filename)
        template_file.save(template_path)

        # 保存源文件
        source_paths = []
        for src_file in source_files:
            src_path = os.path.join(app.config['UPLOAD_FOLDER'], src_file.filename)
            src_file.save(src_path)
            source_paths.append(src_path)

        # 生成输出文件名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        output_filename = f"{template_file.filename.split('.')[0]}_merged_{timestamp}.xlsx"
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)

        # 处理合并
        config = {
            "sheet_names": sheet_names,
            "read_cols": read_cols,
            "write_start_col": write_start_col,
            "data_start_row": data_start_row,
            "data_end_row": data_end_row
        }

        success, message = process_excel_files(template_path, source_paths, output_path, config)

        if success:
            # ✅ 使用 copy 确保文件存在
            try:
                shutil.copy(output_path, output_path)  # 确保文件可读
                print(f"✅ 输出文件已保存: {output_path}")
            except Exception as e:
                print(f"❌ 复制失败: {e}")

            download_url = request.url_root.rstrip('/') + '/download/' + output_filename
            return jsonify({
                'success': True,
                'message': message,
                'download_url': download_url
            })
        else:
            return jsonify({'success': False, 'message': message})

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'message': f'上传失败: {str(e)}'})

@app.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if os.path.exists(file_path):
        return send_file(
            file_path,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    else:
        abort(404)

if __name__ == '__main__':
    app.run(debug=True)