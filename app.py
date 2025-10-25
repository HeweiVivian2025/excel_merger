# app.py
from flask import Flask, request, jsonify, send_file, render_template
import os
import uuid
import openpyxl
from openpyxl.styles import Protection

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB 限制


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_files():
    if 'template' not in request.files or 'sources' not in request.files:
        return jsonify({'error': '缺少文件'}), 400

    template_file = request.files['template']
    source_files = request.files.getlist('sources')

    if template_file.filename == '' or len(source_files) == 0:
        return jsonify({'error': '未选择文件'}), 400

    # 使用 /tmp 目录（Render 唯一可写路径）
    output_filename = f"merged_{uuid.uuid4().hex}.xlsx"
    output_path = os.path.join('/tmp', output_filename)

    try:
        # 加载模板工作簿
        template_wb = openpyxl.load_workbook(template_file)
        template_ws = template_wb.active

        row_index = 2  # 假设从第2行开始写入数据

        for source_file in source_files:
            if source_file.filename == '':
                continue

            source_wb = openpyxl.load_workbook(source_file)
            source_ws = source_wb.active

            # 读取源文件 E1, F1, G1
            e_value = source_ws['E1'].value
            f_value = source_ws['F1'].value
            g_value = source_ws['G1'].value

            # 写入模板对应列
            template_ws[f'E{row_index}'] = e_value
            template_ws[f'F{row_index}'] = f_value
            template_ws[f'G{row_index}'] = g_value

            row_index += 1

        # 确保工作表未受保护（避免保存失败）
        if template_ws.protection.sheet:
            template_ws.protection = Protection(sheet=False)

        # 保存到 /tmp
        template_wb.save(output_path)

        # 调试日志
        print(f"✅ 成功保存合并文件: {output_path}")
        print(f"📁 文件是否存在？{os.path.exists(output_path)}")

        # 返回文件用于下载
        return send_file(
            output_path,
            as_attachment=True,
            download_name='merged_output.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        print(f"❌ 合并失败: {str(e)}")
        return jsonify({'error': f'处理失败: {str(e)}'}), 500


if __name__ == '__main__':
    # Render 要求绑定到 0.0.0.0:$PORT
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)