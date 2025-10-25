import os
import uuid
import logging
import atexit
import shutil
from flask import Flask, request, jsonify, send_file, render_template
from openpyxl import load_workbook
from openpyxl.styles import Protection

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# 初始化Flask应用并显式指定模板目录
template_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), 'templates'))
app = Flask(__name__, template_folder=template_dir)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB 限制

# 允许的文件扩展名
ALLOWED_EXTENSIONS = {'xlsx'}

def allowed_file(filename):
    """检查文件是否为允许的类型"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def cleanup_temp_files():
    """应用退出时清理临时文件"""
    temp_dir = '/tmp'
    if os.path.exists(temp_dir):
        for filename in os.listdir(temp_dir):
            if filename.startswith('merged_') and filename.endswith('.xlsx'):
                file_path = os.path.join(temp_dir, filename)
                try:
                    if os.path.isfile(file_path):
                        os.remove(file_path)
                        logger.info(f"清理临时文件: {filename}")
                except Exception as e:
                    logger.error(f"清理临时文件失败: {str(e)}")

# 注册退出时清理函数
atexit.register(cleanup_temp_files)

@app.route('/')
def index():
    """渲染主页面"""
    try:
        return render_template('index.html')
    except Exception as e:
        logger.error(f"渲染模板失败: {str(e)}")
        return "服务器错误", 500

@app.route('/upload', methods=['POST'])
def upload_files():
    """处理文件上传和合并"""
    # 验证请求中是否包含文件
    if 'template' not in request.files or 'sources' not in request.files:
        return jsonify({'error': '请求中缺少模板文件或源文件'}), 400
    
    template_file = request.files['template']
    source_files = request.files.getlist('sources')
    
    # 验证文件是否被选择
    if template_file.filename == '':
        return jsonify({'error': '未选择模板文件'}), 400
    if len(source_files) == 0 or all(f.filename == '' for f in source_files):
        return jsonify({'error': '未选择源文件'}), 400
    
    # 验证文件类型
    if not allowed_file(template_file.filename):
        return jsonify({'error': '模板文件必须是.xlsx格式'}), 400
    for f in source_files:
        if f.filename != '' and not allowed_file(f.filename):
            return jsonify({'error': f'源文件 {f.filename} 必须是.xlsx格式'}), 400
    
    # 生成输出文件名和路径
    output_filename = f"merged_{uuid.uuid4().hex}.xlsx"
    output_path = os.path.join('/tmp', output_filename)
    
    try:
        # 加载模板工作簿
        template_wb = load_workbook(template_file)
        template_ws = template_wb.active
        row_index = 2  # 从第2行开始写入数据
        
        # 处理每个源文件
        for source_file in source_files:
            if source_file.filename == '':
                continue
                
            try:
                source_wb = load_workbook(source_file)
                source_ws = source_wb.active
                
                # 读取源文件E1, F1, G1单元格的值
                e_value = source_ws['E1'].value if source_ws['E1'].value is not None else ''
                f_value = source_ws['F1'].value if source_ws['F1'].value is not None else ''
                g_value = source_ws['G1'].value if source_ws['G1'].value is not None else ''
                
                # 写入模板对应列
                template_ws[f'E{row_index}'] = e_value
                template_ws[f'F{row_index}'] = f_value
                template_ws[f'G{row_index}'] = g_value
                
                row_index += 1
                logger.info(f"成功处理源文件: {source_file.filename}")
                
            except Exception as e:
                logger.error(f"处理源文件 {source_file.filename} 失败: {str(e)}")
                return jsonify({'error': f'处理源文件 {source_file.filename} 时出错: {str(e)}'}), 500
        
        # 确保工作表未受保护
        if template_ws.protection.sheet:
            template_ws.protection = Protection(sheet=False)
        
        # 保存合并后的文件
        template_wb.save(output_path)
        logger.info(f"成功保存合并文件: {output_path}")
        
        # 验证文件是否存在
        if not os.path.exists(output_path):
            raise Exception("合并文件保存后不存在")
            
        # 返回文件供下载
        return send_file(
            output_path,
            as_attachment=True,
            download_name='merged_output.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        logger.error(f"合并文件处理失败: {str(e)}", exc_info=True)
        return jsonify({'error': f'处理文件时出错: {str(e)}'}), 500

# 错误处理路由
@app.errorhandler(500)
def internal_server_error(e):
    logger.error(f"服务器内部错误: {str(e)}")
    return jsonify({'error': '服务器内部错误，请稍后再试'}), 500

@app.errorhandler(400)
def bad_request(e):
    return jsonify({'error': '请求参数错误，请检查输入'}), 400

@app.errorhandler(404)
def not_found(e):
    return jsonify({'error': '请求的资源不存在'}), 404

if __name__ == '__main__':
    # 获取端口，Render会设置PORT环境变量
    port = int(os.environ.get('PORT', 5000))
    # 生产环境不启用debug
    app.run(host='0.0.0.0', port=port, debug=False)