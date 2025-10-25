import os
import uuid
import logging
import atexit
import shutil
from flask import Flask, request, jsonify, send_file, render_template
from openpyxl import load_workbook
from openpyxl.styles import Protection
from openpyxl.utils import get_column_letter

# 配置日志
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# 初始化Flask应用
template_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), 'templates'))
app = Flask(__name__, template_folder=template_dir)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB 限制

# 允许的文件扩展名
ALLOWED_EXTENSIONS = {'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_merged_cell_value(ws, cell):
    """获取合并单元格的值（增强版）"""
    try:
        # 检查是否为合并单元格
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                # 返回合并区域左上角单元格的值
                return ws[merged_range.start_cell.coordinate].value
        
        # 非合并单元格直接返回值
        return cell.value if cell.value is not None else ''
        
    except Exception as e:
        logger.error(f"获取单元格值失败: {str(e)}")
        return ''

def safe_cell_read(ws, cell_coord):
    """安全读取单元格值的封装函数"""
    try:
        cell = ws[cell_coord]
        # 检查是否为合并单元格
        if cell.data_type == 'm':  # merged cell
            return get_merged_cell_value(ws, cell)
        return cell.value if cell.value is not None else ''
    except Exception as e:
        logger.error(f"读取单元格 {cell_coord} 失败: {str(e)}")
        return ''

def cleanup_temp_files():
    """清理临时文件"""
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

atexit.register(cleanup_temp_files)

@app.route('/')
def index():
    try:
        return render_template('index.html')
    except Exception as e:
        logger.error(f"渲染模板失败: {str(e)}")
        return "服务器错误", 500

@app.route('/upload', methods=['POST'])
def upload_files():
    if 'template' not in request.files or 'sources' not in request.files:
        return jsonify({'error': '请求中缺少模板文件或源文件'}), 400
    
    template_file = request.files['template']
    source_files = request.files.getlist('sources')
    
    if template_file.filename == '':
        return jsonify({'error': '未选择模板文件'}), 400
    if len(source_files) == 0 or all(f.filename == '' for f in source_files):
        return jsonify({'error': '未选择源文件'}), 400
    
    if not allowed_file(template_file.filename):
        return jsonify({'error': '模板文件必须是.xlsx格式'}), 400
    for f in source_files:
        if f.filename != '' and not allowed_file(f.filename):
            return jsonify({'error': f'源文件 {f.filename} 必须是.xlsx格式'}), 400
    
    output_filename = f"merged_{uuid.uuid4().hex}.xlsx"
    output_path = os.path.join('/tmp', output_filename)
    
    try:
        # 加载模板工作簿，禁用只读模式
        template_wb = load_workbook(template_file, read_only=False, data_only=True)
        template_ws = template_wb.active
        row_index = 2
        
        for source_file in source_files:
            if source_file.filename == '':
                continue
                
            try:
                # 加载源文件，禁用只读模式
                source_wb = load_workbook(source_file, read_only=False, data_only=True)
                source_ws = source_wb.active
                
                # 使用安全读取函数读取单元格值
                e_value = safe_cell_read(source_ws, 'E1')
                f_value = safe_cell_read(source_ws, 'F1')
                g_value = safe_cell_read(source_ws, 'G1')
                
                logger.info(f"读取到的值 - E1: {e_value}, F1: {f_value}, G1: {g_value}")
                
                # 写入模板文件（使用值的副本而非引用）
                template_ws[f'E{row_index}'] = str(e_value) if e_value is not None else ''
                template_ws[f'F{row_index}'] = str(f_value) if f_value is not None else ''
                template_ws[f'G{row_index}'] = str(g_value) if g_value is not None else ''
                
                row_index += 1
                logger.info(f"成功处理源文件: {source_file.filename}")
                
            except Exception as e:
                logger.error(f"处理源文件 {source_file.filename} 失败: {str(e)}", exc_info=True)
                return jsonify({'error': f'处理源文件 {source_file.filename} 时出错: {str(e)}'}), 500
        
        # 确保工作表未受保护
        if template_ws.protection.sheet:
            template_ws.protection = Protection(sheet=False)
        
        # 保存合并后的文件
        template_wb.save(output_path)
        logger.info(f"成功保存合并文件: {output_path}")
        
        if not os.path.exists(output_path):
            raise Exception("合并文件保存后不存在")
            
        return send_file(
            output_path,
            as_attachment=True,
            download_name='merged_output.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        logger.error(f"合并文件处理失败: {str(e)}", exc_info=True)
        return jsonify({'error': f'处理文件时出错: {str(e)}'}), 500

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
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)