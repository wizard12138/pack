"""
Flask 服务器入口
负责路由注册、静态文件服务、API 接口。
启动时自动打开浏览器。
"""

import os
import sys
import json
import traceback
import webbrowser
from threading import Timer

from flask import Flask, request, send_file, jsonify
import io

from excel_handler import generate_excel


def _base_dir():
    """定位资源根目录：兼容 PyInstaller 冻结（onedir / onefile）与开发模式。"""
    if getattr(sys, 'frozen', False):
        # one-file: 解包到临时目录 sys._MEIPASS；onedir: 可执行文件所在目录
        return getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


# 资源根目录（开发模式为 backend 目录；冻结后为打包目录）
BACKEND_DIR = _base_dir()
TEMPLATE_PATH = os.path.join(BACKEND_DIR, 'templates', '差旅费报销相关表格.xlsx')
STATIC_DIR = os.path.join(BACKEND_DIR, 'static')

app = Flask(__name__, static_folder=STATIC_DIR, static_url_path='')


@app.route('/')
def index():
    """返回首页"""
    return app.send_static_file('index.html')


@app.route('/api/export', methods=['POST'])
def export_excel():
    """
    导出 Excel 文件（ZIP 或单个 xlsx）

    接收 JSON 格式请求体：
    {
        "selectedForms": ["yingong", "huo", "baoxiao"],
        "formData": { ... }  // 所有表单字段值
    }

    返回: Excel 文件或 ZIP 压缩包
    """
    try:
        req_data = request.get_json(force=True)
        if not req_data:
            return jsonify({'error': '请求体为空'}), 400

        selected_forms = req_data.get('selectedForms', [])
        form_data = req_data.get('formData', {})

        if not selected_forms:
            return jsonify({'error': '请至少选择一个表格'}), 400

        # 检查模板文件是否存在
        if not os.path.exists(TEMPLATE_PATH):
            return jsonify({'error': f'模板文件不存在: {TEMPLATE_PATH}'}), 500

        # 生成 Excel
        file_bytes, file_name = generate_excel(selected_forms, form_data, TEMPLATE_PATH)

        # 判断文件类型
        if file_name.endswith('.zip'):
            mimetype = 'application/zip'
        else:
            mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

        return send_file(
            io.BytesIO(file_bytes),
            mimetype=mimetype,
            as_attachment=True,
            download_name=file_name
        )

    except Exception as e:
        traceback.print_exc()  # 打印详细错误堆栈到控制台
        return jsonify({'error': f'导出失败: {str(e)}'}), 500


@app.route('/api/health', methods=['GET'])
def health_check():
    """健康检查接口"""
    return jsonify({'status': 'ok', 'template_exists': os.path.exists(TEMPLATE_PATH)})


def open_browser():
    """延迟打开浏览器"""
    webbrowser.open('http://localhost:5000')


def create_app():
    """创建 Flask 应用实例"""
    # 确保 static 目录存在
    os.makedirs(STATIC_DIR, exist_ok=True)
    os.makedirs(os.path.join(BACKEND_DIR, 'templates'), exist_ok=True)
    return app


if __name__ == '__main__':
    create_app()
    # 1.5 秒后打开浏览器
    Timer(1.5, open_browser).start()
    print('=' * 60)
    print('  财务报销系统  - 后端服务已启动')
    print('  访问地址: http://localhost:5000')
    print('  按 Ctrl+C 停止服务')
    print('=' * 60)
    app.run(host='127.0.0.1', port=5000, debug=False)