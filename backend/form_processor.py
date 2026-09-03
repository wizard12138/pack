"""
表单数据处理器
负责接收前端提交的扁平 JSON 数据，重建复杂对象。
"""

import json


def process_form_data(raw_data):
    """
    处理前端提交的表单数据
    当前主要做数据完整性校验，未来可扩展数据清洗、验证等

    参数:
        raw_data: dict 前端提交的原始数据

    返回:
        dict 处理后的数据
    """
    # 确保关键字段存在
    defaults = {
        'yinPersonCount': '1',
        'huoPersonCount': '1',
        'chaoPersonCount': '1',
        'huoTimeData': '{}',
    }
    for key, default_val in defaults.items():
        if key not in raw_data or not raw_data[key]:
            raw_data[key] = default_val

    # 确保 huoTimeData 是合法 JSON
    try:
        json.loads(raw_data.get('huoTimeData', '{}'))
    except (json.JSONDecodeError, TypeError):
        raw_data['huoTimeData'] = '{}'

    return raw_data


def validate_selected_forms(selected_forms):
    """
    验证选中的表单类型是否合法

    参数:
        selected_forms: list[str]

    返回:
        tuple (is_valid: bool, error_msg: str)
    """
    valid_forms = {'yingong', 'huo', 'chao', 'baoxiao'}
    for sf in selected_forms:
        if sf not in valid_forms:
            return False, f'无效的表单类型: {sf}'
    if not selected_forms:
        return False, '请至少选择一个表格'
    return True, ''