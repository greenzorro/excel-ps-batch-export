'''
File: psd_renderer.py
Project: excel-ps-batch-export
Created: 2024-09-25 02:07:52
Author: Victor Cheng
Email: hi@victor42.work
Description: PSD渲染器 - 将Excel数据渲染到PSD模板并导出图片
'''

import os
import pandas as pd
from psd_tools import PSDImage
from PIL import Image, ImageDraw, ImageFont
import textwrap
import sys
import csv
from datetime import datetime
import re
from typing import Set, List, Tuple, Dict
from tqdm import tqdm

# 全局变量存储错误和警告
validation_errors = []
validation_warnings = []

def safe_print_message(message):
    """安全打印消息，处理Windows控制台编码问题
    
    :param str message: 要打印的消息
    """
    try:
        print(message)
    except UnicodeEncodeError:
        # 如果直接打印失败，使用安全的编码方式
        safe_message = message.encode('ascii', errors='replace').decode('ascii')
        print(safe_message)

# 作为模块导入时的默认值
file_name = None
font_file = None
image_format = None
quality = 95
optimize = False
current_datetime = ''
output_path = 'export'
excel_file_path = None
psd_file_path = None
text_font = None

def read_excel_file(file_path):
    """读取Excel文件

    :param str file_path: Excel文件路径
    :return pd.DataFrame: 包含Excel数据的DataFrame
    :raises FileNotFoundError: 当文件不存在时抛出异常
    :raises ValueError: 当文件格式不支持时抛出异常
    """
    if not os.path.exists(file_path):
        # 处理Windows路径编码问题
        # 使用ASCII编码确保路径在Windows控制台正确显示
        try:
            # 尝试直接显示路径
            error_msg = f"Excel文件不存在: {file_path}"
        except UnicodeEncodeError:
            # 如果编码失败，使用安全的显示方式
            import traceback
            error_msg = f"Excel文件不存在: {file_path.encode('ascii', errors='replace').decode('ascii')}"
        raise FileNotFoundError(error_msg)
    
    if not file_path.lower().endswith(('.xlsx', '.xls')):
        raise ValueError(f"不支持的文件格式: {file_path}")
    
    try:
        df = pd.read_excel(file_path, sheet_name=0)
        return df
    except Exception as e:
        raise ValueError(f"读取Excel文件失败: {file_path}, 错误: {str(e)}")

def set_layer_visibility(layer, visibility):
    """设置图层可见性

    :param PSDLayer layer: PSD图层对象
    :param bool visibility: 是否可见
    :raises TypeError: 当visibility不是布尔值时抛出异常
    """
    # 处理numpy布尔类型
    if hasattr(visibility, 'item'):
        visibility = visibility.item()
    # 处理pandas布尔类型
    if hasattr(visibility, 'bool'):
        visibility = visibility.bool()
    
    # 正确解析布尔值
    if isinstance(visibility, bool):
        layer.visible = visibility
    elif isinstance(visibility, str):
        # 处理字符串形式的布尔值
        visibility_lower = visibility.lower().strip()
        
        # 空字符串或只有空格的字符串为False
        if not visibility_lower:
            layer.visible = False
        elif visibility_lower in ('true', '1', 'yes', 'on', 't', 'y'):
            layer.visible = True
        elif visibility_lower in ('false', '0', 'no', 'off', 'f', 'n'):
            layer.visible = False
        else:
            # 尝试解析为数字
            try:
                # 尝试转换为浮点数
                num_value = float(visibility_lower)
                layer.visible = num_value != 0
            except ValueError:
                # 无法转换为数字，按照Python的bool()转换规则
                layer.visible = bool(visibility)
    elif isinstance(visibility, (int, float)):
        # 处理数字形式的布尔值
        layer.visible = visibility != 0
    else:
        # 其他类型，使用Python的bool()转换
        layer.visible = bool(visibility)

def get_font_color(font_info):
    """获取文字颜色

    :param dict font_info: 字体信息字典
    :return tuple: 文字颜色 (r, g, b, a)
    """
    if 'FillColor' in font_info['StyleRun']['RunArray'][0]['StyleSheet']['StyleSheetData']:
        argb_color = font_info['StyleRun']['RunArray'][0]['StyleSheet']['StyleSheetData']['FillColor']['Values']
        r = argb_color[1]
        g = argb_color[2]
        b = argb_color[3]
        a = argb_color[0]
        font_color = (r, g, b, a)
        font_color = tuple(int(c * 255) for c in font_color)  # 确保颜色值为整数
    else:
        # 如果没有 'FillColor'，使用默认颜色
        font_color = (0, 0, 0, 255)  # 默认黑色
    return font_color

def calculate_text_position(text, layer_width, font_size, alignment):
    """计算单行文字位置

    :param str text: 文字内容
    :param int layer_width: 图层宽度
    :param int font_size: 字体大小
    :param str alignment: 对齐方式 ('left', 'center', 'right')
    :return tuple: 文字位置 (x, y)
    :raises ValueError: 当参数无效时抛出异常
    """
    # 参数验证
    if font_size <= 0:
        raise ValueError(f"字体大小必须大于0，当前值: {font_size}")
    
    if layer_width < 0:
        raise ValueError(f"图层宽度不能为负数，当前值: {layer_width}")
    
    if alignment not in ['left', 'center', 'right']:
        raise ValueError(f"对齐方式必须是 'left', 'center', 或 'right'，当前值: {alignment}")
    
    # 使用Pillow的字体度量功能获取精确的文本宽度
    try:
        from PIL import Image, ImageDraw, ImageFont
        
        # 创建一个临时图像来计算文本宽度
        temp_image = Image.new('RGB', (1, 1))
        draw = ImageDraw.Draw(temp_image)
        
        # 尝试加载字体，如果失败则使用默认字体
        try:
            # 使用相对路径的默认字体
            script_dir = os.path.dirname(os.path.abspath(__file__))
            font_path = os.path.join(script_dir, 'assets', 'fonts', 'AlibabaPuHuiTi-2-85-Bold.ttf')
            
            if os.path.exists(font_path):
                font = ImageFont.truetype(font_path, font_size)
            else:
                # 如果默认字体不存在，使用Pillow的默认字体
                font = ImageFont.load_default()
        except (OSError, IOError):
            # 字体加载失败，使用默认字体
            font = ImageFont.load_default()
        
        # 使用textbbox获取精确的文本边界框
        bbox = draw.textbbox((0, 0), text, font=font)
        text_width = bbox[2] - bbox[0]  # 右边界 - 左边界
        
    except ImportError:
        # 如果PIL不可用，回退到原始算法（改进版）
        text_width = 0
        for char in text:
            if '\u4e00' <= char <= '\u9fff':  # 判断是否为中文字符
                text_width += font_size * 0.9  # 更准确的中文字符宽度估算
            elif char.isdigit():
                text_width += font_size * 0.6  # 数字字符宽度
            elif char in 'iIl1':  # 窄字符
                text_width += font_size * 0.3
            elif char in 'mwMW':  # 宽字符
                text_width += font_size * 0.8
            else:
                text_width += font_size * 0.5  # 普通英文字符
    
    # 计算位置
    if alignment == 'center':  # 计算居中位置
        x_position = (layer_width - text_width) / 2
    elif alignment == 'right':  # 计算右对齐位置
        x_position = layer_width - text_width
    else:  # 计算左对齐位置
        x_position = 0
    
    # 修正文字位置偏移（基于实际测试调整）
    x_offset = font_size * 0.01
    y_offset = font_size * 0.26
    return x_position - x_offset, -y_offset

def preprocess_text(text_content):
    """预处理文本内容，在写入图片前进行统一清理

    :param str text_content: 原始文本内容
    :return str: 预处理后的文本内容
    """
    if text_content is None:
        return ""

    # 转换为字符串（处理非字符串类型）
    text_content = str(text_content)

    # 清理Excel特殊字符转义字符串
    text_content = text_content.replace('_x000D_', '')  # 回车符
    text_content = text_content.replace('_x000A_', '')  # 换行符
    text_content = text_content.replace('_x0009_', '')  # 制表符

    # 清理首尾成对的英文引号
    if len(text_content) >= 2 and text_content.startswith('"') and text_content.endswith('"'):
        text_content = text_content[1:-1]

    # 清理首尾空白字符（空格、制表符、换行符等）
    text_content = text_content.strip()

    # 文本内容规范化处理
    # 将中文双引号"" (U+201C, U+201D) 替换为中文书名号「」(U+300C, U+300D)
    text_content = text_content.replace(chr(0x201C), '「')  # 左双引号"
    text_content = text_content.replace(chr(0x201D), '」')  # 右双引号"
    # 将斜杠替换为和号
    text_content = text_content.replace('/', '&')

    return text_content

def update_text_layer(layer, text_content, pil_image, text_font='assets/fonts/AlibabaPuHuiTi-2-85-Bold.ttf'):
    """更新文字图层内容

    :param PSDLayer layer: PSD文字图层
    :param str text_content: 新的文字内容
    :param PIL.Image pil_image: PIL图像对象
    :param str text_font: 字体文件路径
    """
    import os
    # 预处理文本内容，统一清理空白字符
    text_content = preprocess_text(text_content)

    # 确保字体路径相对于脚本所在目录
    if not os.path.isabs(text_font):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        text_font = os.path.join(script_dir, text_font)

    layer.visible = False  # 防止PSD原始图层被输出到PIL
    font_info = layer.engine_dict
    font_size = font_info['StyleRun']['RunArray'][0]['StyleSheet']['StyleSheetData']['FontSize']
    font_color = get_font_color(font_info)
    font = ImageFont.truetype(text_font, int(font_size))
    draw = ImageDraw.Draw(pil_image)
    layer_width = layer.size[0]
    # 判断对齐方向
    alignment = 'left'
    if '_c' in layer.name:
        alignment = 'center'
    elif '_r' in layer.name:
        alignment = 'right'
    if '_p' in layer.name:
        # 段落文本处理
        if any('\u4e00' <= char <= '\u9fff' for char in text_content):
            wrapped_text = textwrap.fill(text_content, width=round(layer_width / font_size))
        else:
            wrapped_text = textwrap.fill(text_content, width=round(layer_width / font_size) * 2)
        lines = wrapped_text.split('\n')
        x_position, y_position_line = calculate_text_position(text_content, layer_width, font_size, alignment)
        y_position_line += layer.offset[1]
        # 计算段落文本的总高度
        total_height = len(lines) * font_size * 1.2 - font_size * 0.2
        # 根据垂直对齐方式调整y_position_line
        if '_pm' in layer.name:
            y_position_line += (layer.size[1] - total_height) / 2
        elif '_pb' in layer.name:
            y_position_line += layer.size[1] - total_height
        # 逐行绘制
        for line in lines:
            x_position, y_position = calculate_text_position(line, layer_width, font_size, alignment)
            draw.text((layer.offset[0] + x_position, y_position_line), line, fill=font_color, font=font)
            y_position_line += font_size * 1.2  # 1.2倍行距
    else:
        # 单行文本处理
        x_position, y_position = calculate_text_position(text_content, layer_width, font_size, alignment)
        draw.text((layer.offset[0] + x_position, layer.offset[1] + y_position), text_content, fill=font_color, font=font)

def update_image_layer(layer, new_image_path, pil_image):
    """更新图片图层内容

    :param PSDLayer layer: PSD图片图层
    :param str new_image_path: 新图片路径
    :param PIL.Image pil_image: PIL图像对象
    """
    layer.visible = False  # 防止PSD原始图层被输出到PIL
    if os.path.exists(new_image_path):
        new_image = Image.open(new_image_path).convert('RGBA')
        new_image = new_image.resize(layer.size)
        pil_image.alpha_composite(new_image, (layer.offset[0], layer.offset[1]))
    else:
        print(f"Warning: Image file {new_image_path} does not exist")

def sanitize_filename(filename):
    """清理文件名中的非法字符，确保跨平台兼容性

    :param str filename: 原始文件名
    :return str: 清理后的文件名
    """
    if not filename:
        return "unnamed"

    # 转换为字符串（处理非字符串类型）
    filename = str(filename)

    # 清理Excel特殊字符转义字符串
    filename = filename.replace('_x000D_', '')  # 回车符
    filename = filename.replace('_x000A_', '')  # 换行符
    filename = filename.replace('_x0009_', '')  # 制表符

    # Windows非法字符: / \ : * ? " < > |
    # 其他系统也可能不支持的字符: 控制字符 (0-31)
    illegal_chars = r'[\\/:*?"<>|\x00-\x1f]'

    # 替换非法字符为下划线
    import re
    sanitized = re.sub(illegal_chars, '_', filename)

    # 清理开头和结尾的空格和点（Windows不支持）
    sanitized = sanitized.strip(' .')

    # 限制文件名长度（避免文件系统限制）
    # 大多数文件系统支持255字符，但考虑路径长度限制，使用200字符
    if len(sanitized) > 200:
        sanitized = sanitized[:200]

    # 如果清理后为空，使用默认名称
    if not sanitized:
        sanitized = "unnamed"

    return sanitized

def save_image(output_dir, output_filename, image_format, pil_image):
    """保存PIL图片

    :param str output_dir: 输出目录
    :param str output_filename: 输出文件名
    :param str image_format: 图像格式
    :param PIL.Image pil_image: PIL图像对象
    """
    output_dir = os.path.join(output_dir, f'{current_datetime}_{file_name}')
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    final_output_path = os.path.join(output_dir, f'{output_filename}.{image_format}')
    if image_format.lower() == 'png':
        pil_image.save(final_output_path, format='PNG', optimize=True)
    else:  # 默认保存为jpg
        rgb_image = pil_image.convert('RGB')
        rgb_image.save(final_output_path, quality=quality, optimize=optimize)
    print(f"已导出图片: {final_output_path}")

def export_single_image(row, index, psd_object, psd_file_name):
    """处理单行数据并导出图像（单进程串行版本）

    :param pd.Series row: 包含单行数据的Series
    :param int index: 当前行索引
    :param PSDImage psd_object: 预加载的PSD对象
    :param str psd_file_name: PSD文件名（用于输出文件名）
    """
    pil_image = Image.new('RGBA', psd_object.size)

    def process_layers(layers):
        for layer in layers:
            layer_name = layer.name
            if layer_name and layer_name.startswith('@'):
                parts = layer_name[1:].split('#')
                if len(parts) == 2:
                    field_name, operation_type = parts
                    # 修改图层可见性
                    if operation_type.startswith('v'):
                        visibility = row[field_name]
                        set_layer_visibility(layer, visibility)
                    # 修改文字图层内容
                    elif operation_type.startswith('t'):
                        update_text_layer(layer, str(row[field_name]), pil_image, text_font)
                    # 修改图片图层内容
                    elif operation_type.startswith('i'):
                        update_image_layer(layer, str(row[field_name]), pil_image)
            if layer.is_visible():
                if layer.is_group():
                    # 如果是组，递归处理其子图层
                    process_layers(layer)
                else:
                    # 将非变量图层转换为PIL图像并合并到主图像上
                    layer_image = layer.topil()
                    if layer_image:
                        pil_image.alpha_composite(layer_image, (layer.offset[0], layer.offset[1]))

    # 处理所有图层
    process_layers(psd_object)

    # 输出图片
    # 生成带PSD后缀的输出文件名
    psd_base = os.path.splitext(psd_file_name)[0]
    excel_base = os.path.splitext(os.path.basename(excel_file_path))[0]

    # 更智能地提取PSD特有后缀
    if psd_base.startswith(excel_base):
        # 如果PSD文件名以Excel前缀开头，提取剩余部分作为后缀
        suffix = psd_base[len(excel_base):]
        # 处理井号分隔符
        if suffix.startswith('#'):
            suffix = suffix[1:]  # 去掉开头的井号
    else:
        # 如果PSD文件名不以Excel前缀开头，使用整个PSD文件名作为后缀
        suffix = psd_base

    # 处理空后缀情况
    if suffix:
        suffix = f"_{suffix}" if not suffix.startswith("_") else suffix
    else:
        suffix = ""

    base_filename = row.iloc[0] if pd.notna(row.iloc[0]) else f"image_{index + 1}"
    output_filename = f"{base_filename}{suffix}"
    # 清理文件名中的非法字符，确保跨平台兼容性
    output_filename = sanitize_filename(output_filename)
    save_image(output_path, output_filename, image_format, pil_image)

def get_matching_psds(excel_file):
    """获取匹配的PSD文件列表
    
    :param str excel_file: Excel文件名（不带扩展名）
    :return list: 匹配的PSD文件列表
    """
    base_name = os.path.splitext(excel_file)[0]
    matching_psds = []
    for f in os.listdir():
        if f.endswith('.psd'):
            # 提取文件名前缀（第一个井号前的部分）
            name_without_ext = os.path.splitext(f)[0]
            if '#' in name_without_ext:
                prefix = name_without_ext.split('#', 1)[0]
            else:
                prefix = name_without_ext
            if prefix == base_name:
                matching_psds.append(f)
    return matching_psds

def collect_psd_variables(psd_file_path: str) -> Set[str]:
    """收集PSD文件中的所有变量名
    
    :param str psd_file_path: PSD文件路径
    :return set: 变量名集合
    :raises FileNotFoundError: 当PSD文件不存在时
    :raises Exception: 当PSD文件损坏或读取失败时
    """
    variables = set()
    
    # 检查文件是否存在
    if not os.path.exists(psd_file_path):
        # 处理Windows路径编码问题
        # 使用ASCII编码确保路径在Windows控制台正确显示
        try:
            # 尝试直接显示路径
            error_msg = f"PSD文件不存在: {psd_file_path}"
        except UnicodeEncodeError:
            # 如果编码失败，使用安全的显示方式
            error_msg = f"PSD文件不存在: {psd_file_path.encode('ascii', errors='replace').decode('ascii')}"
        raise FileNotFoundError(error_msg)
    
    # 检查文件扩展名
    if not psd_file_path.lower().endswith('.psd'):
        raise ValueError(f"文件格式不支持，期望.psd文件: {psd_file_path}")
    
    try:
        psd = PSDImage.open(psd_file_path)
    except Exception as e:
        raise Exception(f"无法打开PSD文件 {psd_file_path}: {str(e)}")
    
    def process_layers(layers):
        for layer in layers:
            try:
                layer_name = layer.name
                if layer_name and layer_name.startswith('@'):
                    parts = layer_name[1:].split('#')
                    if len(parts) == 2:
                        field_name = parts[0]
                        variables.add(field_name)
                if layer.is_group():
                    process_layers(layer)
            except Exception as e:
                # 记录图层处理错误但继续处理其他图层
                print(f"警告：处理图层时出错: {str(e)}")
                continue
    
    try:
        process_layers(psd)
    except Exception as e:
        raise Exception(f"处理PSD文件图层时出错 {psd_file_path}: {str(e)}")
    
    return variables

def is_image_column(operation_type: str) -> bool:
    """判断是否为图片列
    
    :param str operation_type: 操作类型
    :return bool: 是否为图片列
    """
    return operation_type.startswith('i')

def validate_data(dataframe: pd.DataFrame, psd_templates: List[str]) -> Tuple[List[str], List[str]]:
    """验证Excel数据与PSD模板的匹配性
    
    :param pd.DataFrame dataframe: Excel数据
    :param list psd_templates: PSD模板文件列表
    :return tuple: (错误列表, 警告列表)
    """
    global validation_errors, validation_warnings
    validation_errors = []
    validation_warnings = []
    
    # 收集所有PSD变量
    all_psd_variables = set()
    image_columns = set()
    
    for psd_file in psd_templates:
        if not os.path.exists(psd_file):
            validation_errors.append(f"PSD template file does not exist: {psd_file}")
            continue
            
        variables = collect_psd_variables(psd_file)
        all_psd_variables.update(variables)
        
        # 识别图片列
        psd = PSDImage.open(psd_file)
        
        def check_image_layers(layers):
            for layer in layers:
                layer_name = layer.name
                if layer_name and layer_name.startswith('@'):
                    parts = layer_name[1:].split('#')
                    if len(parts) == 2:
                        field_name, operation_type = parts
                        if is_image_column(operation_type):
                            image_columns.add(field_name)
                if layer.is_group():
                    check_image_layers(layer)
        
        check_image_layers(psd)
    
    # 列名校验
    excel_columns = set(dataframe.columns)
    
    # 检查Excel中是否有PSD不存在的列
    extra_columns = excel_columns - all_psd_variables
    if extra_columns:
        for col in extra_columns:
            if col != 'File_name':  # File_name是特殊列，不算错误
                validation_warnings.append(f"Column '{col}' in Excel does not exist in PSD template")
    
    # 检查PSD必需变量在Excel中是否存在
    missing_columns = all_psd_variables - excel_columns
    if missing_columns:
        validation_errors.append(f"PSD模板中必需的变量在Excel中缺失: {', '.join(missing_columns)}")
    
    # 文件路径校验
    for image_col in image_columns:
        if image_col in dataframe.columns:
            for idx, file_path in enumerate(dataframe[image_col]):
                if pd.notna(file_path) and str(file_path).strip():
                    # 检查文件是否存在
                    if not os.path.exists(str(file_path)):
                        validation_errors.append(f"Image file does not exist: Row {idx+2}, Column '{image_col}', Path: {file_path}")
    
    return validation_errors, validation_warnings

def report_validation_results(errors: List[str], warnings: List[str]):
    """报告验证结果
    
    :param list errors: 错误列表
    :param list warnings: 警告列表
    """
    if not errors and not warnings:
        print("数据验证通过")
        return True
    
    print("\n" + "="*60)
    print("数据验证报告")
    print("="*60)
    
    if warnings:
        safe_print_message("\n警告:")
        for warning in warnings:
            safe_print_message(f"  - {warning}")
    
    if errors:
        safe_print_message("\n错误:")
        for error in errors:
            safe_print_message(f"  - {error}")
        safe_print_message("\n请修复上述错误后重新运行程序")
        return False
    
    return True

def preload_psd_templates(psd_files: List[str]) -> dict:
    """预加载PSD模板文件
    
    :param list psd_files: PSD文件列表
    :return dict: 预加载的PSD对象字典
    """
    psd_objects = {}
    print("\n预加载PSD模板...")
    
    for psd_file in psd_files:
        try:
            psd_objects[psd_file] = PSDImage.open(psd_file)
            print(f"  已加载: {psd_file}")
        except Exception as e:
            print(f"  加载失败: {psd_file} - {str(e)}")
            psd_objects[psd_file] = None
    
    return psd_objects

def log_export_activity(excel_file, image_count):
    """记录导出活动到日志文件

    :param str excel_file: 使用的Excel文件名
    :param int image_count: 导出的图片数量
    """
    log_file = 'log.csv'

    # 检查日志文件是否存在
    file_exists = os.path.exists(log_file)

    # 使用简单的字符串写入确保跨平台兼容性
    with open(log_file, 'a', encoding='utf-8') as f:
        # 如果文件不存在，写入表头
        if not file_exists:
            f.write('生成时间,图片数量,所用Excel文件\n')

        # 写入日志记录
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        f.write(f'{timestamp},{image_count},{excel_file}\n')

def psd_renderer_images():
    """批量输出图片
    """
    # ========== 调试代码开始 ==========
    print("="*50)
    print(f"Excel文件: {excel_file_path}")
    matching_psds = get_matching_psds(excel_file_path)
    print(f"匹配PSD: {matching_psds}")
    # ========== 调试代码结束 ==========
    
    # 读取Excel数据
    df = read_excel_file(excel_file_path)
    
    # 数据验证
    print("\n正在验证数据...")
    errors, warnings = validate_data(df, matching_psds)
    
    if not report_validation_results(errors, warnings):
        safe_print_message("数据验证失败，程序终止")
        sys.exit(1)
    
    # 预加载PSD模板
    psd_objects = preload_psd_templates(matching_psds)
    
    # 检查是否有PSD加载失败
    failed_psds = [psd_file for psd_file, psd_obj in psd_objects.items() if psd_obj is None]
    if failed_psds:
        safe_print_message(f"\n以下PSD模板加载失败，请检查文件完整性:")
        for failed_psd in failed_psds:
            safe_print_message(f"  - {failed_psd}")
        sys.exit(1)
    
    # 单进程串行处理
    total_images = 0
    success_count = 0
    error_count = 0

    # 统计总任务数
    for psd_file in matching_psds:
        if psd_objects[psd_file] is not None:
            total_images += len(df)

    # 如果没有找到任何任务，直接退出
    if total_images == 0:
        safe_print_message("\n警告：没有找到匹配的PSD模板或数据，跳过图片生成")
        return

    # 开始串行处理
    print(f"\n开始单进程串行处理 {total_images} 个任务...")

    # 使用tqdm显示进度条
    with tqdm(total=total_images, desc="正在导出图片", unit="张") as pbar:
        for psd_file in matching_psds:
            if psd_objects[psd_file] is not None:
                for index, row in df.iterrows():
                    try:
                        # 使用原有的串行函数
                        export_single_image(row, index, psd_objects[psd_file], psd_file)
                        success_count += 1
                    except Exception as e:
                        error_count += 1
                        error_msg = f"第 {index+1} 行数据处理失败: {str(e)}"
                        safe_print_message(error_msg)

                        # 根据错误类型提供建议
                        if "PermissionError" in str(e) or "权限" in str(e):
                            safe_print_message("  提示：请检查文件权限设置")
                        elif "FileNotFoundError" in str(e) or "文件不存在" in str(e):
                            safe_print_message("  提示：请检查相关文件是否存在")
                        elif "MemoryError" in str(e) or "内存" in str(e):
                            safe_print_message("  提示：内存不足，尝试关闭其他程序")
                    finally:
                        pbar.update(1)  # 无论成功还是失败，都更新进度条

    # 输出统计信息
    print(f"\n处理统计:")
    print(f"  总任务数: {total_images} 张")
    print(f"  成功: {success_count} 张")
    print(f"  失败: {error_count} 张")

    if error_count > 0:
        print(f"\n有 {error_count} 个任务失败，请检查错误信息")
        # 如果错误率过高，给出建议
        error_rate = error_count / total_images
        if error_rate > 0.5:  # 超过50%失败率
            safe_print_message("  建议：检查PSD模板和Excel数据格式是否正确")
        elif error_rate > 0.2:  # 超过20%失败率
            safe_print_message("  建议：检查资源文件是否存在")

    print(f"\n处理完成，共生成 {success_count} 张图片")

    # 记录日志 - 只在有成功生成的图片时才记录
    if success_count > 0:
        log_export_activity(file_name, success_count)
        print("批量导出完成！")
    else:
        print("警告：没有成功生成任何图片，跳过日志记录")

    # 打开输出文件夹
    # 在第一次保存图片后获取准确的输出目录
    first_image_output_dir = os.path.join(output_path, f'{current_datetime}_{file_name}')
    # 跨平台兼容的文件夹打开方式
    if os.name == 'nt':  # Windows
        os.system(f'explorer "{first_image_output_dir}"')
    else:  # macOS/Linux
        os.system(f'open "{first_image_output_dir}"')


if __name__ == "__main__":
    # 切换到脚本所在目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)
    
    # 设置项
    if len(sys.argv) < 4:
        print("用法: python batch_export.py [Excel文件前缀] [字体文件] [输出格式]")
        print("示例: python batch_export.py 1 AlibabaPuHuiTi-2-85-Bold.ttf jpg")
        sys.exit(1)
    
    file_name = sys.argv[1]  # 从命令行参数获取使用第几套数据和模版
    font_file = sys.argv[2]  # 从命令行参数获取字体文件
    image_format = sys.argv[3]  # 从命令行参数获取输出图片格式
    
    quality = 95
    optimize = False
    
    # 文件路径
    output_path = 'export'
    excel_file_path = f'{file_name}.xlsx'
    psd_file_path = f'{file_name}.psd'
    text_font = font_file

    # 批量输出图片
    current_datetime = datetime.now().strftime('%Y%m%d_%H%M%S')
    psd_renderer_images()
