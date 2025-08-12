'''
File: batch_export.py
Project: excel-ps-batch-export
Created: 2024-09-25 02:07:52
Author: Victor Cheng
Email: greenzorromail@gmail.com
Description: 单次批量输出图片
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
from concurrent.futures import ProcessPoolExecutor, as_completed
from tqdm import tqdm
import multiprocessing
import copy

# 全局变量存储错误和警告
validation_errors = []
validation_warnings = []

# 设置项
file_name = sys.argv[1]  # 从命令行参数获取使用第几套数据和模版
font_file = sys.argv[2]  # 从命令行参数获取字体文件
image_format = sys.argv[3]  # 从命令行参数获取输出图片格式

quality = 95
optimize = False
current_datetime = ''

# 文件路径
output_path = 'export'
excel_file_path = f'{file_name}.xlsx'
psd_file_path = f'{file_name}.psd'
text_font = f'assets/fonts/{font_file}'

def read_excel_file(file_path):
    """读取Excel文件

    :param str file_path: Excel文件路径
    :return pd.DataFrame: 包含Excel数据的DataFrame
    :raises FileNotFoundError: 当文件不存在时抛出异常
    :raises ValueError: 当文件格式不支持时抛出异常
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Excel文件不存在: {file_path}")
    
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
    if not isinstance(visibility, bool):
        raise TypeError(f"visibility必须是布尔值，收到类型: {type(visibility).__name__}")
    layer.visible = visibility

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
    
    # 计算文字宽度，考虑中文和英文字符占的宽度不同
    text_width = 0
    for char in text:
        if '\u4e00' <= char <= '\u9fff':  # 判断是否为中文字符
            text_width += font_size  # 中文字符宽度为字体大小
        else:
            text_width += font_size * 0.5  # 英文字符宽度为字体大小的一半
    
    # 计算位置
    if alignment == 'center':  # 计算居中位置
        x_position = (layer_width - text_width) / 2
    elif alignment == 'right':  # 计算右对齐位置
        x_position = layer_width - text_width
    else:  # 计算左对齐位置
        x_position = 0
    
    # 修正文字位置偏移
    x_offset = font_size * 0.01
    y_offset = font_size * 0.26
    return x_position - x_offset, -y_offset

def update_text_layer(layer, text_content, pil_image):
    """更新文字图层内容

    :param PSDLayer layer: PSD文字图层
    :param str text_content: 新的文字内容
    :param PIL.Image pil_image: PIL图像对象
    """
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

def export_single_image_task(task_data):
    """并行处理单行数据并导出图像
    
    :param dict task_data: 任务数据包含 row, index, psd_object, psd_file_name, excel_file_path, output_path, image_format, text_font, quality, optimize, current_datetime
    :return str: 输出的文件路径
    """
    row = task_data['row']
    index = task_data['index']
    psd_object = task_data['psd_object']
    psd_file_name = task_data['psd_file_name']
    excel_file_path = task_data['excel_file_path']
    output_path = task_data['output_path']
    image_format = task_data['image_format']
    text_font = task_data['text_font']
    quality = task_data['quality']
    optimize = task_data['optimize']
    current_datetime = task_data['current_datetime']
    
    # 创建PSD对象的深拷贝，避免并发问题
    psd_copy = copy.deepcopy(psd_object)
    pil_image = Image.new('RGBA', psd_copy.size)

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
                        update_text_layer(layer, str(row[field_name]), pil_image)
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
    process_layers(psd_copy)
    
    # 输出图片
    # 生成带PSD后缀的输出文件名
    psd_base = os.path.splitext(psd_file_name)[0]
    excel_base = os.path.splitext(os.path.basename(excel_file_path))[0]
    suffix = psd_base.replace(excel_base, "")  # 提取PSD特有后缀
    
    # 处理空后缀情况
    if suffix:
        suffix = f"_{suffix}" if not suffix.startswith("_") else suffix
    else:
        suffix = ""
    
    base_filename = row.iloc[0] if pd.notna(row.iloc[0]) else f"image_{index + 1}"
    output_filename = f"{base_filename}{suffix}"
    
    # 保存图片
    output_dir = os.path.join(output_path, f'{current_datetime}_{os.path.splitext(os.path.basename(excel_file_path))[0]}')
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    final_output_path = os.path.join(output_dir, f'{output_filename}.{image_format}')
    
    if image_format.lower() == 'png':
        pil_image.save(final_output_path, format='PNG', optimize=True)
    else:  # 默认保存为jpg
        rgb_image = pil_image.convert('RGB')
        rgb_image.save(final_output_path, quality=quality, optimize=optimize)
    
    return final_output_path

def export_single_image(row, index, psd_object, psd_file_name):
    """处理单行数据并导出图像（串行版本）

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
                        update_text_layer(layer, str(row[field_name]), pil_image)
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
    suffix = psd_base.replace(excel_base, "")  # 提取PSD特有后缀
    
    # 处理空后缀情况
    if suffix:
        suffix = f"_{suffix}" if not suffix.startswith("_") else suffix
    else:
        suffix = ""
    
    base_filename = row.iloc[0] if pd.notna(row.iloc[0]) else f"image_{index + 1}"
    output_filename = f"{base_filename}{suffix}"
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
    """
    variables = set()
    psd = PSDImage.open(psd_file_path)
    
    def process_layers(layers):
        for layer in layers:
            layer_name = layer.name
            if layer_name and layer_name.startswith('@'):
                parts = layer_name[1:].split('#')
                if len(parts) == 2:
                    field_name = parts[0]
                    variables.add(field_name)
            if layer.is_group():
                process_layers(layer)
    
    process_layers(psd)
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
        print("✅ 数据验证通过")
        return True
    
    print("\n" + "="*60)
    print("📋 数据验证报告")
    print("="*60)
    
    if warnings:
        print("\n⚠️  警告:")
        for warning in warnings:
            print(f"  - {warning}")
    
    if errors:
        print("\n❌ 错误:")
        for error in errors:
            print(f"  - {error}")
        print("\n❗ 请修复上述错误后重新运行程序")
        return False
    
    return True

def preload_psd_templates(psd_files: List[str]) -> dict:
    """预加载PSD模板文件
    
    :param list psd_files: PSD文件列表
    :return dict: 预加载的PSD对象字典
    """
    psd_objects = {}
    print("\n🔄 预加载PSD模板...")
    
    for psd_file in psd_files:
        try:
            psd_objects[psd_file] = PSDImage.open(psd_file)
            print(f"  ✅ 已加载: {psd_file}")
        except Exception as e:
            print(f"  ❌ 加载失败: {psd_file} - {str(e)}")
            psd_objects[psd_file] = None
    
    return psd_objects

def log_export_activity(excel_file, image_count):
    """记录导出活动到日志文件
    
    :param str excel_file: 使用的Excel文件名
    :param int image_count: 导出的图片数量
    """
    log_file = 'log.csv'
    log_entry = {
        '生成时间': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        '图片数量': image_count,
        '所用Excel文件': excel_file
    }
    
    # 检查日志文件是否存在
    file_exists = os.path.exists(log_file)
    
    # 写入日志
    with open(log_file, 'a', newline='', encoding='utf-8') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=['生成时间', '图片数量', '所用Excel文件'])
        
        # 如果文件不存在，写入表头
        if not file_exists:
            writer.writeheader()
        
        # 写入日志记录
        writer.writerow(log_entry)

def batch_export_images():
    """批量输出图片
    """
    # ========== 调试代码开始 ==========
    print("="*50)
    print(f"📁 Excel文件: {excel_file_path}")
    matching_psds = get_matching_psds(excel_file_path)
    print(f"🔍 匹配PSD: {matching_psds}")
    # ========== 调试代码结束 ==========
    
    # 读取Excel数据
    df = read_excel_file(excel_file_path)
    
    # 数据验证
    print("\n🔍 正在验证数据...")
    errors, warnings = validate_data(df, matching_psds)
    
    if not report_validation_results(errors, warnings):
        print("❌ 数据验证失败，程序终止")
        sys.exit(1)
    
    # 预加载PSD模板
    psd_objects = preload_psd_templates(matching_psds)
    
    # 检查是否有PSD加载失败
    failed_psds = [psd_file for psd_file, psd_obj in psd_objects.items() if psd_obj is None]
    if failed_psds:
        print(f"\n❌ 以下PSD模板加载失败，请检查文件完整性:")
        for failed_psd in failed_psds:
            print(f"  - {failed_psd}")
        sys.exit(1)
    
    # 准备并行任务
    tasks = []
    total_images = 0
    
    # 为每个PSD文件和每行数据创建任务
    for psd_file in matching_psds:
        if psd_objects[psd_file] is not None:
            for index, row in df.iterrows():
                task_data = {
                    'row': row,
                    'index': index,
                    'psd_object': psd_objects[psd_file],
                    'psd_file_name': psd_file,
                    'excel_file_path': excel_file_path,
                    'output_path': output_path,
                    'image_format': image_format,
                    'text_font': text_font,
                    'quality': quality,
                    'optimize': optimize,
                    'current_datetime': current_datetime
                }
                tasks.append(task_data)
                total_images += 1
    
    # 并行处理
    print(f"\n🚀 开始并行处理 {total_images} 个任务...")
    
    # 使用CPU核心数的80%作为最大工作进程数
    max_workers = min(multiprocessing.cpu_count(), max(1, int(multiprocessing.cpu_count() * 0.8)))
    
    with ProcessPoolExecutor(max_workers=max_workers) as executor:
        # 使用tqdm显示进度
        futures = [executor.submit(export_single_image_task, task) for task in tasks]
        
        # 等待所有任务完成并显示进度
        for future in tqdm(as_completed(futures), total=len(futures), desc="正在导出图片", unit="张"):
            try:
                result = future.result()
                # 可以在这里记录成功导出的文件
            except Exception as e:
                print(f"❌ 任务执行失败: {str(e)}")
    
    print(f"\n✅ 并行处理完成，共处理 {total_images} 张图片")
    
    # 记录日志
    log_export_activity(file_name, total_images)
    print("批量导出完成！")
    
    # 打开输出文件夹
    # 在第一次保存图片后获取准确的输出目录
    first_image_output_dir = os.path.join(output_path, f'{current_datetime}_{file_name}')
    os.system(f'open "{first_image_output_dir}"')


if __name__ == "__main__":
    # 切换到脚本所在目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)

    # 批量输出图片
    current_datetime = datetime.now().strftime('%Y%0m%d_%H%M%S')
    batch_export_images()
