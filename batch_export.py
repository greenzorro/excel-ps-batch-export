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

# 设置项
file_name = sys.argv[1]  # 从命令行参数获取使用第几套数据和模版
font_file = sys.argv[2]  # 从命令行参数获取字体文件
image_format = sys.argv[3]  # 从命令行参数获取输出图片格式

# file_name = '1'  # 手动选择使用哪套数据和模版
# font_file = 'AlibabaPuHuiTi-2-85-Bold.ttf'
# image_format = 'jpg'  # jpg/png
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
    """
    df = pd.read_excel(file_path, sheet_name=0)
    return df

def set_layer_visibility(layer, visibility):
    """设置图层可见性

    :param PSDLayer layer: PSD图层对象
    :param bool visibility: 是否可见
    """
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
    """
    # 计算文字宽度，考虑中文和英文字符占的宽度不同
    text_width = 0
    for char in text:
        if '\u4e00' <= char <= '\u9fff':  # 判断是否为中文字符
            text_width += font_size  # 中文字符宽度为字体大小
        else:
            text_width += font_size * 0.5  # 英文字符宽度为字体大小的一半
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
        print(f"警告：图片文件 {new_image_path} 不存在")

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

def export_single_image(row, index, current_psd_path):
    """处理单行数据并导出图像

    :param pd.Series row: 包含单行数据的Series
    :param int index: 当前行索引
    :param str current_psd_path: 当前处理的PSD文件路径
    """
    psd = PSDImage.open(current_psd_path)
    pil_image = Image.new('RGBA', psd.size)

    def process_layers(layers):
        for layer in layers:
            layer_name = layer.name
            if layer_name.startswith('@'):
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
    process_layers(psd)
    
    # 输出图片
    # 生成带PSD后缀的输出文件名
    psd_base = os.path.splitext(os.path.basename(current_psd_path))[0]
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
    
    df = read_excel_file(excel_file_path)
    total_images = 0
    
    # 处理所有匹配的PSD文件
    for psd_file in matching_psds:
        print(f"\n处理PSD文件: {psd_file}")
        for index, row in df.iterrows():
            print(f"  正在处理第 {index + 1} 行数据...")
            export_single_image(row, index, psd_file)
            total_images += 1
    
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
