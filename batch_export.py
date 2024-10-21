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
import sys
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

# 读取Excel文件
def read_excel_file(file_path):
    """读取Excel文件

    :param str file_path: Excel文件路径
    :return pd.DataFrame: 包含Excel数据的DataFrame
    """
    df = pd.read_excel(file_path, sheet_name=0)
    return df

# 修改图层可见性
def set_layer_visibility(layer, visibility):
    """设置图层可见性

    :param PSDLayer layer: PSD图层对象
    :param bool visibility: 是否可见
    """
    layer.visible = visibility

# 修改文字图层内容
def update_text_layer(layer, text_content, pil_image):
    """更新文字图层内容

    :param PSDLayer layer: PSD文字图层
    :param str text_content: 新的文字内容
    :param PIL.Image pil_image: PIL图像对象
    """
    layer.visible = False  # 防止PSD原始图层被输出到PIL
    font_info = layer.engine_dict
    font_size = font_info['StyleRun']['RunArray'][0]['StyleSheet']['StyleSheetData']['FontSize']
    # 检查是否存在 'FillColor' 键
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
    font = ImageFont.truetype(text_font, int(font_size))
    draw = ImageDraw.Draw(pil_image)
    layer_width = layer.size[0]
    text_width = draw.textbbox((0, 0), text_content, font=font)[2] - draw.textbbox((0, 0), text_content, font=font)[0]
    # 从图层名读取文字对齐方向
    if layer.name.endswith('t_c'):  # 计算居中位置
        x_position = (layer.offset[0] + (layer_width - text_width) / 2)
    elif layer.name.endswith('t_r'):  # 计算右对齐位置
        x_position = layer.offset[0] + layer_width - text_width
    else:  # 计算左对齐位置
        x_position = layer.offset[0]
    # 修正文字位置偏移
    x_offset = font_size * 0.04
    y_offset = font_size * 0.25
    draw.text((x_position - x_offset, layer.offset[1] - y_offset), text_content, fill=font_color, font=font)

# 修改图片图层内容
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

# 保存PIL图片
def save_image(output_dir, output_filename, image_format, pil_image):
    """保存处理后的图像

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

# 输出单张图片
def export_single_image(row, index):
    """处理单行数据并导出图像

    :param pd.Series row: 包含单行数据的Series
    :param int index: 当前行索引
    """
    psd = PSDImage.open(psd_file_path)
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
    output_filename = row.iloc[0] if pd.notna(row.iloc[0]) else f"image_{index + 1}"
    save_image(output_path, output_filename, image_format, pil_image)

# 批量输出图片
def batch_export_images():
    df = read_excel_file(excel_file_path)
    for index, row in df.iterrows():
        print(f"正在处理第 {index + 1} 行数据...")
        export_single_image(row, index)
    print("批量导出完成！")


if __name__ == "__main__":
    # 切换到脚本所在目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)

    # 批量输出图片
    current_datetime = datetime.now().strftime('%Y%0m%d_%H%M%S')
    batch_export_images()
