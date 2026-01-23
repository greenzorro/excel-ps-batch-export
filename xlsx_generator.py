"""
File: xlsx_generator.py
Project: excel-ps-batch-export
Created: 2024-10-14 11:41:59
Author: Victor Cheng
Email: hi@victor42.work
Description: Excel生成器 - 为PSD模板创建Excel配置文件，根据变量图层初始化列
"""

import os
import pandas as pd
from psd_tools import PSDImage


def main():
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    # 按前缀分组PSD文件（从workspace目录读取）
    psd_groups = {}
    workspace_dir = "workspace"
    for file in os.listdir(workspace_dir):
        if file.endswith('.psd'):
            base_name = os.path.splitext(file)[0]
            # 提取前缀（第一个井号前的部分）
            if '#' in base_name:
                prefix = base_name.split('#', 1)[0]
            else:
                prefix = base_name

            if prefix not in psd_groups:
                psd_groups[prefix] = []
            psd_groups[prefix].append(file)

    # 为每组创建共享Excel文件
    for prefix, psd_files in psd_groups.items():
        excel_file = f"{prefix}.xlsx"
        excel_path = os.path.join(workspace_dir, excel_file)
        # 跳过已存在Excel的情况
        if not os.path.exists(excel_path):
            # 收集所有PSD的变量
            all_text_columns = []
            all_visibility_columns = []
            all_image_columns = []

            for psd_file in psd_files:
                psd_path = os.path.join(workspace_dir, psd_file)
                psd = PSDImage.open(psd_path)
                text_columns, visibility_columns, image_columns = extract_variables(psd)
                all_text_columns.extend([c for c in text_columns if c not in all_text_columns])
                all_visibility_columns.extend([c for c in visibility_columns if c not in all_visibility_columns])
                all_image_columns.extend([c for c in image_columns if c not in all_image_columns])

            # 创建DataFrame
            columns = ['File_name'] + all_text_columns + all_visibility_columns + all_image_columns
            df = pd.DataFrame(columns=columns)

            # 添加示例行
            example_data = {'File_name': "文件名"}
            for col in all_text_columns:
                example_data[col] = "示例文字"
            for col in all_visibility_columns:
                example_data[col] = "TRUE"
            for col in all_image_columns:
                example_data[col] = "文件/路径/图片.jpg"

            df.loc[0] = example_data
            df.to_excel(excel_path, index=False)
            print(f"已创建共享Excel文件: {excel_path}")

def extract_variables(psd):
    """从PSD提取所有变量图层
    
    :param PSDImage psd: PSD文件对象
    :return tuple: (文本列, 可见性列, 图片列)
    """
    text_columns = []
    visibility_columns = []
    image_columns = []
    
    def process_layers(layers):
        for layer in layers:
            if layer.name.startswith('@'):
                parts = layer.name[1:].split('#')
                if len(parts) == 2:
                    field_name = parts[0]
                    category = parts[1][0]
                    if category == 't' and field_name not in text_columns:
                        text_columns.append(field_name)
                    elif category == 'v' and field_name not in visibility_columns:
                        visibility_columns.append(field_name)
                    elif category == 'i' and field_name not in image_columns:
                        image_columns.append(field_name)
            if layer.is_group():
                # 递归处理组内图层
                process_layers(layer)
    
    process_layers(psd)
    return text_columns, visibility_columns, image_columns

if __name__ == "__main__":
    main()
