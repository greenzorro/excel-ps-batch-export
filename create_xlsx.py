'''
File: create_xlsx.py
Project: excel-ps-batch-export
Created: 2024-10-14 11:41:59
Author: Victor Cheng
Email: greenzorromail@gmail.com
Description: 为PSD模板创建同名Excel文件，并根据PSD里的变量图层初始化Excel列
'''

import os
import pandas as pd
from psd_tools import PSDImage

def init_xlsx(file):
    """初始化Excel文件

    :param str file: PSD文件名
    """
    base_name = os.path.splitext(file)[0]
    excel_file_xlsx = f'{base_name}.xlsx'
    psd_file_path = f'{base_name}.psd'
    
    # 读取PSD文件
    psd = PSDImage.open(psd_file_path)
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
                # 如果是组，递归处理其子图层
                process_layers(layer)
    
    # 处理所有图层
    process_layers(psd)
    
    # 按文本、可见性、图片的顺序排列列名，并在前面增加"File_name"列
    columns = ['File_name'] + text_columns + visibility_columns + image_columns
    
    # 创建DataFrame并写入Excel文件
    df = pd.DataFrame(columns=columns)
    
    # 增加一行示例数据
    example_data = {'File_name': "文件名"}
    for col in text_columns:
        example_data[col] = "示例文字"
    for col in visibility_columns:
        example_data[col] = "TRUE"
    for col in image_columns:
        example_data[col] = "文件/路径/图片.jpg"
    
    df.loc[0] = example_data
    df.to_excel(excel_file_xlsx, index=False)
    print(f"已初始化文件: {excel_file_xlsx}")

def create_xlsx(file):
    """创建Excel文件

    :param str file: PSD文件名
    """
    base_name = os.path.splitext(file)[0]
    excel_file_xlsx = f'{base_name}.xlsx'
    excel_file_xls = f'{base_name}.xls'
    if not os.path.exists(excel_file_xlsx) and not os.path.exists(excel_file_xls):
        # 创建空的DataFrame
        df = pd.DataFrame()
        # 保存为Excel文件
        df.to_excel(excel_file_xlsx, index=False)
        init_xlsx(file)
        print(f"已创建文件: {excel_file_xlsx}")

def main():
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    for file in os.listdir():
        if file.endswith('.psd'):
            create_xlsx(file)

if __name__ == "__main__":
    main()
