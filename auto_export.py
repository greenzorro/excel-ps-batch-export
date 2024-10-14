'''
File: auto_export.py
Project: excel-ps-batch-export
Created: 2024-09-26 10:36:43
Author: Victor Cheng
Email: greenzorromail@gmail.com
Description: 监控数据文件并自动执行批量图片输出
'''

# 设置项
font_file = 'AlibabaPuHuiTi-2-85-Bold.ttf'
image_format = 'jpg'  # jpg/png

import os
import sys
import subprocess
import asyncio
import create_xlsx

async def monitor_excel_file(excel_file_path):
    """监控Excel文件变化

    :param str excel_file_path: Excel文件路径
    """
    print(f"正在监控数据文件 {excel_file_path}……")
    last_modified_time = os.path.getmtime(excel_file_path)
    while True:
        await asyncio.sleep(5)  # 每5秒检查一次
        current_modified_time = os.path.getmtime(excel_file_path)
        if current_modified_time != last_modified_time:
            print(f"{excel_file_path} 文件已被修改，正在执行 batch_export.py...")
            print(f'test: {os.path.splitext(excel_file_path)[0]}')
            subprocess.run([sys.executable, 'batch_export.py', str(os.path.splitext(excel_file_path)[0]), font_file, image_format])
            last_modified_time = current_modified_time
            print(f"正在监控数据文件……")

async def main():
    """主函数"""
    tasks = []
    for excel_file, psd_file in excel_psd_pairs:
        tasks.append(monitor_excel_file(excel_file))
    await asyncio.gather(*tasks)

if __name__ == "__main__":
    create_xlsx.main()
    
    # 自动获取当前文件夹中所有.xlsx或.xls文件，并检查是否有同名的.psd文件
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    excel_psd_pairs = []
    for file in os.listdir():
        if file.endswith(('.xlsx', '.xls')):
            base_name = os.path.splitext(file)[0]
            psd_file = f'{base_name}.psd'
            if os.path.exists(psd_file):
                excel_psd_pairs.append((file, psd_file))

    asyncio.run(main())
