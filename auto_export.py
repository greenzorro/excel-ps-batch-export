'''
File: auto_export.py
Project: excel-ps-batch-export
Created: 2024-09-26 10:36:43
Author: Victor Cheng
Email: greenzorromail@gmail.com
Description: 监控数据文件并自动执行批量图片输出
'''

import os
import subprocess
import asyncio

# 自动获取当前文件夹中所有以数字+"_"开头的.xlsx文件，提取数据集编号
nums = [int(file.split('_')[0]) for file in os.listdir() if file.endswith('.xlsx') and file[0].isdigit()]

async def monitor_excel_file(file_path, num):
    """监控Excel文件变化

    :param str file_path: Excel文件路径
    :param int num: 数据集编号
    """
    print(f"正在监控数据文件 {file_path}……")
    last_modified_time = os.path.getmtime(file_path)
    while True:
        await asyncio.sleep(5)  # 每5秒检查一次
        current_modified_time = os.path.getmtime(file_path)
        if current_modified_time != last_modified_time:
            print(f"{file_path} 文件已被修改，正在执行 batch_export.py...")
            subprocess.run(['python', 'batch_export.py', str(num)])
            last_modified_time = current_modified_time
            print(f"正在监控数据文件……")

async def main():
    """主函数"""
    tasks = []
    for num in nums:
        excel_file_path = f'{num}_data.xlsx'
        tasks.append(monitor_excel_file(excel_file_path, num))
    await asyncio.gather(*tasks)

if __name__ == "__main__":
    asyncio.run(main())
