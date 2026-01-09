'''
File: file_monitor.py
Project: excel-ps-batch-export
Created: 2024-09-26 10:36:43
Author: Victor Cheng
Email: hi@victor42.work
Description: 监控数据文件并自动执行批量图片输出
'''

# 设置项
image_format = 'jpg'  # jpg/png

import os
import sys
import subprocess
import asyncio
import xlsx_generator

async def monitor_excel_file(excel_file_path, psd_files):
    """监控Excel文件变化

    :param str excel_file_path: Excel文件路径
    :param list psd_files: 关联的PSD模板文件列表
    """
    print(f"正在监控数据文件 {excel_file_path} (关联PSD: {', '.join(psd_files)})……")
    # 初始化为当前时间，避免启动时的误触发
    last_modified_time = os.path.getmtime(excel_file_path)
    while True:
        await asyncio.sleep(5)  # 每5秒检查一次
        current_modified_time = os.path.getmtime(excel_file_path)
        if current_modified_time != last_modified_time:
            print(f"{excel_file_path} 文件已被修改，正在执行 psd_renderer.py...")
            subprocess.run([sys.executable, 'psd_renderer.py', str(os.path.splitext(excel_file_path)[0]), image_format])
            last_modified_time = current_modified_time
            print(f"正在监控数据文件……")

async def main():
    """主函数"""
    tasks = []
    for excel_file, psd_files in excel_psd_pairs:
        tasks.append(monitor_excel_file(excel_file, psd_files))
    await asyncio.gather(*tasks)

if __name__ == "__main__":
    xlsx_generator.main()
    
    # 自动获取当前文件夹中所有.xlsx或.xls文件，并匹配对应的PSD模板（支持多个模板）
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    excel_psd_pairs = []
    for file in os.listdir():
        if file.endswith(('.xlsx', '.xls')):
            base_name = os.path.splitext(file)[0]
            # 匹配所有 base_name#*.psd 文件
            matching_psds = [f for f in os.listdir() 
                            if f.startswith(f"{base_name}#") and f.endswith('.psd')]
            # 如果没有带#的PSD，尝试匹配 base_name.psd
            if not matching_psds:
                single_psd = f"{base_name}.psd"
                if os.path.exists(single_psd):
                    matching_psds = [single_psd]
            
            if matching_psds:
                excel_psd_pairs.append((file, matching_psds))

    asyncio.run(main())
