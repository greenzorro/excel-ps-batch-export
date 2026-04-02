"""
File: file_monitor.py
Project: excel-ps-batch-export
Created: 2024-09-26 10:36:43
Author: Victor Cheng
Email: hi@victor42.work
Description: 监控数据文件并自动执行批量图片输出
"""

# 设置项
image_format = 'jpg'  # jpg/png
output_dir = '../export'  # 输出目录（默认../export）

import os
import sys
import subprocess
import asyncio
import hashlib
import xlsx_generator

def get_file_hash(file_path):
    """计算文件的 MD5 哈希值"""
    if not os.path.exists(file_path):
        return None
    try:
        hasher = hashlib.md5()
        with open(file_path, 'rb') as f:
            # 每次读取 8KB，避免大文件占用过多内存
            for chunk in iter(lambda: f.read(8192), b""):
                hasher.update(chunk)
        return hasher.hexdigest()
    except Exception as e:
        print(f"警告：无法计算文件哈希 {file_path}: {e}")
        return None

async def monitor_excel_file(base_name, file_path, psd_files):
    """监控数据文件变化

    :param str base_name: Excel文件基础名（不含扩展名和目录）
    :param str file_path: 数据文件完整路径（.xlsx 或 _raw.csv）
    :param list psd_files: 关联的PSD模板文件列表
    """
    print(f"正在监控数据文件 {file_path} (关联PSD: {', '.join(psd_files)})……")
    
    # 初始化哈希值和修改时间
    last_modified_time = os.path.getmtime(file_path) if os.path.exists(file_path) else 0
    last_hash = get_file_hash(file_path)
    
    while True:
        await asyncio.sleep(5)  # 每5秒检查一次
        if not os.path.exists(file_path):
            continue
            
        current_modified_time = os.path.getmtime(file_path)
        
        # 只有当 mtime 发生变化时，才去计算哈希（性能优化）
        if current_modified_time != last_modified_time:
            # 去抖动：等待 1 秒，确保 Resilio Sync 或 Excel 完成写入
            await asyncio.sleep(1)
            
            current_hash = get_file_hash(file_path)
            
            # 核心判断：只有内容哈希变了，才视为真正修改
            if current_hash and current_hash != last_hash:
                print(f"\n[{base_name}] 内容已更新，开始渲染...")
                cmd = [sys.executable, 'psd_renderer.py', base_name, image_format]
                if output_dir != '../export':
                    cmd.append(output_dir)
                subprocess.run(cmd)
                
                # 更新状态
                last_hash = current_hash
                last_modified_time = os.path.getmtime(file_path)
                print(f"[{base_name}] 处理完成，继续监控中……")
            else:
                # 如果哈希没变，说明只是元数据（如同步状态、访问时间）变化，忽略它
                # 但仍需更新 last_modified_time 以避免下次循环重复检测
                last_modified_time = current_modified_time

async def main():
    """主函数"""
    tasks = []
    for base_name, excel_file_path, psd_files in excel_psd_pairs:
        tasks.append(monitor_excel_file(base_name, excel_file_path, psd_files))
    await asyncio.gather(*tasks)

if __name__ == "__main__":
    xlsx_generator.main()

    # 自动获取 workspace 文件夹中所有.xlsx或.xls文件，并匹配对应的PSD模板（支持多个模板）
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    workspace_dir = "../workspace"
    excel_psd_pairs = []
    for file in os.listdir(workspace_dir):
        if file.endswith(('.xlsx', '.xls')):
            base_name = os.path.splitext(file)[0]
            excel_file_path = os.path.join(workspace_dir, file)
            # 匹配所有 base_name#*.psd 文件
            matching_psds = [f for f in os.listdir(workspace_dir)
                            if f.startswith(f"{base_name}#") and f.endswith('.psd')]
            # 如果没有带#的PSD，尝试匹配 base_name.psd
            if not matching_psds:
                single_psd = f"{base_name}.psd"
                single_psd_path = os.path.join(workspace_dir, single_psd)
                if os.path.exists(single_psd_path):
                    matching_psds = [single_psd]

            if matching_psds:
                # 检查是否有变换规则文件
                json_rule_path = os.path.join(workspace_dir, f"{base_name}.json")
                if os.path.exists(json_rule_path):
                    # 有变换规则：监控 _raw.csv
                    raw_csv_path = os.path.join(workspace_dir, f"{base_name}_raw.csv")
                    if os.path.exists(raw_csv_path):
                        excel_psd_pairs.append((base_name, raw_csv_path, matching_psds))
                    else:
                        excel_psd_pairs.append((base_name, raw_csv_path, matching_psds))
                        print(f"  注意: {raw_csv_path} 尚未创建，将在创建后开始监控")
                else:
                    # 无变换规则：监控 .xlsx
                    excel_psd_pairs.append((base_name, excel_file_path, matching_psds))

    asyncio.run(main())
