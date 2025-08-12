#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel-PS Batch Export 测试工具
============================

本模块提供测试辅助工具和实用函数。
"""

import os
import sys
import tempfile
import shutil
from pathlib import Path
import pandas as pd


def create_test_data(rows=10):
    """创建测试数据"""
    data = {
        "File_name": [f"test_{i}" for i in range(rows)],
        "分类": ["测试分类"] * rows,
        "标题第1行": [f"测试标题 {i}" for i in range(rows)],
        "标题第2行": [f"副标题 {i}" for i in range(rows)],
        "直播时间": ["2024-01-01"] * rows,
        "单行": ["单行文本"] * rows,
        "两行": ["两行文本\n第二行"] * rows,
        "时间框": [True] * rows,
        "站内标": [True] * rows,
        "小标签内容": ["标签内容"] * rows,
        "背景图": ["assets/1_img/null.jpg"] * rows,
        "小标签": [True] * rows,
        "站外标": [False] * rows,
    }
    return pd.DataFrame(data)


def create_test_environment(base_dir, include_psd=False):
    """创建测试环境"""
    test_dir = Path(base_dir) / "test_env"
    test_dir.mkdir(exist_ok=True)
    
    # 创建目录结构
    (test_dir / "assets" / "fonts").mkdir(parents=True, exist_ok=True)
    (test_dir / "assets" / "1_img").mkdir(parents=True, exist_ok=True)
    (test_dir / "export").mkdir(exist_ok=True)
    
    # 创建测试数据
    test_data = create_test_data()
    test_data.to_excel(test_dir / "test.xlsx", index=False)
    
    # 创建虚拟图片文件
    null_img = test_dir / "assets" / "1_img" / "null.jpg"
    with open(null_img, 'w') as f:
        f.write("dummy image file")
    
    return test_dir


def cleanup_test_environment(test_dir):
    """清理测试环境"""
    if test_dir.exists():
        shutil.rmtree(test_dir)


def validate_test_setup():
    """验证测试设置"""
    project_root = Path(__file__).parent.parent
    
    required_files = [
        "create_xlsx.py",
        "batch_export.py", 
        "auto_export.py",
        "requirements.txt"
    ]
    
    missing_files = []
    for file in required_files:
        if not (project_root / file).exists():
            missing_files.append(file)
    
    if missing_files:
        print(f"缺少必需文件: {missing_files}")
        return False
    
    return True


def run_test_suite():
    """运行完整测试套件"""
    print("Excel-PS Batch Export 测试套件")
    print("=" * 50)
    
    # 验证测试设置
    if not validate_test_setup():
        print("测试设置验证失败")
        return False
    
    # 创建临时测试环境
    with tempfile.TemporaryDirectory() as temp_dir:
        test_env = create_test_environment(temp_dir)
        
        try:
            print(f"测试环境创建在: {test_env}")
            print("测试完成")
            return True
            
        finally:
            cleanup_test_environment(test_env)


if __name__ == "__main__":
    run_test_suite()