#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试日志记录功能

测试 log_export_activity 函数的正确性和健壮性
包括重复记录检测、文件格式验证等
"""

import os
import sys
import tempfile
import pandas as pd
from datetime import datetime
from pathlib import Path

# 添加项目根目录到路径
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from psd_renderer import log_export_activity


def test_log_export_activity_basic_functionality():
    """测试日志记录基本功能"""
    with tempfile.TemporaryDirectory() as temp_dir:
        # 在临时目录中测试
        original_dir = os.getcwd()
        os.chdir(temp_dir)

        try:
            # 第一次调用 - 应该创建文件并写入表头
            log_export_activity("test1.xlsx", 5)

            # 验证文件存在
            assert os.path.exists('log.csv'), "日志文件应该被创建"

            # 读取文件内容
            with open('log.csv', 'r', encoding='utf-8') as f:
                content = f.read()

            # 验证表头
            lines = content.strip().split('\n')
            assert lines[0] == '生成时间,图片数量,所用Excel文件', "表头格式不正确"

            # 验证第一条记录
            assert len(lines) == 2, f"应该有2行（表头+数据），实际有{len(lines)}行"

            # 验证数据格式
            data_parts = lines[1].split(',')
            assert len(data_parts) == 3, "数据记录应该有3个字段"
            assert data_parts[1] == '5', "图片数量应该为5"
            assert data_parts[2] == 'test1.xlsx', "Excel文件名应该为test1.xlsx"

            # 第二次调用 - 应该追加记录
            log_export_activity("test2.xlsx", 10)

            # 验证追加的记录
            with open('log.csv', 'r', encoding='utf-8') as f:
                content = f.read()

            lines = content.strip().split('\n')
            assert len(lines) == 3, f"应该有3行（表头+2条数据），实际有{len(lines)}行"

            # 验证第二条记录
            data_parts = lines[2].split(',')
            assert data_parts[1] == '10', "第二条记录的图片数量应该为10"
            assert data_parts[2] == 'test2.xlsx', "第二条记录的Excel文件名应该为test2.xlsx"

        finally:
            os.chdir(original_dir)


def test_log_export_activity_duplicate_logging():
    """测试重复数据多次记录行为"""
    with tempfile.TemporaryDirectory() as temp_dir:
        original_dir = os.getcwd()
        os.chdir(temp_dir)

        try:
            # 模拟重复调用场景（类似之前的问题）
            log_export_activity("test.xlsx", 5)

            # 读取初始记录数
            with open('log.csv', 'r', encoding='utf-8') as f:
                initial_lines = len(f.readlines())

            # 再次调用相同数据 - 应该只增加一条记录
            log_export_activity("test.xlsx", 5)

            # 验证记录数
            with open('log.csv', 'r', encoding='utf-8') as f:
                final_lines = len(f.readlines())

            assert final_lines == initial_lines + 1, f"应该只增加一条记录，初始{initial_lines}，最终{final_lines}"

        finally:
            os.chdir(original_dir)


def test_log_export_activity_zero_count_handling():
    """测试零图片数量处理"""
    with tempfile.TemporaryDirectory() as temp_dir:
        original_dir = os.getcwd()
        os.chdir(temp_dir)

        try:
            # 测试零图片数量的记录
            log_export_activity("empty.xlsx", 0)

            # 验证记录被正确写入
            with open('log.csv', 'r', encoding='utf-8') as f:
                content = f.read()

            lines = content.strip().split('\n')
            assert len(lines) == 2, "应该有表头和数据行"

            data_parts = lines[1].split(',')
            assert data_parts[1] == '0', "图片数量应该为0"

        finally:
            os.chdir(original_dir)


def test_log_export_activity_file_format_integrity():
    """测试日志文件格式完整性"""
    with tempfile.TemporaryDirectory() as temp_dir:
        original_dir = os.getcwd()
        os.chdir(temp_dir)

        try:
            # 多次写入测试
            for i in range(3):
                log_export_activity(f"test{i}.xlsx", i * 5)

            # 验证文件格式
            with open('log.csv', 'r', encoding='utf-8') as f:
                content = f.read()

            lines = content.strip().split('\n')

            # 验证总行数
            assert len(lines) == 4, f"应该有4行（表头+3条数据），实际有{len(lines)}行"

            # 验证每行的字段数
            for i, line in enumerate(lines):
                parts = line.split(',')
                assert len(parts) == 3, f"第{i+1}行应该有3个字段，实际有{len(parts)}个"

            # 验证时间戳格式
            for i in range(1, len(lines)):
                timestamp_str = lines[i].split(',')[0]
                try:
                    # 尝试解析时间戳
                    datetime.strptime(timestamp_str, '%Y-%m-%d %H:%M:%S')
                except ValueError:
                    assert False, f"第{i+1}行的时间戳格式不正确: {timestamp_str}"

        finally:
            os.chdir(original_dir)


def test_log_export_activity_cross_platform_compatibility():
    """测试跨平台兼容性"""
    with tempfile.TemporaryDirectory() as temp_dir:
        original_dir = os.getcwd()
        os.chdir(temp_dir)

        try:
            # 测试特殊字符文件名
            log_export_activity("测试文件.xlsx", 5)
            log_export_activity("file with spaces.xlsx", 10)

            # 验证文件可读
            with open('log.csv', 'r', encoding='utf-8') as f:
                content = f.read()

            lines = content.strip().split('\n')
            assert len(lines) == 3, "应该有3行记录"

            # 验证特殊字符被正确处理
            assert "测试文件.xlsx" in content, "中文文件名应该被正确记录"
            assert "file with spaces.xlsx" in content, "带空格文件名应该被正确记录"

        finally:
            os.chdir(original_dir)


def test_log_export_activity_serial_simulation():
    """模拟串行场景测试"""
    with tempfile.TemporaryDirectory() as temp_dir:
        original_dir = os.getcwd()
        os.chdir(temp_dir)

        try:
            # 模拟快速连续调用（类似文件监控器的场景）
            records_to_create = 10

            for i in range(records_to_create):
                log_export_activity(f"concurrent{i}.xlsx", i + 1)

            # 验证所有记录都被写入
            with open('log.csv', 'r', encoding='utf-8') as f:
                content = f.read()

            lines = content.strip().split('\n')
            assert len(lines) == records_to_create + 1, f"应该有{records_to_create + 1}行记录"

            # 验证没有重复记录
            excel_files = [line.split(',')[2] for line in lines[1:]]
            assert len(excel_files) == len(set(excel_files)), "不应该有重复的Excel文件名"

        finally:
            os.chdir(original_dir)


if __name__ == "__main__":
    # 运行所有测试
    test_log_export_activity_basic_functionality()
    print("✓ 基本功能测试通过")

    test_log_export_activity_duplicate_logging()
    print("✓ 重复记录行为测试通过")

    test_log_export_activity_zero_count_handling()
    print("✓ 零图片数量处理测试通过")

    test_log_export_activity_file_format_integrity()
    print("✓ 文件格式完整性测试通过")

    test_log_export_activity_cross_platform_compatibility()
    print("✓ 跨平台兼容性测试通过")

    test_log_export_activity_serial_simulation()
    print("✓ 并发场景测试通过")

    print("\n🎉 所有日志记录功能测试通过！")