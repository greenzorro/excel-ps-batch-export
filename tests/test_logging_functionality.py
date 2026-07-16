#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试日志记录功能

log_export_activity 始终写入项目根目录的 log.csv（与 cwd 无关）。
测试通过临时替换该文件验证行为，结束后恢复原内容。
"""

import os
import sys
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from src.psd_renderer import log_export_activity

LOG_CSV = project_root / "log.csv"


@contextmanager
def isolated_export_log():
    """隔离项目根目录 log.csv，避免污染真实导出日志"""
    backup = LOG_CSV.read_bytes() if LOG_CSV.exists() else None
    if LOG_CSV.exists():
        LOG_CSV.unlink()
    try:
        yield LOG_CSV
    finally:
        if LOG_CSV.exists():
            LOG_CSV.unlink()
        if backup is not None:
            LOG_CSV.write_bytes(backup)


def test_log_export_activity_basic_functionality():
    """测试日志记录基本功能"""
    with isolated_export_log() as log_path:
        log_export_activity("test1.xlsx", 5)

        assert log_path.exists(), "日志文件应该被创建在项目根目录"

        with open(log_path, "r", encoding="utf-8") as f:
            content = f.read()

        lines = content.strip().split("\n")
        assert lines[0] == "生成时间,图片数量,所用Excel文件", "表头格式不正确"
        assert len(lines) == 2, f"应该有2行（表头+数据），实际有{len(lines)}行"

        data_parts = lines[1].split(",")
        assert len(data_parts) == 3, "数据记录应该有3个字段"
        assert data_parts[1] == "5", "图片数量应该为5"
        assert data_parts[2] == "test1.xlsx", "Excel文件名应该为test1.xlsx"

        log_export_activity("test2.xlsx", 10)

        with open(log_path, "r", encoding="utf-8") as f:
            content = f.read()

        lines = content.strip().split("\n")
        assert len(lines) == 3, f"应该有3行（表头+2条数据），实际有{len(lines)}行"

        data_parts = lines[2].split(",")
        assert data_parts[1] == "10", "第二条记录的图片数量应该为10"
        assert data_parts[2] == "test2.xlsx", "第二条记录的Excel文件名应该为test2.xlsx"


def test_log_export_activity_duplicate_logging():
    """测试重复数据多次记录行为"""
    with isolated_export_log() as log_path:
        log_export_activity("test.xlsx", 5)

        with open(log_path, "r", encoding="utf-8") as f:
            initial_lines = len(f.readlines())

        log_export_activity("test.xlsx", 5)

        with open(log_path, "r", encoding="utf-8") as f:
            final_lines = len(f.readlines())

        assert final_lines == initial_lines + 1, (
            f"应该只增加一条记录，初始{initial_lines}，最终{final_lines}"
        )


def test_log_export_activity_zero_count_handling():
    """测试零图片数量处理"""
    with isolated_export_log() as log_path:
        log_export_activity("empty.xlsx", 0)

        with open(log_path, "r", encoding="utf-8") as f:
            content = f.read()

        lines = content.strip().split("\n")
        assert len(lines) == 2, "应该有表头和数据行"

        data_parts = lines[1].split(",")
        assert data_parts[1] == "0", "图片数量应该为0"


def test_log_export_activity_file_format_integrity():
    """测试日志文件格式完整性"""
    with isolated_export_log() as log_path:
        for i in range(3):
            log_export_activity(f"test{i}.xlsx", i * 5)

        with open(log_path, "r", encoding="utf-8") as f:
            content = f.read()

        lines = content.strip().split("\n")
        assert len(lines) == 4, f"应该有4行（表头+3条数据），实际有{len(lines)}行"

        for i, line in enumerate(lines):
            parts = line.split(",")
            assert len(parts) == 3, f"第{i + 1}行应该有3个字段，实际有{len(parts)}个"

        for i in range(1, len(lines)):
            timestamp_str = lines[i].split(",")[0]
            try:
                datetime.strptime(timestamp_str, "%Y-%m-%d %H:%M:%S")
            except ValueError:
                assert False, f"第{i + 1}行的时间戳格式不正确: {timestamp_str}"


def test_log_export_activity_cross_platform_compatibility():
    """测试跨平台兼容性"""
    with isolated_export_log() as log_path:
        log_export_activity("测试文件.xlsx", 5)
        log_export_activity("file with spaces.xlsx", 10)

        with open(log_path, "r", encoding="utf-8") as f:
            content = f.read()

        lines = content.strip().split("\n")
        assert len(lines) == 3, "应该有3行记录"
        assert "测试文件.xlsx" in content, "中文文件名应该被正确记录"
        assert "file with spaces.xlsx" in content, "带空格文件名应该被正确记录"


def test_log_export_activity_serial_simulation():
    """模拟串行场景测试"""
    with isolated_export_log() as log_path:
        records_to_create = 10

        for i in range(records_to_create):
            log_export_activity(f"concurrent{i}.xlsx", i + 1)

        with open(log_path, "r", encoding="utf-8") as f:
            content = f.read()

        lines = content.strip().split("\n")
        assert len(lines) == records_to_create + 1, (
            f"应该有{records_to_create + 1}行记录"
        )

        excel_files = [line.split(",")[2] for line in lines[1:]]
        assert len(excel_files) == len(set(excel_files)), "不应该有重复的Excel文件名"


def test_log_export_activity_ignores_cwd():
    """日志写入位置不依赖当前工作目录"""
    import tempfile

    with isolated_export_log() as log_path:
        original_dir = os.getcwd()
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_log = Path(temp_dir) / "log.csv"
            try:
                os.chdir(temp_dir)
                log_export_activity("cwd_independent.xlsx", 3)
            finally:
                os.chdir(original_dir)

            assert log_path.exists(), "即使 cwd 变化，也应写入项目根目录 log.csv"
            assert not temp_log.exists(), "不应在 cwd 下创建 log.csv"

            with open(log_path, "r", encoding="utf-8") as f:
                content = f.read()
            assert "cwd_independent.xlsx" in content


if __name__ == "__main__":
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
    print("✓ 串行场景测试通过")

    test_log_export_activity_ignores_cwd()
    print("✓ cwd 无关性测试通过")

    print("\n所有日志记录功能测试通过！")
