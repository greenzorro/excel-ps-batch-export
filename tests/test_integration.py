#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel-PS Batch Export 集成测试
===========================

端到端集成测试，测试完整的程序执行流程和真实场景。
"""

import ast
import inspect
import os
import subprocess
import sys
import tempfile
import shutil
import pytest

from pathlib import Path
from unittest.mock import patch

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 项目根目录
PROJECT_ROOT = Path(__file__).parent.parent
SCRIPT_PATH = PROJECT_ROOT / "psd_renderer.py"


class TestRequiredDependencies:
    """测试必要的依赖是否可用"""

    DEPENDENCIES = [
        'psd_tools',
        'pandas',
        'PIL',
        'tqdm',
        'datetime',
        'multiprocessing'
    ]

    @pytest.mark.parametrize("dep", DEPENDENCIES)
    def test_import_succeeds(self, dep):
        """每个关键依赖都应能成功导入"""
        try:
            __import__(dep)
        except ImportError as e:
            pytest.fail(f"缺少必要的依赖: {dep} - {e}")


class TestPsdRendererScriptExists:
    """测试批量导出脚本是否存在且可执行"""

    def test_script_file_exists(self):
        """脚本文件存在"""
        assert SCRIPT_PATH.exists(), f"批量导出脚本不存在: {SCRIPT_PATH}"

    def test_script_file_readable(self):
        """脚本文件可读"""
        assert os.access(SCRIPT_PATH, os.R_OK), f"批量导出脚本不可读: {SCRIPT_PATH}"

    def test_script_has_main_entry_point(self):
        """脚本包含 if __name__ == '__main__' 入口"""
        content = SCRIPT_PATH.read_text(encoding='utf-8')
        assert 'if __name__ == "__main__"' in content, "缺少主程序入口点"

    def test_script_has_core_function(self):
        """脚本包含 psd_renderer_images 核心函数"""
        content = SCRIPT_PATH.read_text(encoding='utf-8')
        assert 'def psd_renderer_images' in content, "缺少 psd_renderer_images 函数"


class TestDatetimeFormatHandling:
    """测试日期时间格式处理"""

    def test_format_produces_expected_pattern(self):
        """strftime('%Y%m%d_%H%M%S') 产出 YYYYMMDD_HHMMSS 格式"""
        from datetime import datetime

        dt = datetime(2024, 3, 15, 9, 30, 45)
        formatted = dt.strftime('%Y%m%d_%H%M%S')

        assert formatted == "20240315_093045", f"日期格式不正确: {formatted}"

    def test_format_length_and_separator(self):
        """格式化结果固定为 15 个字符，第 9 位为下划线"""
        from datetime import datetime

        formatted = datetime.now().strftime('%Y%m%d_%H%M%S')
        assert len(formatted) == 15, f"日期格式长度不正确: {formatted}"
        assert formatted[8] == '_', f"分隔符不正确: {formatted}"

    def test_format_components_are_valid(self):
        """解析出的年/月/日/时/分/秒 各字段在合理范围内"""
        from datetime import datetime

        formatted = datetime.now().strftime('%Y%m%d_%H%M%S')
        year = int(formatted[0:4])
        month = int(formatted[4:6])
        day = int(formatted[6:8])
        hour = int(formatted[9:11])
        minute = int(formatted[11:13])
        second = int(formatted[13:15])

        assert 2000 <= year <= 2100, f"年份不在合理范围: {year}"
        assert 1 <= month <= 12, f"月份不在合理范围: {month}"
        assert 1 <= day <= 31, f"日期不在合理范围: {day}"
        assert 0 <= hour <= 23, f"小时不在合理范围: {hour}"
        assert 0 <= minute <= 59, f"分钟不在合理范围: {minute}"
        assert 0 <= second <= 59, f"秒不在合理范围: {second}"


class TestProgramStartupBasic:
    """测试程序基本启动行为"""

    def test_no_args_shows_usage_and_exits_nonzero(self):
        """不提供参数时应打印用法说明并以非零退出码退出"""
        result = subprocess.run(
            [sys.executable, str(SCRIPT_PATH)],
            capture_output=True, text=True, timeout=30
        )

        assert result.returncode != 0, "缺少参数时程序应以非零退出码退出"

        combined = result.stdout + result.stderr
        assert "用法" in combined, (
            f"缺少参数时程序应输出用法说明。\nstdout: {result.stdout}\nstderr: {result.stderr}"
        )

    def test_no_args_shows_example_command(self):
        """不提供参数时输出应包含示例命令"""
        result = subprocess.run(
            [sys.executable, str(SCRIPT_PATH)],
            capture_output=True, text=True, timeout=30
        )

        combined = result.stdout + result.stderr
        assert "psd_renderer.py" in combined, (
            "用法说明中应包含 psd_renderer.py 示例"
        )
        assert "jpg" in combined or "png" in combined, (
            "用法说明中应包含输出格式示例 (jpg/png)"
        )


class TestErrorHandlingStartup:
    """测试启动时的错误处理"""

    def test_nonexistent_file_produces_file_not_found_error(self):
        """指定不存在的文件时应产生 FileNotFoundError"""
        result = subprocess.run(
            [sys.executable, str(SCRIPT_PATH), "nonexistent_file", "jpg"],
            capture_output=True, text=True, timeout=30
        )

        assert result.returncode != 0, "文件不存在时程序应以非零退出码退出"

        combined = result.stdout + result.stderr
        assert "FileNotFoundError" in combined or "不存在" in combined, (
            f"文件不存在时应报告 FileNotFoundError 或'不存在'错误。\n"
            f"stdout: {result.stdout}\nstderr: {result.stderr}"
        )

    def test_no_internal_traceback_on_missing_file(self):
        """文件不存在时 stdout 不应包含 Python Traceback 堆栈"""
        result = subprocess.run(
            [sys.executable, str(SCRIPT_PATH), "nonexistent_file", "jpg"],
            capture_output=True, text=True, timeout=30
        )

        assert "Traceback" not in result.stdout, (
            f"stdout 中不应包含未捕获的 Traceback。\nstdout: {result.stdout}"
        )


class TestProgramStructure:
    """测试程序结构完整性 — 通过 AST 分析函数签名"""

    def test_core_functions_have_correct_parameter_counts(self):
        """核心函数的参数数量应符合预期（AST 检查）"""
        source = SCRIPT_PATH.read_text(encoding='utf-8')
        tree = ast.parse(source)

        # 收集所有顶层函数定义
        func_defs = {
            node.name: [arg.arg for arg in node.args.args]
            for node in ast.walk(tree)
            if isinstance(node, ast.FunctionDef)
        }

        # read_excel_file 应接受 file_path 参数
        assert "read_excel_file" in func_defs, "缺少 read_excel_file 函数"
        assert "file_path" in func_defs["read_excel_file"], (
            "read_excel_file 应接受 file_path 参数"
        )

        # calculate_text_position 应接受 6 个参数
        assert "calculate_text_position" in func_defs, "缺少 calculate_text_position 函数"
        expected_args = ["text", "layer_width", "font_size", "alignment", "draw", "font"]
        actual_args = func_defs["calculate_text_position"]
        assert actual_args == expected_args, (
            f"calculate_text_position 参数不匹配，期望: {expected_args}，实际: {actual_args}"
        )

        # update_image_layer 应接受 3 个参数
        assert "update_image_layer" in func_defs, "缺少 update_image_layer 函数"
        assert len(func_defs["update_image_layer"]) == 3, (
            f"update_image_layer 应接受 3 个参数，实际: {func_defs['update_image_layer']}"
        )

        # psd_renderer_images 应是无参函数
        assert "psd_renderer_images" in func_defs, "缺少 psd_renderer_images 函数"
        assert len(func_defs["psd_renderer_images"]) == 0, (
            "psd_renderer_images 应不接受参数（使用全局变量）"
        )

    def test_no_broken_datetime_format_string(self):
        """源码中不应包含已知的错误日期格式 %Y%0m%d"""
        source = SCRIPT_PATH.read_text(encoding='utf-8')
        assert '%Y%0m%d_%H%M%S' not in source, "仍存在错误的日期格式字符串 %Y%0m%d"

    def test_no_emoji_characters_in_source(self):
        """源码中不应包含 emoji 字符"""
        source = SCRIPT_PATH.read_text(encoding='utf-8')
        emoji_chars = ['\U0001F4C1', '\U0001F50D', '\U0001F504', '\U0001F680',
                       '\U0001F4A1', '\U0001F4CA', '\u26A0\uFE0F', '\u274C', '\u2705']
        for emoji in emoji_chars:
            assert emoji not in source, f"源码中仍存在 emoji 字符: {emoji}"


class TestMainFunctionLogic:
    """测试主函数逻辑 — 调用实际函数并验证行为"""

    def test_read_excel_file_raises_on_missing_file(self):
        """read_excel_file 对不存在的文件应抛出 FileNotFoundError"""
        # 先设置 sys.argv 以便模块能导入
        original_argv = sys.argv
        sys.argv = ['psd_renderer.py', 'test', 'jpg']
        try:
            from psd_renderer import read_excel_file

            with pytest.raises(FileNotFoundError, match="Excel文件不存在"):
                read_excel_file("/nonexistent/path/to/file.xlsx")
        finally:
            sys.argv = original_argv

    def test_read_excel_file_raises_on_wrong_extension(self):
        """read_excel_file 对非 Excel 文件应抛出 ValueError"""
        # 创建一个临时文件，但扩展名不是 .xlsx/.xls
        with tempfile.NamedTemporaryFile(suffix=".txt", delete=False) as tmp:
            tmp.write(b"dummy content")
            tmp_path = tmp.name

        original_argv = sys.argv
        sys.argv = ['psd_renderer.py', 'test', 'jpg']
        try:
            from psd_renderer import read_excel_file

            with pytest.raises(ValueError, match="不支持的文件格式"):
                read_excel_file(tmp_path)
        finally:
            sys.argv = original_argv
            os.unlink(tmp_path)

    def test_collect_psd_variables_raises_on_missing_file(self):
        """collect_psd_variables 对不存在的文件应抛出 FileNotFoundError"""
        original_argv = sys.argv
        sys.argv = ['psd_renderer.py', 'test', 'jpg']
        try:
            from psd_renderer import collect_psd_variables

            with pytest.raises(FileNotFoundError, match="PSD文件不存在"):
                collect_psd_variables("/nonexistent/path.psd")
        finally:
            sys.argv = original_argv

    def test_collect_psd_variables_raises_on_wrong_extension(self):
        """collect_psd_variables 对非 PSD 文件应抛出 ValueError"""
        with tempfile.NamedTemporaryFile(suffix=".txt", delete=False) as tmp:
            tmp.write(b"dummy content")
            tmp_path = tmp.name

        original_argv = sys.argv
        sys.argv = ['psd_renderer.py', 'test', 'jpg']
        try:
            from psd_renderer import collect_psd_variables

            with pytest.raises(ValueError, match="文件格式不支持"):
                collect_psd_variables(tmp_path)
        finally:
            sys.argv = original_argv
            os.unlink(tmp_path)

    def test_preprocess_text_handles_none(self):
        """preprocess_text(None) 应返回空字符串"""
        original_argv = sys.argv
        sys.argv = ['psd_renderer.py', 'test', 'jpg']
        try:
            from psd_renderer import preprocess_text

            assert preprocess_text(None) == ""
        finally:
            sys.argv = original_argv

    def test_preprocess_text_strips_quotes(self):
        """preprocess_text 应去除首尾成对的英文引号"""
        original_argv = sys.argv
        sys.argv = ['psd_renderer.py', 'test', 'jpg']
        try:
            from psd_renderer import preprocess_text

            assert preprocess_text('"hello"') == "hello"
            assert preprocess_text('"hello') == '"hello'  # 不成对，不去除
        finally:
            sys.argv = original_argv

    def test_preprocess_text_replaces_slash(self):
        """preprocess_text 应将斜杠替换为 &"""
        original_argv = sys.argv
        sys.argv = ['psd_renderer.py', 'test', 'jpg']
        try:
            from psd_renderer import preprocess_text

            assert "a/b" not in preprocess_text("a/b")
            assert preprocess_text("a/b") == "a&b"
        finally:
            sys.argv = original_argv

    def test_sanitize_filename_removes_illegal_chars(self):
        """sanitize_filename 应清理非法字符"""
        original_argv = sys.argv
        sys.argv = ['psd_renderer.py', 'test', 'jpg']
        try:
            from psd_renderer import sanitize_filename

            assert ":" not in sanitize_filename("test:file")
            assert "*" not in sanitize_filename("test*file")
            assert "?" not in sanitize_filename("test?file")
        finally:
            sys.argv = original_argv

    def test_sanitize_filename_handles_empty_input(self):
        """sanitize_filename 对空输入返回 'unnamed'"""
        original_argv = sys.argv
        sys.argv = ['psd_renderer.py', 'test', 'jpg']
        try:
            from psd_renderer import sanitize_filename

            assert sanitize_filename("") == "unnamed"
            assert sanitize_filename(None) == "unnamed"
        finally:
            sys.argv = original_argv

    def test_parse_rotation_from_name_extracts_angle(self):
        """parse_rotation_from_name 应正确解析旋转角度"""
        original_argv = sys.argv
        sys.argv = ['psd_renderer.py', 'test', 'jpg']
        try:
            from psd_renderer import parse_rotation_from_name

            assert parse_rotation_from_name("@标题#t_a15") == 15.0
            assert parse_rotation_from_name("@标题#t_a-30") == -30.0
            assert parse_rotation_from_name("@标题#t") is None
            assert parse_rotation_from_name("") is None
        finally:
            sys.argv = original_argv

    def test_get_psd_prefix_extracts_correctly(self):
        """get_psd_prefix 应正确提取 PSD 文件名前缀"""
        original_argv = sys.argv
        sys.argv = ['psd_renderer.py', 'test', 'jpg']
        try:
            from psd_renderer import get_psd_prefix

            assert get_psd_prefix("1#海报.psd") == "1"
            assert get_psd_prefix("2.psd") == "2"
            assert get_psd_prefix("产品#横版#v2.psd") == "产品"
        finally:
            sys.argv = original_argv


if __name__ == "__main__":
    pytest.main([__file__, "-v", "-s"])
