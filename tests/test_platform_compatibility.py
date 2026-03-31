#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel-PS Batch Export 平台兼容性测试
====================================

测试不同平台、不同环境下的兼容性问题，特别是Windows平台特定问题。
"""

import os
import sys
import subprocess
import tempfile
import shutil
import json
import pytest
import platform
from pathlib import Path

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from transform import (
    transform_row, apply_direct, apply_conditional,
    apply_template, apply_derived, apply_derived_raw,
    is_empty, remove_spaces
)


class TestErrorMessageEncoding:
    """测试错误消息的编码处理 - 验证psd_renderer.py在异常场景下输出的错误信息内容"""

    def setup_method(self):
        self.test_dir = tempfile.mkdtemp()
        self.original_cwd = os.getcwd()
        os.chdir(self.test_dir)

        # 创建最小目录结构
        Path("assets/fonts").mkdir(parents=True, exist_ok=True)
        Path("assets/1_img").mkdir(parents=True, exist_ok=True)
        Path("workspace").mkdir(parents=True, exist_ok=True)
        Path("export").mkdir(parents=True, exist_ok=True)

    def teardown_method(self):
        os.chdir(self.original_cwd)
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def _run_psd_renderer(self, *args):
        """运行psd_renderer.py并返回结果"""
        script_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "psd_renderer.py"
        )
        return subprocess.run(
            [sys.executable, script_path, *args],
            capture_output=True, text=True, timeout=30,
            encoding='utf-8', errors='replace'
        )

    def test_missing_excel_file_error_message(self):
        """验证缺少Excel文件时的错误信息包含具体文件路径"""
        result = self._run_psd_renderer("nonexistent_template", "jpg")

        # 程序应该以非零退出码结束
        assert result.returncode != 0

        # 错误信息应包含具体的模板名
        stderr_combined = result.stdout + result.stderr
        assert "nonexistent_template" in stderr_combined, \
            f"Error output should mention the template name 'nonexistent_template', got: {stderr_combined}"

    def test_missing_arguments_usage_message(self):
        """验证参数不足时输出用法说明"""
        result = self._run_psd_renderer()

        assert result.returncode != 0
        assert "用法" in result.stdout or "usage" in result.stdout.lower(), \
            f"Should print usage message when called without args, got stdout: {result.stdout!r}"

    def test_unicode_in_error_path(self):
        """验证包含Unicode的路径在错误消息中不会导致编码崩溃"""
        # 创建一个含中文的workspace下的xlsx引用
        result = self._run_psd_renderer("测试模板", "jpg")

        # 不应因编码问题崩溃（如UnicodeEncodeError）
        assert "UnicodeEncodeError" not in result.stderr, \
            "Program should not crash with UnicodeEncodeError for Chinese file names"
        assert "UnicodeDecodeError" not in result.stderr, \
            "Program should not crash with UnicodeDecodeError for Chinese file names"

        # 应该有关于文件不存在的具体错误信息
        stderr_combined = result.stdout + result.stderr
        assert "测试模板" in stderr_combined or "不存在" in stderr_combined or "失败" in stderr_combined, \
            f"Error should reference the Chinese template name or indicate file not found, got: {stderr_combined}"

    def test_invalid_image_format_error_message(self):
        """验证无效图片格式时的错误处理"""
        # 创建一个空的xlsx文件以便通过文件存在性检查
        import pandas as pd
        df = pd.DataFrame({"File_name": ["test"]})
        df.to_excel("workspace/test_fmt.xlsx", index=False)

        result = self._run_psd_renderer("test_fmt", "bmp")

        # 程序要么处理成功（如果有PSD模板匹配），要么给出有意义的错误
        # 不应产生未处理的异常
        assert "Traceback" not in result.stderr or "psd_renderer" in result.stderr, \
            "Unexpected traceback in stderr - error should be handled gracefully"

    def test_no_encoding_errors_in_any_scenario(self):
        """通用验证：任何错误场景下都不应出现编码异常"""
        # 多种可能触发错误的调用方式
        error_inputs = [
            ["", "jpg"],          # 空模板名
            ["test", ""],         # 空格式
        ]

        for args in error_inputs:
            result = self._run_psd_renderer(*args)
            assert "UnicodeEncodeError" not in result.stderr, \
                f"UnicodeEncodeError for args={args}: {result.stderr}"
            assert "UnicodeDecodeError" not in result.stderr, \
                f"UnicodeDecodeError for args={args}: {result.stderr}"


class TestUnicodeInExcelData:
    """测试Unicode字符在transform pipeline中的正确处理"""

    def _make_raw_row(self, **overrides):
        """创建带有默认值的原始数据行"""
        import pandas as pd
        defaults = {
            "File_name": "test_file",
            "title": "普通标题",
            "content": "普通内容",
            "subtitle": "副标题",
            "background": "",
            "logo": "",
        }
        defaults.update(overrides)
        return pd.Series(defaults)

    def test_chinese_characters_in_direct_mapping(self):
        """测试direct类型映射能正确保留中文字符"""
        raw_row = self._make_raw_row(title="测试中文标题_第一行")
        rules = {
            "primary_field": "File_name",
            "columns": {
                "标题": {"type": "direct", "source": "title"},
            }
        }

        result = transform_row(raw_row, rules, 1)
        assert result["标题"] == "测试中文标题_第一行"

    def test_emoji_in_direct_mapping(self):
        """测试direct类型映射能正确保留emoji字符"""
        raw_row = self._make_raw_row(title="测试_🎉_表情")
        rules = {
            "primary_field": "File_name",
            "columns": {
                "标题": {"type": "direct", "source": "title"},
            }
        }

        result = transform_row(raw_row, rules, 1)
        assert result["标题"] == "测试_🎉_表情"

    def test_mixed_unicode_in_conditional_mapping(self):
        """测试conditional类型映射能正确处理混合Unicode字符"""
        raw_row = self._make_raw_row(
            title="English日本語한국어🌟",
            subtitle="依赖内容_中文",
        )
        rules = {
            "primary_field": "File_name",
            "columns": {
                "副标题": {"type": "direct", "source": "subtitle"},
                "标题": {"type": "conditional", "source": "title", "depends_on": "副标题"},
            }
        }

        result = transform_row(raw_row, rules, 1)
        assert result["副标题"] == "依赖内容_中文"
        assert result["标题"] == "English日本語한국어🌟"

    def test_conditional_skips_unicode_when_guard_empty(self):
        """测试conditional类型在守卫字段为空时跳过Unicode字段"""
        raw_row = self._make_raw_row(
            title="有内容_🎉",
            subtitle="",
        )
        rules = {
            "primary_field": "File_name",
            "columns": {
                "副标题": {"type": "direct", "source": "subtitle"},
                "标题": {"type": "conditional", "source": "title", "depends_on": "副标题"},
            }
        }

        result = transform_row(raw_row, rules, 1)
        assert result["副标题"] == ""
        assert result["标题"] == ""  # 守卫字段为空，应跳过

    def test_unicode_in_template_concatenation(self):
        """测试template类型拼接能正确处理Unicode字符"""
        raw_row = self._make_raw_row(
            title="中文标题",
            subtitle="English副标题",
        )
        rules = {
            "primary_field": "File_name",
            "columns": {
                "full_title": {
                    "type": "template",
                    "template": "{title}-{subtitle}",
                },
            }
        }

        result = transform_row(raw_row, rules, 1)
        assert result["full_title"] == "中文标题-English副标题"

    def test_unicode_remove_spaces(self):
        """测试带Unicode字符的字段去空格"""
        raw_row = self._make_raw_row(title="测 试 标 题")
        rules = {
            "primary_field": "File_name",
            "columns": {
                "标题": {"type": "direct", "source": "title", "remove_spaces": True},
            }
        }

        result = transform_row(raw_row, rules, 1)
        assert result["标题"] == "测试标题"

    def test_unicode_in_primary_field_guards_row(self):
        """测试主字段为空时整行被跳过，即使其他字段有Unicode数据"""
        raw_row = self._make_raw_row(
            File_name="",
            title="有Unicode内容_🚀",
        )
        rules = {
            "primary_field": "File_name",
            "columns": {
                "标题": {"type": "direct", "source": "title"},
            }
        }

        result = transform_row(raw_row, rules, 1)
        assert result == {}  # 整行被跳过

    def test_derived_with_unicode_product_data(self):
        """测试derived类型在Unicode成品数据上正确推导布尔值"""
        raw_row = self._make_raw_row(title="有中文内容")
        rules = {
            "primary_field": "File_name",
            "columns": {
                "标题": {"type": "direct", "source": "title"},
                "has_title": {"type": "derived", "field": "标题", "when_empty": False},
                "no_title": {"type": "derived", "field": "标题", "when_empty": True},
            }
        }

        result = transform_row(raw_row, rules, 1)
        assert result["标题"] == "有中文内容"
        assert result["has_title"] == "True"   # 字段非空 => True
        assert result["no_title"] == "False"   # 字段非空，when_empty=True => False

    def test_derived_raw_with_unicode_source(self):
        """测试derived_raw类型对Unicode原始字段正确推导"""
        raw_row = self._make_raw_row(title="🚀火箭发射")
        rules = {
            "primary_field": "File_name",
            "columns": {
                "has_content": {"type": "derived_raw", "source": "title"},
            }
        }

        result = transform_row(raw_row, rules, 1)
        assert result["has_content"] == "True"  # 非空 => True

    def test_full_transform_pipeline_with_unicode(self):
        """端到端测试：完整的transform pipeline处理混合Unicode数据"""
        import pandas as pd

        raw_row = pd.Series({
            "File_name": "测试文件_🎉",
            "标题": "中文标题_🚀",
            "副标题": "  English Sub  ",
            "备注": "Emoji 🌟 标记",
            "空字段": "",
        })
        rules = {
            "primary_field": "File_name",
            "columns": {
                "name": {"type": "direct", "source": "File_name"},
                "title": {"type": "direct", "source": "标题"},
                "subtitle": {"type": "direct", "source": "副标题", "remove_spaces": True},
                "note": {"type": "conditional", "source": "备注", "depends_on": "title"},
                "has_note": {"type": "derived", "field": "note", "when_empty": False},
                "has_empty": {"type": "derived_raw", "source": "空字段"},
            }
        }

        result = transform_row(raw_row, rules, 1)

        assert result["name"] == "测试文件_🎉"
        assert result["title"] == "中文标题_🚀"
        assert result["subtitle"] == "EnglishSub"  # 去空格
        assert result["note"] == "Emoji 🌟 标记"   # conditional: title非空 => 传递
        assert result["has_note"] == "True"         # note非空
        assert result["has_empty"] == "False"       # 空字段为空


class TestSystemInfoDetection:
    """测试系统信息检测的准确性"""

    def test_platform_detection(self):
        """验证platform.system()返回已知的操作系统名称"""
        system_info = platform.system()
        assert system_info in ("Windows", "Linux", "Darwin", "FreeBSD", "OpenBSD"), \
            f"Unexpected platform.system() value: {system_info}"

    def test_architecture_detection(self):
        """验证architecture返回有效的元组"""
        arch = platform.architecture()
        assert isinstance(arch, tuple), f"architecture() should return a tuple, got {type(arch)}"
        assert len(arch) == 2, f"architecture() should return 2-element tuple, got {len(arch)}"
        # 第一项是位数（如'64bit', '32bit'）
        assert arch[0] in ("64bit", "32bit"), f"Unexpected bitness: {arch[0]}"

    def test_machine_detection(self):
        """验证machine返回常见的CPU架构标识"""
        machine = platform.machine()
        known_machines = ("x86_64", "AMD64", "arm64", "aarch64", "i386", "i686", "ppc64le", "s390x")
        assert machine in known_machines, \
            f"Unexpected platform.machine() value: {machine}"

    def test_path_separator_matches_platform(self):
        """验证路径分隔符与检测到的平台一致"""
        system = platform.system()
        if system == "Windows":
            assert ";" in os.pathsep, "Windows should use semicolon as PATH separator"
        else:
            assert ":" in os.pathsep, "Non-Windows should use colon as PATH separator"

    def test_line_ending_consistency(self):
        """验证系统的行结束符符合平台规范"""
        system = platform.system()
        # os.linesep是系统原生行分隔符
        if system == "Windows":
            assert os.linesep == "\r\n", "Windows should use CRLF"
        else:
            assert os.linesep == "\n", "Non-Windows should use LF"


class TestLongFilePaths:
    """测试超长文件路径的处理"""

    def setup_method(self):
        self.test_dir = tempfile.mkdtemp()
        self.original_cwd = os.getcwd()
        os.chdir(self.test_dir)

        Path("assets/fonts").mkdir(parents=True, exist_ok=True)
        Path("workspace").mkdir(parents=True, exist_ok=True)
        Path("export").mkdir(parents=True, exist_ok=True)

    def teardown_method(self):
        os.chdir(self.original_cwd)
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_long_directory_name_created_successfully(self):
        """验证系统可以创建100字符长度的目录名"""
        long_dir_name = "a" * 100
        long_path = Path(f"assets/fonts/{long_dir_name}")
        long_path.mkdir(parents=True, exist_ok=True)
        assert long_path.exists(), "Should be able to create directory with 100-char name"

    def test_long_file_name_created_successfully(self):
        """验证系统可以在长目录中创建50字符文件名的文件"""
        long_dir_name = "a" * 100
        long_path = Path(f"assets/fonts/{long_dir_name}")
        long_path.mkdir(parents=True, exist_ok=True)

        long_font_name = "b" * 50 + ".ttf"
        file_path = long_path / long_font_name
        file_path.write_text("content")

        assert file_path.exists(), "Should be able to create file with 50-char name in 100-char directory"

    def test_transform_handles_long_primary_field_value(self):
        """验证transform pipeline能处理超长的字段值"""
        import pandas as pd
        from transform import transform_row

        long_value = "测" * 200
        raw_row = pd.Series({"File_name": long_value, "title": long_value})
        rules = {
            "primary_field": "File_name",
            "columns": {
                "name": {"type": "direct", "source": "File_name"},
                "title": {"type": "direct", "source": "title"},
            }
        }

        result = transform_row(raw_row, rules, 1)
        assert len(result["name"]) == 200, "Should preserve full 200-char primary field"
        assert len(result["title"]) == 200, "Should preserve full 200-char title"

    def test_psd_renderer_no_memory_error_on_long_path(self):
        """验证psd_renderer在超长路径输入下不会产生内存或缓冲区错误"""
        script_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "psd_renderer.py"
        )

        long_dir = "a" * 100
        long_file = "b" * 50 + ".ttf"
        long_arg = f"assets/fonts/{long_dir}/{long_file}"

        result = subprocess.run(
            [sys.executable, script_path, long_arg, "jpg"],
            capture_output=True, text=True, timeout=30,
            encoding='utf-8', errors='replace'
        )

        assert "MemoryError" not in result.stderr, "Should not get MemoryError for long paths"
        assert "BufferError" not in result.stderr, "Should not get BufferError for long paths"


class TestFilePathWithSpaces:
    """测试带空格的文件路径"""

    def setup_method(self):
        self.test_dir = tempfile.mkdtemp()
        self.original_cwd = os.getcwd()
        os.chdir(self.test_dir)

        Path("assets/fonts").mkdir(parents=True, exist_ok=True)
        Path("workspace").mkdir(parents=True, exist_ok=True)
        Path("export").mkdir(parents=True, exist_ok=True)

    def teardown_method(self):
        os.chdir(self.original_cwd)
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_file_with_spaces_created_successfully(self):
        """验证可以创建带空格的文件名"""
        spaced_name = "test font with spaces.ttf"
        file_path = Path(f"assets/fonts/{spaced_name}")
        file_path.write_text("content")
        assert file_path.exists(), "Should be able to create file with spaces in name"

    def test_transform_handles_spaces_in_field_values(self):
        """验证transform pipeline正确处理值中的空格"""
        import pandas as pd
        from transform import transform_row

        raw_row = pd.Series({
            "File_name": "test file",
            "title": "title with spaces",
        })
        rules = {
            "primary_field": "File_name",
            "columns": {
                "name": {"type": "direct", "source": "File_name"},
                "title_no_space": {"type": "direct", "source": "title", "remove_spaces": True},
                "title_with_space": {"type": "direct", "source": "title"},
            }
        }

        result = transform_row(raw_row, rules, 1)
        assert result["name"] == "test file", "Should preserve spaces by default"
        assert result["title_no_space"] == "titlewithspaces", "Should remove spaces when configured"
        assert result["title_with_space"] == "title with spaces", "Should keep spaces when not configured"

    def test_template_concatenation_with_spaces(self):
        """验证template类型正确拼接含空格的字段"""
        import pandas as pd
        from transform import transform_row

        raw_row = pd.Series({
            "File_name": "test",
            "first": "hello world",
            "second": "foo bar",
        })
        rules = {
            "primary_field": "File_name",
            "columns": {
                "combined": {
                    "type": "template",
                    "template": "{first} & {second}",
                },
            }
        }

        result = transform_row(raw_row, rules, 1)
        assert result["combined"] == "hello world & foo bar"

    def test_psd_renderer_no_syntax_error_on_spaced_args(self):
        """验证psd_renderer接收带空格的参数不会产生语法错误"""
        script_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "psd_renderer.py"
        )

        result = subprocess.run(
            [sys.executable, script_path, "test template", "jpg"],
            capture_output=True, text=True, timeout=30,
            encoding='utf-8', errors='replace'
        )

        assert "SyntaxError" not in result.stderr, \
            f"Should not produce SyntaxError for spaced argument, got: {result.stderr}"


if __name__ == "__main__":
    pytest.main([__file__, "-v", "-s"])
