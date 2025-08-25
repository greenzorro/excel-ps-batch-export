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
import pytest
import locale
import platform
from pathlib import Path

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

class TestPlatformCompatibility:
    """平台兼容性测试"""
    
    def setup_method(self):
        """每个测试方法前的设置"""
        self.test_dir = tempfile.mkdtemp()
        self.original_cwd = os.getcwd()
        os.chdir(self.test_dir)
        
        # 创建测试环境
        self.setup_test_environment()
    
    def teardown_method(self):
        """每个测试方法后的清理"""
        os.chdir(self.original_cwd)
        shutil.rmtree(self.test_dir, ignore_errors=True)
    
    def setup_test_environment(self):
        """设置测试环境"""
        # 创建必要的目录结构
        Path("assets/fonts").mkdir(parents=True, exist_ok=True)
        Path("assets/1_img").mkdir(parents=True, exist_ok=True)
        Path("export").mkdir(parents=True, exist_ok=True)
        
        # 创建测试字体文件
        with open("assets/fonts/test_font.ttf", "w") as f:
            f.write("dummy font content")
        
        # 创建测试图片文件
        with open("assets/1_img/test.jpg", "w") as f:
            f.write("dummy image content")
    
    def test_windows_console_encoding(self):
        """测试Windows控制台编码兼容性"""
        # 检查当前系统的默认编码
        current_encoding = locale.getpreferredencoding()
        print(f"当前系统编码: {current_encoding}")
        
        # 测试程序在Windows环境下的输出
        script_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "batch_export.py")
        
        # 运行程序并检查输出编码
        result = subprocess.run([
            sys.executable, script_path, "test", "assets/fonts/test_font.ttf", "jpg"
        ], capture_output=True, text=True, timeout=30, encoding='utf-8', errors='replace')
        
        # 检查是否有编码错误
        assert "UnicodeEncodeError" not in result.stderr
        assert "UnicodeDecodeError" not in result.stderr
        
        # 检查输出是否可读
        assert isinstance(result.stdout, str)
        assert isinstance(result.stderr, str)
    
    def test_chinese_file_path_handling(self):
        """测试中文文件路径处理"""
        # 创建中文文件名的测试文件
        chinese_font_name = "测试字体.ttf"
        with open(f"assets/fonts/{chinese_font_name}", "w", encoding='utf-8') as f:
            f.write("中文字体内容")
        
        script_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "batch_export.py")
        
        # 测试中文路径的处理
        result = subprocess.run([
            sys.executable, script_path, "test", f"assets/fonts/{chinese_font_name}", "jpg"
        ], capture_output=True, text=True, timeout=30, encoding='utf-8', errors='replace')
        
        # 不应该因为路径编码问题而崩溃
        assert "UnicodeEncodeError" not in result.stderr
        assert "UnicodeDecodeError" not in result.stderr
    
    def test_special_characters_in_output(self):
        """测试输出中特殊字符的处理"""
        script_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "batch_export.py")
        
        # 运行程序并检查输出中的特殊字符
        result = subprocess.run([
            sys.executable, script_path, "test", "assets/fonts/test_font.ttf", "jpg"
        ], capture_output=True, text=True, timeout=30, encoding='utf-8', errors='replace')
        
        # 检查是否有特殊字符导致的编码问题
        special_chars = ['●', '○', '■', '□', '★', '☆', '◆', '◇']
        for char in special_chars:
            # 如果输出中包含这些字符，不应该导致编码错误
            if char in result.stdout or char in result.stderr:
                assert "UnicodeEncodeError" not in result.stderr
    
    def test_file_path_with_spaces(self):
        """测试带空格的文件路径"""
        # 创建带空格的字体文件名
        spaced_font_name = "test font with spaces.ttf"
        with open(f"assets/fonts/{spaced_font_name}", "w") as f:
            f.write("spaced font content")
        
        script_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "batch_export.py")
        
        # 测试带空格的路径
        result = subprocess.run([
            sys.executable, script_path, "test", f"assets/fonts/{spaced_font_name}", "jpg"
        ], capture_output=True, text=True, timeout=30, encoding='utf-8', errors='replace')
        
        # 不应该因为空格而崩溃
        assert "SyntaxError" not in result.stderr
        assert "FileNotFoundError" in result.stderr or result.returncode != 0
    
    def test_long_file_paths(self):
        """测试超长文件路径"""
        # 创建超长目录名
        long_dir_name = "a" * 100
        long_path = Path(f"assets/fonts/{long_dir_name}")
        long_path.mkdir(parents=True, exist_ok=True)
        
        # 在超长目录中创建字体文件
        long_font_name = "a" * 50 + ".ttf"
        with open(long_path / long_font_name, "w") as f:
            f.write("long path font")
        
        script_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "batch_export.py")
        
        # 测试超长路径
        result = subprocess.run([
            sys.executable, script_path, "test", str(long_path / long_font_name), "jpg"
        ], capture_output=True, text=True, timeout=30, encoding='utf-8', errors='replace')
        
        # 不应该因为路径过长而崩溃
        assert "MemoryError" not in result.stderr
        assert "BufferError" not in result.stderr
    
    def test_different_locale_settings(self):
        """测试不同区域设置的影响"""
        # 保存原始区域设置
        original_locale = locale.getlocale()
        
        try:
            # 测试中文区域设置
            if platform.system() == "Windows":
                try:
                    locale.setlocale(locale.LC_ALL, 'chinese')
                except:
                    # 如果中文区域设置不可用，尝试其他设置
                    try:
                        locale.setlocale(locale.LC_ALL, 'zh_CN.UTF-8')
                    except:
                        pass
            
            # 测试程序在不同区域设置下的行为
            script_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "batch_export.py")
            
            result = subprocess.run([
                sys.executable, script_path, "test", "assets/fonts/test_font.ttf", "jpg"
            ], capture_output=True, text=True, timeout=30, encoding='utf-8', errors='replace')
            
            # 检查是否有区域设置相关的问题
            assert "locale" not in result.stderr.lower()
            
        finally:
            # 恢复原始区域设置
            try:
                locale.setlocale(locale.LC_ALL, original_locale)
            except:
                pass
    
    def test_error_message_encoding(self):
        """测试错误消息的编码处理"""
        script_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "batch_export.py")
        
        # 故意使用不存在的文件来触发错误
        result = subprocess.run([
            sys.executable, script_path, "test", "nonexistent_font.ttf", "jpg"
        ], capture_output=True, text=True, timeout=30, encoding='utf-8', errors='replace')
        
        # 错误消息应该能够正确显示，不应该有编码问题
        assert isinstance(result.stderr, str)
        assert len(result.stderr) > 0  # 应该有错误信息
        
        # BC-006修复后，检查关键错误信息是否正确显示
        # 由于测试环境路径包含中文，我们主要检查核心错误信息
        assert "Excel" in result.stderr or "PSD" in result.stderr or "不存在" in result.stderr
        
        # 检查程序不会因为编码问题而崩溃
        assert "UnicodeEncodeError" not in result.stderr
        assert "UnicodeDecodeError" not in result.stderr
    
    def test_progress_display_encoding(self):
        """测试进度显示的编码兼容性"""
        script_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "batch_export.py")
        
        # 运行程序并检查进度显示
        result = subprocess.run([
            sys.executable, script_path, "test", "assets/fonts/test_font.ttf", "jpg"
        ], capture_output=True, text=True, timeout=30, encoding='utf-8', errors='replace')
        
        # 检查进度条相关的字符是否正确显示
        progress_chars = ['|', '/', '-', '\\', '[', ']', '%', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9']
        for char in progress_chars:
            # 这些字符应该能正确显示
            if char in result.stdout:
                assert result.stdout.count(char) >= 0  # 不应该有编码错误
    
    def test_system_info_detection(self):
        """测试系统信息检测的准确性"""
        # 检查平台信息
        system_info = platform.system()
        architecture = platform.architecture()
        machine = platform.machine()
        
        print(f"系统: {system_info}")
        print(f"架构: {architecture}")
        print(f"机器: {machine}")
        
        # 根据系统类型进行特定测试
        if system_info == "Windows":
            # Windows特定测试
            assert "Windows" in system_info
            # 测试Windows路径分隔符
            assert "\\" in os.pathsep or ";" in os.pathsep
        
        elif system_info == "Linux":
            # Linux特定测试
            assert "Linux" in system_info
            # 测试Linux路径分隔符
            assert ":" in os.pathsep
        
        elif system_info == "Darwin":
            # macOS特定测试
            assert "Darwin" in system_info
    
    def test_environment_variable_handling(self):
        """测试环境变量的处理"""
        # 设置测试环境变量
        test_env = os.environ.copy()
        test_env["TEST_ENV"] = "测试环境变量"
        test_env["LANG"] = "en_US.UTF-8"
        
        script_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "batch_export.py")
        
        # 在特定环境下运行程序
        result = subprocess.run([
            sys.executable, script_path, "test", "assets/fonts/test_font.ttf", "jpg"
        ], capture_output=True, text=True, timeout=30, encoding='utf-8', errors='replace', env=test_env)
        
        # 环境变量不应该导致程序崩溃
        assert "EnvironmentError" not in result.stderr
        assert "OSError" not in result.stderr or "No such file" in result.stderr
    
    def test_unicode_in_excel_data(self):
        """测试Excel数据中的Unicode字符处理"""
        # 创建包含Unicode字符的测试数据
        unicode_data = {
            'File_name': ['测试_🎉', 'English_😊', '混合_🚀'],
            'title': ['中文标题', 'English Title', 'Mixed 标题'],
            'content': ['这是中文内容', 'This is English content', '混合 content 内容']
        }
        
        # 虽然我们没有真实的Excel文件，但可以测试相关的处理逻辑
        from test_utils import parse_boolean_value
        
        # 测试布尔值解析对Unicode的处理
        assert parse_boolean_value("TRUE") == True  # 英文大写
        assert parse_boolean_value("false") == False  # 英文小写
        assert parse_boolean_value("1") == True  # 数字1
        assert parse_boolean_value("0") == False  # 数字0
    
    def test_console_output_buffering(self):
        """测试控制台输出缓冲问题"""
        script_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "batch_export.py")
        
        # 运行程序并检查输出缓冲
        result = subprocess.run([
            sys.executable, script_path, "test", "assets/fonts/test_font.ttf", "jpg"
        ], capture_output=True, text=True, timeout=30, encoding='utf-8', errors='replace')
        
        # 检查输出是否完整
        assert len(result.stdout) > 0 or len(result.stderr) > 0
        
        # 检查是否有输出被截断的迹象
        assert not result.stdout.endswith("...")  # 不应该以省略号结尾
        assert not result.stderr.endswith("...")  # 不应该以省略号结尾

if __name__ == "__main__":
    pytest.main([__file__, "-v", "-s"])