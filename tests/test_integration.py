#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel-PS Batch Export 集成测试
===========================

端到端集成测试，测试完整的程序执行流程和真实场景。
"""

import os
import sys
import subprocess
import tempfile
import shutil
import pytest
import time
import importlib.util
from pathlib import Path
import pandas as pd

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

class TestIntegration:
    """端到端集成测试"""
    
    def setup_method(self):
        """每个测试方法前的设置"""
        self.test_dir = tempfile.mkdtemp()
        self.original_cwd = os.getcwd()
        os.chdir(self.test_dir)
        
        # 复制必要的测试文件
        project_root = Path(__file__).parent.parent
        self.copy_test_files(project_root)
    
    def teardown_method(self):
        """每个测试方法后的清理"""
        os.chdir(self.original_cwd)
        shutil.rmtree(self.test_dir, ignore_errors=True)
    
    def copy_test_files(self, project_root):
        """复制测试文件到临时目录"""
        # 创建必要的目录结构
        Path("assets/fonts").mkdir(parents=True, exist_ok=True)
        Path("assets/1_img").mkdir(parents=True, exist_ok=True)
        Path("export").mkdir(parents=True, exist_ok=True)
        
        # 复制字体文件
        font_src = project_root / "assets/fonts/AlibabaPuHuiTi-2-85-Bold.ttf"
        if font_src.exists():
            shutil.copy2(font_src, "assets/fonts/")
        
        # 复制示例图片
        img_src = project_root / "assets/1_img/null.jpg"
        if img_src.exists():
            shutil.copy2(img_src, "assets/1_img/")
        
        # 复制PSD和Excel文件（如果存在）
        for file_name in ["1.psd", "1.xlsx", "2.psd", "2.xlsx", "3#1.psd", "3#2.psd", "3.xlsx"]:
            file_src = project_root / file_name
            if file_src.exists():
                shutil.copy2(file_src, ".")
    
    def test_program_startup_basic(self):
        """测试程序基本启动功能"""
        # 测试不提供参数时的启动
        result = subprocess.run([
            sys.executable, 
            os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "psd_renderer.py")
        ], capture_output=True, text=True, timeout=30)
        
        # 程序应该因为缺少参数而退出，但不应该因为代码错误而崩溃
        assert result.returncode != 0  # 应该有错误退出
        assert "ValueError" not in result.stderr  # 不应该有ValueError
        assert "UnicodeEncodeError" not in result.stderr  # 不应该有编码错误
    
    def test_program_initialization(self):
        """测试程序初始化流程"""
        # 创建最小测试环境
        with open("test_font.ttf", "w") as f:
            f.write("dummy font file")
        
        # 测试程序启动但不执行完整导出
        result = subprocess.run([
            sys.executable,
            os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "psd_renderer.py"),
            "nonexistent",  # 不存在的文件名，应该会失败但不会崩溃
            "test_font.ttf",
            "jpg"
        ], capture_output=True, text=True, timeout=30)
        
        # 检查是否有日期格式错误
        assert "Invalid format string" not in result.stderr
        
        # 检查是否有编码错误
        assert "UnicodeEncodeError" not in result.stderr
        
        # 检查是否有其他初始化错误
        assert "datetime" not in result.stderr.lower()
    
    def test_datetime_format_handling(self):
        """测试日期时间格式处理"""
        # 导入相关模块测试日期格式
        from datetime import datetime
        
        # 测试修复后的日期格式
        try:
            current_datetime = datetime.now().strftime('%Y%m%d_%H%M%S')
            # 如果能正常执行，说明日期格式是正确的
            assert len(current_datetime) == 15  # YYYYMMDD_HHMMSS 格式
            assert current_datetime[8] == '_'  # 分隔符正确
        except ValueError as e:
            pytest.fail(f"日期格式错误: {e}")
    
    def test_command_line_argument_parsing(self):
        """测试命令行参数解析"""
        # 创建测试字体文件
        with open("test.ttf", "w") as f:
            f.write("dummy font")
        
        # 测试不同参数组合
        test_cases = [
            ["test", "test.ttf", "jpg"],
            ["test", "test.ttf", "png"],
            ["long_name", "test.ttf", "jpg"],
        ]
        
        for args in test_cases:
            result = subprocess.run([
                sys.executable,
                os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "psd_renderer.py")
            ] + args, capture_output=True, text=True, timeout=30)
            
            # 不应该因为参数解析而崩溃
            assert "IndexError" not in result.stderr
            assert "TypeError" not in result.stderr
    
    def test_file_path_handling(self):
        """测试文件路径处理"""
        # 创建测试文件
        with open("test.ttf", "w") as f:
            f.write("dummy font")
        
        # 测试相对路径
        result = subprocess.run([
            sys.executable,
            os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "psd_renderer.py"),
            "test",
            "./test.ttf",
            "jpg"
        ], capture_output=True, text=True, timeout=30)
        
        # 不应该因为路径问题而崩溃
        assert "FileNotFoundError" in result.stderr or result.returncode != 0
        
        # 测试绝对路径
        abs_font_path = os.path.abspath("test.ttf")
        result = subprocess.run([
            sys.executable,
            os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "psd_renderer.py"),
            "test",
            abs_font_path,
            "jpg"
        ], capture_output=True, text=True, timeout=30)
        
        # 不应该因为路径问题而崩溃
        assert "AttributeError" not in result.stderr
    
    def test_error_handling_startup(self):
        """测试启动时的错误处理"""
        # 测试不存在的字体文件
        result = subprocess.run([
            sys.executable,
            os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "psd_renderer.py"),
            "test",
            "nonexistent_font.ttf",
            "jpg"
        ], capture_output=True, text=True, timeout=30)
        
        # 应该优雅地处理错误，而不是崩溃
        assert "Traceback" not in result.stdout
        assert result.returncode != 0
    
    def test_psd_renderer_script_exists(self):
        """测试批量导出脚本是否存在且可执行"""
        script_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "psd_renderer.py")
        
        # 检查脚本文件是否存在
        assert os.path.exists(script_path), f"批量导出脚本不存在: {script_path}"
        
        # 检查脚本文件是否可读
        assert os.access(script_path, os.R_OK), f"批量导出脚本不可读: {script_path}"
        
        # 检查脚本是否有基本的Python结构
        with open(script_path, 'r', encoding='utf-8') as f:
            content = f.read()
            assert 'def psd_renderer_images' in content, "缺少psd_renderer_images函数"
            assert 'if __name__ == "__main__"' in content, "缺少主程序入口点"
    
    def test_required_dependencies(self):
        """测试必要的依赖是否可用"""
        # 测试关键依赖的导入
        dependencies = [
            'psd_tools',
            'pandas', 
            'PIL',
            'tqdm',
            'datetime',
            'multiprocessing'
        ]
        
        for dep in dependencies:
            try:
                __import__(dep)
            except ImportError as e:
                pytest.fail(f"缺少必要的依赖: {dep} - {e}")
    
    def test_program_structure(self):
        """测试程序结构完整性"""
        script_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "psd_renderer.py")
        
        with open(script_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 检查关键函数是否存在
        required_functions = [
            'read_excel_file',
            'calculate_text_position',
            'update_text_layer',
            'update_image_layer',
            'validate_data',
            'psd_renderer_images'
        ]
        
        for func in required_functions:
            assert f'def {func}' in content, f"缺少关键函数: {func}"
        
        # 检查是否有修复过的问题
        assert '%Y%0m%d_%H%M%S' not in content, "仍存在错误的日期格式字符串"
        
        # 检查是否还有emoji字符
        emoji_chars = ['📁', '🔍', '🔄', '🚀', '💡', '📊', '⚠️', '❌', '✅']
        for emoji in emoji_chars:
            assert emoji not in content, f"仍存在emoji字符: {emoji}"
    
    def test_main_function_logic(self):
        """测试主函数逻辑"""
        # 导入主模块测试关键逻辑
        script_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "psd_renderer.py")
        
        # 测试能够导入模块
        try:
            spec = importlib.util.spec_from_file_location("psd_renderer", script_path)
            module = importlib.util.module_from_spec(spec)
            
            # 模拟sys.argv以避免导入时执行
            original_argv = sys.argv
            sys.argv = ['psd_renderer.py', 'test', 'test.ttf', 'jpg']
            
            try:
                spec.loader.exec_module(module)
                
                # 检查关键变量和函数是否存在
                assert hasattr(module, 'psd_renderer_images'), "缺少psd_renderer_images函数"
                assert hasattr(module, 'read_excel_file'), "缺少read_excel_file函数"
                
            finally:
                sys.argv = original_argv
                
        except Exception as e:
            pytest.fail(f"无法导入主模块: {e}")

if __name__ == "__main__":
    pytest.main([__file__, "-v", "-s"])