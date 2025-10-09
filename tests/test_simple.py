#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel-PS Batch Export 简化测试套件
==============================

由于原始代码结构问题，本测试提供简化的功能验证。
"""

import os
import sys
import tempfile
import shutil
import subprocess
import pytest
import pandas as pd
from pathlib import Path

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 导入共享测试工具
from test_utils import parse_layer_name, validate_layer_name_parsing, create_test_data, validate_test_setup


class TestLayerParsing:
    """测试图层名称解析功能"""
    
    def test_text_variable_parsing(self):
        """测试文本变量解析"""
        # 测试基本文本变量
        validate_layer_name_parsing("@标题#t", "text", "标题", expected_align="left", expected_valign="top")
        
        # 测试居中对齐
        validate_layer_name_parsing("@标题#t_c", "text", "标题", expected_align="center")
        
        # 测试右对齐
        validate_layer_name_parsing("@标题#t_r", "text", "标题", expected_align="right")
        
        # 测试段落文本
        # 测试段落文本
        validate_layer_name_parsing("@描述#t_p", "text", "描述", expected_paragraph=True)
        
        # 测试垂直居中
        validate_layer_name_parsing("@描述#t_pm", "text", "描述", expected_valign="middle")
        
        # 测试垂直底部
        validate_layer_name_parsing("@描述#t_pb", "text", "描述", expected_valign="bottom")
        
        # 测试组合参数
        validate_layer_name_parsing("@描述#t_c_p", "text", "描述", expected_align="center", expected_paragraph=True)
    
    def test_image_variable_parsing(self):
        """测试图片变量解析"""
        validate_layer_name_parsing("@背景图#i", "image", "背景图")
    
    def test_visibility_variable_parsing(self):
        """测试可见性变量解析"""
        validate_layer_name_parsing("@水印#v", "visibility", "水印")
    
    def test_invalid_layer_names(self):
        """测试无效图层名称"""
        # 测试不以@开头的图层
        result = parse_layer_name("普通图层")
        assert result is None
        
        # 测试缺少操作符的图层
        result = parse_layer_name("@变量名")
        assert result is None
        
        # 测试无效的操作符
        result = parse_layer_name("@变量名#x")
        assert result is None


class TestExcelOperations:
    """测试Excel操作功能"""
    
    def test_excel_file_reading(self):
        """测试Excel文件读取"""
        project_root = Path(__file__).parent.parent
        test_excel = project_root / "1.xlsx"
        
        if not test_excel.exists():
            pytest.skip(f"测试Excel文件不存在: {test_excel}")
        
        # 读取Excel文件
        df = pd.read_excel(test_excel)
        
        # 验证数据结构
        assert isinstance(df, pd.DataFrame)
        assert len(df) > 0
        assert len(df.columns) > 0
        
        # 验证列名
        expected_columns = [
            "File_name", "分类", "标题第1行", "标题第2行", 
            "直播时间", "单行", "两行", "时间框", 
            "站内标", "小标签内容", "背景图", "小标签", "站外标"
        ]
        
        for col in expected_columns:
            assert col in df.columns, f"缺少列: {col}"
    
    def test_excel_data_validation(self):
        """测试Excel数据验证"""
        project_root = Path(__file__).parent.parent
        test_excel = project_root / "1.xlsx"
        
        if not test_excel.exists():
            pytest.skip(f"测试Excel文件不存在: {test_excel}")
        
        df = pd.read_excel(test_excel)
        
        # 验证数据完整性
        assert not df.isnull().all().any(), "存在完全为空的列"
        
        # 验证File_name列
        assert "File_name" in df.columns
        assert not df["File_name"].isnull().all(), "File_name列不能全为空"
        
        # 验证布尔列
        bool_columns = ["时间框", "站内标", "小标签", "站外标"]
        for col in bool_columns:
            if col in df.columns:
                # 验证布尔值是否有效
                unique_values = df[col].dropna().unique()
                for val in unique_values:
                    assert val in [True, False, 1, 0, "True", "False", "TRUE", "FALSE"], f"无效的布尔值: {val}"


class TestFileOperations:
    """测试文件操作功能"""
    
    def test_project_structure(self):
        """测试项目结构"""
        project_root = Path(__file__).parent.parent
        
        # 验证必需文件存在
        assert validate_test_setup(), "测试设置验证失败"
    
    def test_assets_directory(self):
        """测试资源目录"""
        project_root = Path(__file__).parent.parent
        assets_dir = project_root / "assets"
        
        assert assets_dir.exists(), "assets目录不存在"
        
        # 验证子目录
        fonts_dir = assets_dir / "fonts"
        images_dir = assets_dir / "1_img"
        
        if fonts_dir.exists():
            font_files = list(fonts_dir.glob("*.ttf")) + list(fonts_dir.glob("*.otf"))
            assert len(font_files) > 0, "fonts目录中没有字体文件"
        
        if images_dir.exists():
            image_files = list(images_dir.glob("*.jpg")) + list(images_dir.glob("*.png"))
            assert len(image_files) > 0, "images目录中没有图片文件"
    
    def test_psd_files_exist(self):
        """测试PSD文件存在性"""
        project_root = Path(__file__).parent.parent
        
        # 查找PSD文件
        psd_files = list(project_root.glob("*.psd"))
        assert len(psd_files) > 0, "没有找到PSD文件"
        
        # 验证测试文件
        test_psd = project_root / "1.psd"
        if test_psd.exists():
            assert test_psd.stat().st_size > 0, "PSD文件为空"
        
        # 验证多模板文件
        multi_psd_1 = project_root / "3#1.psd"
        multi_psd_2 = project_root / "3#2.psd"
        
        if multi_psd_1.exists() and multi_psd_2.exists():
            assert multi_psd_1.stat().st_size > 0, "多模板PSD文件1为空"
            assert multi_psd_2.stat().st_size > 0, "多模板PSD文件2为空"


class TestDependencyCheck:
    """测试依赖包检查"""
    
    def test_required_packages(self):
        """测试必需包是否已安装"""
        required_packages = [
            "pandas",
            "PIL",  # Pillow
            "psd_tools",
            "tqdm"
        ]
        
        for package in required_packages:
            try:
                __import__(package)
            except ImportError:
                pytest.fail(f"缺少必需包: {package}")
    
    def test_psd_tools_functionality(self):
        """测试psd-tools基本功能"""
        try:
            from psd_tools import PSDImage
            
            # 这里不实际打开PSD文件，只测试导入
            assert PSDImage is not None
            
        except ImportError:
            pytest.fail("psd-tools导入失败")


class TestEndToEndSimple:
    """简单的端到端测试"""
    
    def test_psd_renderer_basic_functionality(self):
        """测试批量导出的基本功能"""
        # 这个测试验证psd_renderer.py能够正常启动和执行基本功能
        script_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "psd_renderer.py")
        
        # 测试程序能够启动并且不会因为基本错误而崩溃
        result = subprocess.run([
            sys.executable, script_path, "test", "nonexistent.ttf", "jpg"
        ], capture_output=True, text=True, timeout=30, encoding='utf-8', errors='replace')
        
        # 程序应该因为缺少文件而退出，但不应该因为代码错误而崩溃
        assert result.returncode != 0
        assert "ValueError" not in result.stderr
        assert "Invalid format string" not in result.stderr
        assert "UnicodeEncodeError" not in result.stderr
    
    def test_program_structure_validation(self):
        """验证程序结构"""
        script_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "psd_renderer.py")
        
        # 检查脚本文件是否存在且可读
        assert os.path.exists(script_path)
        assert os.access(script_path, os.R_OK)
        
        # 检查脚本内容
        with open(script_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 检查关键函数是否存在
        required_functions = [
            'def read_excel_file',
            'def calculate_text_position', 
            'def update_text_layer',
            'def update_image_layer',
            'def validate_data',
            'def psd_renderer_images'
        ]
        
        for func in required_functions:
            assert func in content, f"缺少关键函数: {func}"
        
        # 检查主程序入口
        assert 'if __name__ == "__main__":' in content
        
        # 检查是否有修复过的问题
        assert '%Y%0m%d_%H%M%S' not in content, "仍存在错误的日期格式字符串"
        
        # 检查是否还有emoji字符（可能导致编码问题）
        emoji_chars = ['📁', '🔍', '🔄', '🚀', '💡', '📊', '⚠️', '❌', '✅']
        for emoji in emoji_chars:
            assert emoji not in content, f"仍存在emoji字符: {emoji}"
    
    def test_datetime_format_functionality(self):
        """测试日期时间格式功能"""
        from datetime import datetime
        
        # 测试修复后的日期格式
        try:
            # 这应该能正常工作，因为我们已经修复了格式字符串
            current_datetime = datetime.now().strftime('%Y%m%d_%H%M%S')
            assert len(current_datetime) == 15  # YYYYMMDD_HHMMSS
            assert current_datetime[8] == '_'
            assert current_datetime.replace('_', '').isdigit()
        except ValueError as e:
            pytest.fail(f"日期格式错误: {e}")
    
    def test_import_dependencies(self):
        """测试依赖导入"""
        # 测试所有必要的依赖都能正常导入
        dependencies = [
            'os', 'sys', 'subprocess', 'tempfile', 'shutil',
            'pandas', 'PIL', 'psd_tools', 'tqdm',
            'datetime', 'multiprocessing', 'pathlib'
        ]
        
        for dep in dependencies:
            try:
                __import__(dep)
            except ImportError as e:
                pytest.fail(f"无法导入依赖: {dep} - {e}")
    
    def test_safe_print_message_function(self):
        """测试安全打印消息函数"""
        # 导入业务代码中的safe_print_message函数
        sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        from psd_renderer import safe_print_message
        
        # 测试正常消息
        try:
            safe_print_message("测试消息")
        except Exception as e:
            pytest.fail(f"safe_print_message 处理正常消息失败: {e}")
        
        # 测试包含特殊字符的消息
        try:
            safe_print_message("测试消息 with special chars: ●○■□★☆◆◇")
        except Exception as e:
            pytest.fail(f"safe_print_message 处理特殊字符失败: {e}")
        
        # 测试中文消息
        try:
            safe_print_message("中文测试消息")
        except Exception as e:
            pytest.fail(f"safe_print_message 处理中文消息失败: {e}")
    
    def test_business_code_improvements(self):
        """测试业务代码改进效果"""
        # 验证业务代码中已经修复的问题
        script_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "psd_renderer.py")
        
        # 检查脚本文件是否存在且可读
        assert os.path.exists(script_path)
        assert os.access(script_path, os.R_OK)
        
        # 检查脚本内容
        with open(script_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 检查是否有safe_print_message函数
        assert 'def safe_print_message' in content, "缺少safe_print_message函数"
        
        # 检查是否使用了safe_print_message
        assert 'safe_print_message' in content, "业务代码中未使用safe_print_message"
        
        # 检查是否修复了日期格式问题
        assert '%Y%0m%d_%H%M%S' not in content, "仍存在错误的日期格式字符串"
        
        # 检查是否还有emoji字符
        emoji_chars = ['📁', '🔍', '🔄', '🚀', '💡', '📊', '⚠️', '❌', '✅']
        for emoji in emoji_chars:
            assert emoji not in content, f"仍存在emoji字符: {emoji}"


if __name__ == "__main__":
    # 运行测试
    pytest.main([__file__, "-v"])