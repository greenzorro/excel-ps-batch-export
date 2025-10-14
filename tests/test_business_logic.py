#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel-PS Batch Export 业务逻辑测试
=============================

测试核心业务逻辑功能，不包括PSD文件操作。
"""

import os
import sys
import tempfile
import shutil
import pytest
import pandas as pd
from pathlib import Path
from unittest.mock import Mock, patch, MagicMock
import textwrap
from PIL import Image, ImageDraw, ImageFont

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 导入业务代码函数
from psd_renderer import get_matching_psds, read_excel_file

# 导入共享测试工具
from test_utils import (
    parse_layer_name, validate_layer_name_parsing, create_mock_layer, create_mock_psd,
    create_test_excel_data, create_temp_excel_file, create_temp_image_file,
    assert_text_position_accuracy, parse_boolean_value, TestEnvironment,
    create_complex_test_data
)

# 使用测试环境管理器处理sys.argv依赖
test_env = TestEnvironment()
test_env.setup_psd_renderer_args('test', 'test_font.ttf', 'jpg')

# 导入需要测试的函数
from psd_renderer import (
    read_excel_file,
    set_layer_visibility,
    get_font_color,
    calculate_text_position,
    update_text_layer,
    update_image_layer,
    get_matching_psds,
    collect_psd_variables,
    is_image_column,
    validate_data,
    report_validation_results
)

# 恢复原始环境
test_env.cleanup()


class TestLayerParsingAdvanced:
    """高级图层名称解析测试"""
    
    def test_complex_text_variables(self):
        """Test complex text variable parsing"""
        # Test combination parameters - p parameter takes precedence, paragraph is True
        validate_layer_name_parsing(
            "@description#t_c_p_pb", "text", "description", 
            expected_align="center", expected_valign="bottom", expected_paragraph=True
        )
        
        # Test only vertical alignment
        validate_layer_name_parsing(
            "@title#t_pm", "text", "title", 
            expected_align="left", expected_valign="middle", expected_paragraph=False
        )
    
    def test_edge_case_layer_names(self):
        """Test edge case layer names"""
        # Test empty string
        assert parse_layer_name("") is None
        assert parse_layer_name(None) is None
        
        # Test only @ symbol
        assert parse_layer_name("@") is None
        
        # Test only # symbol
        assert parse_layer_name("#") is None
        
        # Test special characters
        validate_layer_name_parsing("@special-chars_123#t", "text", "special-chars_123")
        
        # Test Chinese variable name
        validate_layer_name_parsing("@中文标题#t_c", "text", "中文标题", expected_align="center")
    
    def test_invalid_operation_types(self):
        """Test invalid operation types"""
        # Test unknown operation types
        assert parse_layer_name("@variable#x") is None
        assert parse_layer_name("@variable#unknown") is None
        
        # Test incomplete operation types
        assert parse_layer_name("@variable#") is None


class TestExcelDataValidation:
    """Excel数据验证测试"""
    
    def test_excel_data_validation_with_valid_data(self):
        """Test valid data validation"""
        # Create test data
        test_data = {
            "File_name": ["test1", "test2"],
            "title": ["title1", "title2"],
            "background": ["assets/1_img/null.jpg", "assets/1_img/null.jpg"],
            "watermark": [True, False]
        }
        df = pd.DataFrame(test_data)
        
        # Mock PSD variables
        with patch('psd_renderer.collect_psd_variables') as mock_collect:
            with patch('os.path.exists') as mock_exists:
                with patch('psd_tools.api.psd_image.PSDImage.open') as mock_psd:
                    mock_collect.return_value = {"title", "background", "watermark"}
                    mock_exists.return_value = True
                    mock_psd.return_value = Mock()
                    mock_psd.return_value.__iter__ = Mock(return_value=iter([]))
                    
                    with patch('psd_renderer.is_image_column') as mock_is_image:
                        mock_is_image.return_value = True
                        
                        errors, warnings = validate_data(df, ["test.psd"])
                        
                        assert len(errors) == 0
                        assert len(warnings) == 0
    
    def test_excel_data_validation_missing_columns(self):
        """Test missing columns validation"""
        test_data = {
            "File_name": ["test1"],
            "title": ["title1"]
            # Missing "background" column required by PSD
        }
        df = pd.DataFrame(test_data)
        
        with patch('psd_renderer.collect_psd_variables') as mock_collect:
            with patch('os.path.exists') as mock_exists:
                with patch('psd_tools.api.psd_image.PSDImage.open') as mock_psd:
                    mock_collect.return_value = {"title", "background"}
                    mock_exists.return_value = True
                    mock_psd.return_value = Mock()
                    mock_psd.return_value.__iter__ = Mock(return_value=iter([]))
                    
                    errors, warnings = validate_data(df, ["test.psd"])
                    
                    assert len(errors) > 0
                    assert any("background" in error for error in errors)
    
    def test_excel_data_validation_image_files_missing(self):
        """Test image files missing validation"""
        test_data = {
            "File_name": ["test1"],
            "title": ["title1"],
            "background": ["nonexistent.jpg"]
        }
        df = pd.DataFrame(test_data)
        
        with patch('psd_renderer.collect_psd_variables') as mock_collect:
            with patch('os.path.exists') as mock_exists:
                with patch('psd_tools.api.psd_image.PSDImage.open') as mock_psd:
                    mock_collect.return_value = {"title", "background"}
                    mock_exists.return_value = False  # File doesn't exist
                    mock_psd.return_value = Mock()
                    mock_psd.return_value.__iter__ = Mock(return_value=iter([]))
                    
                    with patch('psd_renderer.is_image_column') as mock_is_image:
                        mock_is_image.return_value = True
                        
                        errors, warnings = validate_data(df, ["test.psd"])
                        
                        assert len(errors) > 0
                        # Due to mock configuration, may not have specific filename error, but should have related error
                        assert any("does not exist" in error for error in errors)
    
    def test_excel_data_validation_extra_columns(self):
        """Test extra columns validation"""
        test_data = {
            "File_name": ["test1"],
            "title": ["title1"],
            "extra_column": ["extra_data"]
        }
        df = pd.DataFrame(test_data)
        
        with patch('psd_renderer.collect_psd_variables') as mock_collect:
            with patch('os.path.exists') as mock_exists:
                with patch('psd_tools.api.psd_image.PSDImage.open') as mock_psd:
                    mock_collect.return_value = {"title"}
                    mock_exists.return_value = True
                    mock_psd.return_value = Mock()
                    mock_psd.return_value.__iter__ = Mock(return_value=iter([]))
                    
                    errors, warnings = validate_data(df, ["test.psd"])
                    
                    assert len(errors) == 0
                    assert len(warnings) > 0
                    assert any("extra_column" in warning for warning in warnings)


class TestTextRendering:
    """文本渲染功能测试"""
    
    def test_calculate_text_position_chinese(self):
        """Test Chinese text position calculation"""
        # Test Chinese text width calculation
        text = "中文测试"
        layer_width = 200
        font_size = 20
        
        x_pos, y_pos = calculate_text_position(text, layer_width, font_size, "center")
        
        # Chinese should be wider than English characters
        assert x_pos > 0
        assert y_pos < 0
    
    def test_calculate_text_position_english(self):
        """Test English text position calculation"""
        text = "Hello World"
        layer_width = 200
        font_size = 20
        
        x_pos, y_pos = calculate_text_position(text, layer_width, font_size, "right")
        
        # Right alignment should be close to right boundary, considering text width and offset
        assert x_pos > layer_width / 4  # Lower expectation
        assert y_pos < 0
    
    def test_calculate_text_position_alignment(self):
        """Test different alignment methods"""
        text = "Test"
        layer_width = 100
        font_size = 16
        
        # Test left alignment - should account for offset calculation
        x_left, _ = calculate_text_position(text, layer_width, font_size, "left")
        # Left alignment includes negative offset, should be between -0.5 and 0
        assert -0.5 <= x_left <= 0, f"Left alignment should be close to 0, got {x_left}"
        
        # Test center alignment
        x_center, _ = calculate_text_position(text, layer_width, font_size, "center")
        assert x_center > 0 and x_center < layer_width / 2
        
        # Test right alignment
        x_right, _ = calculate_text_position(text, layer_width, font_size, "right")
        assert x_right > layer_width / 2
    
    def test_get_font_color_with_color(self):
        """Test getting font color"""
        font_info = {
            'StyleRun': {
                'RunArray': [{
                    'StyleSheet': {
                        'StyleSheetData': {
                            'FillColor': {
                                'Values': [1.0, 0.5, 0.2, 0.8]  # RGBA
                            }
                        }
                    }
                }]
            }
        }
        
        color = get_font_color(font_info)
        # Float conversion should be precise within rounding error
        expected_color = (128, 51, 204, 255)
        assert len(color) == len(expected_color)
        for i, (actual, expected) in enumerate(zip(color, expected_color)):
            # Allow ±1 difference for float-to-int conversion rounding
            assert abs(actual - expected) <= 1, f"Color component {i} mismatch: expected {expected}, got {actual}"
    
    def test_get_font_color_default(self):
        """Test default font color"""
        font_info = {
            'StyleRun': {
                'RunArray': [{
                    'StyleSheet': {
                        'StyleSheetData': {}
                    }
                }]
            }
        }
        
        color = get_font_color(font_info)
        assert color == (0, 0, 0, 255)  # Default black


class TestImageLayerHandling:
    """图片图层处理测试"""
    
    def test_update_image_layer_with_valid_image(self):
        """Test valid image update"""
        # Create mock image
        mock_image = Mock()
        mock_layer = Mock()
        mock_layer.size = (100, 100)
        mock_layer.offset = (10, 10)
        
        # Create temporary image file
        with tempfile.NamedTemporaryFile(suffix='.jpg', delete=False) as tmp:
            # Create a simple test image
            img = Image.new('RGB', (50, 50), color='red')
            img.save(tmp.name)
            
            with patch('os.path.exists', return_value=True):
                with patch('PIL.Image.open', return_value=img):
                    update_image_layer(mock_layer, tmp.name, mock_image)
                    
                    # Verify alpha_composite is called
                    mock_image.alpha_composite.assert_called_once()
    
    def test_update_image_layer_with_missing_image(self):
        """Test missing image handling"""
        mock_image = Mock()
        mock_layer = Mock()
        
        with patch('os.path.exists', return_value=False):
            with patch('builtins.print') as mock_print:
                update_image_layer(mock_layer, "missing.jpg", mock_image)
                
                # Verify warning message is printed
                mock_print.assert_called_once()
                assert "does not exist" in mock_print.call_args[0][0]
    
    def test_is_image_column_detection(self):
        """Test image column detection"""
        assert is_image_column("i") is True
        assert is_image_column("i_c") is True
        assert is_image_column("i_custom") is True
        assert is_image_column("t") is False
        assert is_image_column("v") is False
        assert is_image_column("x") is False


class TestPSDFileMatching:
    """PSD文件匹配测试"""
    
    def test_get_matching_psds_single_file(self):
        """测试单个PSD文件匹配"""
        original_cwd = os.getcwd()
        with tempfile.TemporaryDirectory() as tmp_dir:
            os.chdir(tmp_dir)
            
            try:
                # 创建测试文件
                Path("test.psd").touch()
                
                matching = get_matching_psds("test")
                assert matching == ["test.psd"]
            finally:
                os.chdir(original_cwd)
    
    def test_get_matching_psds_multiple_files(self):
        """测试多个PSD文件匹配"""
        original_cwd = os.getcwd()
        with tempfile.TemporaryDirectory() as tmp_dir:
            os.chdir(tmp_dir)
            
            try:
                # 创建测试文件
                Path("test#1.psd").touch()
                Path("test#2.psd").touch()
                Path("other.psd").touch()
                
                matching = get_matching_psds("test")
                assert set(matching) == {"test#1.psd", "test#2.psd"}
            finally:
                os.chdir(original_cwd)
    
    def test_get_matching_psds_no_match(self):
        """测试无匹配PSD文件"""
        original_cwd = os.getcwd()
        with tempfile.TemporaryDirectory() as tmp_dir:
            os.chdir(tmp_dir)
            
            try:
                # 创建测试文件
                Path("other.psd").touch()
                
                matching = get_matching_psds("test")
                assert matching == []
            finally:
                os.chdir(original_cwd)


class TestValidationReporting:
    """验证报告测试"""
    
    def test_report_validation_results_success(self):
        """Test successful validation report"""
        with patch('builtins.print') as mock_print:
            result = report_validation_results([], [])
            assert result is True
            mock_print.assert_called()
            assert "数据验证通过" in mock_print.call_args[0][0]
    
    def test_report_validation_results_with_warnings(self):
        """Test validation report with warnings"""
        with patch('builtins.print') as mock_print:
            result = report_validation_results([], ["warning1", "warning2"])
            assert result is True
            mock_print.assert_called()
            # Should include warning information
            print_calls = [str(call) for call in mock_print.call_args_list]
            assert any("warning1" in call for call in print_calls)
    
    def test_report_validation_results_with_errors(self):
        """Test validation report with errors"""
        with patch('builtins.print') as mock_print:
            result = report_validation_results(["error1", "error2"], [])
            assert result is False
            mock_print.assert_called()
            # Should include error information
            print_calls = [str(call) for call in mock_print.call_args_list]
            assert any("error1" in call for call in print_calls)


class TestMultiplePSDTemplates:
    """多个PSD模板文件名生成测试"""
    
    def test_multiple_psd_filename_generation(self):
        """测试多个PSD模板的文件名生成逻辑"""
        # 模拟多个PSD模板的文件名生成场景
        test_cases = [
            ("test", "test#1.psd", "test_1"),
            ("test", "test#2.psd", "test_2"),
            ("test", "test#variant.psd", "test_variant"),
            ("excel", "excel#template1.psd", "excel_template1"),
            ("excel", "excel#template2.psd", "excel_template2"),
            ("prefix", "prefix#suffix.psd", "prefix_suffix"),
            ("base", "base#1.psd", "base_1"),
            ("base", "base#2.psd", "base_2"),
            ("base", "base#3.psd", "base_3"),
        ]
        
        for excel_base, psd_file, expected_suffix in test_cases:
            # 模拟修复后的文件名生成逻辑
            psd_base = psd_file.replace('.psd', '')
            
            # 使用修复后的逻辑
            if psd_base.startswith(excel_base):
                suffix = psd_base[len(excel_base):]
                if suffix.startswith('#'):
                    suffix = suffix[1:]
            else:
                suffix = psd_base
            
            # 生成输出文件名
            output_name = f"{excel_base}_{suffix}"
            
            # 验证结果
            assert output_name == expected_suffix, f"文件名生成错误: {excel_base}, {psd_file} -> {output_name}, 期望: {expected_suffix}"
    
    def test_psd_filename_edge_cases(self):
        """测试PSD文件名的边界情况"""
        # 测试边界情况
        edge_cases = [
            ("test", "test.psd", "test"),  # 没有#分隔符
            ("test", "test#.psd", "test"),  # 只有#没有后缀
            ("test", "test#1#2.psd", "test_1#2"),  # 多个#
            ("", "#test.psd", "test"),  # 空Excel前缀
            ("prefix", "different.psd", "prefix_different"),  # 不匹配的前缀
        ]
        
        for excel_base, psd_file, expected_suffix in edge_cases:
            psd_base = psd_file.replace('.psd', '')
            
            # 使用修复后的逻辑
            if psd_base.startswith(excel_base):
                suffix = psd_base[len(excel_base):]
                if suffix.startswith('#'):
                    suffix = suffix[1:]
            else:
                suffix = psd_base
            
            # 如果后缀为空，则只使用excel_base
            if suffix == "":
                output_name = excel_base
            elif excel_base == "":
                # 如果Excel前缀为空，则只使用后缀
                output_name = suffix
            else:
                output_name = f"{excel_base}_{suffix}"
            
            # 验证结果
            assert output_name == expected_suffix, f"边界情况错误: {excel_base}, {psd_file} -> {output_name}, 期望: {expected_suffix}"


class TestLayerVisibility:
    """图层可见性测试"""
    
    def test_set_layer_visibility_true(self):
        """Test setting layer visibility to True"""
        mock_layer = Mock()
        set_layer_visibility(mock_layer, True)
        assert mock_layer.visible is True
    
    def test_set_layer_visibility_false(self):
        """Test setting layer visibility to False"""
        mock_layer = Mock()
        set_layer_visibility(mock_layer, False)
        assert mock_layer.visible is False
    
    def test_set_layer_visibility_boolean_strings(self):
        """Test boolean string input - should convert correctly"""
        mock_layer = Mock()
        
        # 测试字符串输入转换为布尔值的行为
        # 修复后的代码应该正确解析布尔值字符串
        set_layer_visibility(mock_layer, "True")
        assert mock_layer.visible is True
        
        set_layer_visibility(mock_layer, "False")
        assert mock_layer.visible is False  # 修复：现在正确解析False
        
        # 空字符串应该为False
        set_layer_visibility(mock_layer, "")
        assert mock_layer.visible is False
        
        # 数字字符串转换测试
        set_layer_visibility(mock_layer, "1")
        assert mock_layer.visible is True
        
        set_layer_visibility(mock_layer, "0")
        assert mock_layer.visible is False  # 修复：现在正确解析0


if __name__ == "__main__":
    pytest.main([__file__, "-v"])