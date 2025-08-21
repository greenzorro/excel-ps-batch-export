#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel-PS Batch Export 错误处理测试
==============================

测试各种错误场景的处理能力。
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

# Mock sys.argv to avoid command line argument issues
original_argv = sys.argv.copy()
sys.argv = ['batch_export.py', '1', 'test_font.ttf', 'jpg']

# 导入需要测试的函数
from batch_export import (
    read_excel_file,
    validate_data,
    update_text_layer,
    update_image_layer,
    get_matching_psds,
    preload_psd_templates,
    export_single_image_task
)

# Restore original argv
sys.argv = original_argv


class TestExcelFileErrors:
    """Excel文件错误处理测试"""
    
    def test_read_nonexistent_excel_file(self):
        """测试读取不存在的Excel文件"""
        with pytest.raises(FileNotFoundError):
            read_excel_file("nonexistent.xlsx")
    
    def test_read_corrupted_excel_file(self):
        """测试读取损坏的Excel文件"""
        tmp_path = None
        try:
            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
                # 写入一些无效数据
                tmp.write(b"This is not a valid Excel file")
                tmp.flush()
                tmp_path = tmp.name
            
            with pytest.raises(Exception):
                read_excel_file(tmp_path)
        finally:
            if tmp_path and os.path.exists(tmp_path):
                try:
                    os.unlink(tmp_path)
                except PermissionError:
                    pass  # 忽略权限错误，Windows下可能发生
    
    def test_read_empty_excel_file(self):
        """测试读取空Excel文件"""
        tmp_path = None
        try:
            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
                # 创建空的DataFrame
                df = pd.DataFrame()
                df.to_excel(tmp.name, index=False)
                tmp.flush()
                tmp_path = tmp.name
            
            # 应该能读取，但返回空DataFrame
            result = read_excel_file(tmp_path)
            assert isinstance(result, pd.DataFrame)
            assert len(result) == 0
        finally:
            if tmp_path and os.path.exists(tmp_path):
                try:
                    os.unlink(tmp_path)
                except PermissionError:
                    pass  # 忽略权限错误，Windows下可能发生
    
    def test_excel_file_with_missing_columns(self):
        """测试Excel文件缺失必需列"""
        test_data = {
            "Title": ["Test1", "Test2"],  # 缺少File_name列
            "Content": ["Content1", "Content2"]
        }
        df = pd.DataFrame(test_data)
        
        tmp_path = None
        try:
            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
                df.to_excel(tmp.name, index=False)
                tmp.flush()
                tmp_path = tmp.name
            
            # 读取应该成功，但后续验证会失败
            result = read_excel_file(tmp_path)
            assert isinstance(result, pd.DataFrame)
            assert "File_name" not in result.columns
        finally:
            if tmp_path and os.path.exists(tmp_path):
                try:
                    os.unlink(tmp_path)
                except PermissionError:
                    pass  # 忽略权限错误，Windows下可能发生


class TestImageDataErrors:
    """图片数据错误处理测试"""
    
    def test_update_image_layer_with_invalid_path(self):
        """Test invalid image path handling"""
        mock_image = Mock()
        mock_layer = Mock()
        mock_layer.size = (100, 100)
        mock_layer.offset = (0, 0)
        
        # Test non-existent file
        with patch('builtins.print') as mock_print:
            update_image_layer(mock_layer, "nonexistent.jpg", mock_image)
            mock_print.assert_called_once()
            assert "does not exist" in mock_print.call_args[0][0]
    
    def test_update_image_layer_with_unsupported_format(self):
        """Test unsupported image format"""
        mock_image = Mock()
        mock_layer = Mock()
        mock_layer.size = (100, 100)
        mock_layer.offset = (0, 0)
        
        tmp_path = None
        try:
            with tempfile.NamedTemporaryFile(suffix='.xyz', delete=False) as tmp:
                # Write invalid image data
                tmp.write(b"Not an image")
                tmp.flush()
                tmp_path = tmp.name
            
            with patch('os.path.exists', return_value=True):
                # Test unsupported format will throw exception
                with pytest.raises(Exception):
                    update_image_layer(mock_layer, tmp_path, mock_image)
        finally:
            if tmp_path and os.path.exists(tmp_path):
                try:
                    os.unlink(tmp_path)
                except PermissionError:
                    pass  # 忽略权限错误，Windows下可能发生
    
    def test_update_image_layer_with_corrupted_image(self):
        """Test corrupted image file"""
        mock_image = Mock()
        mock_layer = Mock()
        mock_layer.size = (100, 100)
        mock_layer.offset = (0, 0)
        
        tmp_path = None
        try:
            with tempfile.NamedTemporaryFile(suffix='.jpg', delete=False) as tmp:
                # Write corrupted image data
                tmp.write(b"Corrupted image data")
                tmp.flush()
                tmp_path = tmp.name
            
            with patch('os.path.exists', return_value=True):
                # Test corrupted image will throw exception
                with pytest.raises(Exception):
                    update_image_layer(mock_layer, tmp_path, mock_image)
        finally:
            if tmp_path and os.path.exists(tmp_path):
                try:
                    os.unlink(tmp_path)
                except PermissionError:
                    pass  # 忽略权限错误，Windows下可能发生


class TestTextRenderingErrors:
    """文本渲染错误处理测试"""
    
    def test_update_text_layer_with_invalid_font(self):
        """Test invalid font file handling"""
        mock_layer = Mock()
        mock_layer.name = "@title#t"
        mock_layer.engine_dict = {
            'StyleRun': {
                'RunArray': [{
                    'StyleSheet': {
                        'StyleSheetData': {
                            'FontSize': 20
                        }
                    }
                }]
            }
        }
        mock_layer.size = (200, 50)
        mock_layer.offset = (0, 0)
        
        mock_image = Mock()
        
        # Mock ImageFont.truetype to raise OSError
        with patch('batch_export.ImageFont.truetype') as mock_font:
            mock_font.side_effect = OSError("Font not found")
            
            # Test invalid font will throw exception
            with pytest.raises(OSError):
                update_text_layer(mock_layer, "Test text", mock_image)
    
    def test_update_text_layer_with_empty_text(self):
        """Test empty text handling"""
        mock_layer = Mock()
        mock_layer.name = "@title#t"
        mock_layer.engine_dict = {
            'StyleRun': {
                'RunArray': [{
                    'StyleSheet': {
                        'StyleSheetData': {
                            'FontSize': 20
                        }
                    }
                }]
            }
        }
        mock_layer.size = (200, 50)
        mock_layer.offset = (0, 0)
        
        mock_image = Mock()
        
        # Test empty string
        with patch('batch_export.ImageFont.truetype') as mock_font:
            with patch('batch_export.ImageDraw.Draw') as mock_draw:
                update_text_layer(mock_layer, "", mock_image)
                # Should still handle empty string
                mock_font.assert_called_once()
                mock_draw.assert_called_once()
    
    def test_update_text_layer_with_special_characters(self):
        """Test special characters handling"""
        mock_layer = Mock()
        mock_layer.name = "@title#t"
        mock_layer.engine_dict = {
            'StyleRun': {
                'RunArray': [{
                    'StyleSheet': {
                        'StyleSheetData': {
                            'FontSize': 20
                        }
                    }
                }]
            }
        }
        mock_layer.size = (200, 50)
        mock_layer.offset = (0, 0)
        
        mock_image = Mock()
        
        # Test special characters
        with patch('batch_export.ImageFont.truetype') as mock_font:
            with patch('batch_export.ImageDraw.Draw') as mock_draw:
                special_text = "Special chars: @#$%^&*()_+-=[]{}|;':\",./<>?"
                update_text_layer(mock_layer, special_text, mock_image)
                # Should handle special characters
                mock_font.assert_called_once()
                mock_draw.assert_called_once()


class TestPSDTemplateErrors:
    """PSD模板错误处理测试"""
    
    def test_get_matching_psds_with_nonexistent_excel(self):
        """测试不存在的Excel文件"""
        original_cwd = os.getcwd()
        with tempfile.TemporaryDirectory() as tmp_dir:
            try:
                os.chdir(tmp_dir)
                
                # 没有对应的Excel文件
                matching = get_matching_psds("nonexistent.xlsx")
                assert matching == []
            finally:
                os.chdir(original_cwd)
    
    def test_preload_psd_templates_with_missing_files(self):
        """测试预加载缺失的PSD文件"""
        psd_files = ["missing1.psd", "missing2.psd"]
        
        with patch('builtins.print') as mock_print:
            result = preload_psd_templates(psd_files)
            
            # 应该返回包含None的字典
            assert len(result) == 2
            assert result["missing1.psd"] is None
            assert result["missing2.psd"] is None
            
            # 应该打印错误信息
            assert mock_print.called
    
    def test_preload_psd_templates_with_corrupted_files(self):
        """测试预加载损坏的PSD文件"""
        original_cwd = os.getcwd()
        with tempfile.TemporaryDirectory() as tmp_dir:
            try:
                os.chdir(tmp_dir)
                
                # 创建损坏的PSD文件
                corrupted_psd = "corrupted.psd"
                with open(corrupted_psd, 'w') as f:
                    f.write("Corrupted PSD data")
                
                with patch('builtins.print') as mock_print:
                    result = preload_psd_templates([corrupted_psd])
                    
                    # 应该返回None
                    assert result[corrupted_psd] is None
                    
                    # 应该打印错误信息
                    assert mock_print.called
            finally:
                os.chdir(original_cwd)


class TestValidationErrorHandling:
    """验证错误处理测试"""
    
    def test_validate_data_with_missing_psd_files(self):
        """Test validation with missing PSD files"""
        test_data = {
            "File_name": ["test1"],
            "title": ["title1"]
        }
        df = pd.DataFrame(test_data)
        
        # PSD file doesn't exist
        errors, warnings = validate_data(df, ["nonexistent.psd"])
        
        assert len(errors) > 0
        assert any("does not exist" in error for error in errors)
    
    def test_validate_data_with_invalid_boolean_values(self):
        """Test invalid boolean values validation"""
        test_data = {
            "File_name": ["test1"],
            "title": ["title1"],
            "visibility": ["invalid_boolean"]  # Invalid boolean value
        }
        df = pd.DataFrame(test_data)
        
        with patch('batch_export.collect_psd_variables') as mock_collect:
            mock_collect.return_value = {"title", "visibility"}
            
            with patch('batch_export.is_image_column') as mock_is_image:
                mock_is_image.return_value = False
                
                errors, warnings = validate_data(df, ["test.psd"])
                
                # Should have errors because boolean value is invalid
                assert len(errors) > 0
    
    def test_validate_data_with_empty_dataframe(self):
        """Test empty DataFrame validation"""
        df = pd.DataFrame()
        
        with patch('batch_export.collect_psd_variables') as mock_collect:
            with patch('os.path.exists') as mock_exists:
                with patch('psd_tools.api.psd_image.PSDImage.open') as mock_psd:
                    mock_collect.return_value = set()
                    mock_exists.return_value = True
                    mock_psd.return_value = Mock()
                    mock_psd.return_value.__iter__ = Mock(return_value=iter([]))
                    
                    errors, warnings = validate_data(df, ["test.psd"])
                    
                    # Empty DataFrame should pass validation
                    assert len(errors) == 0
                    assert len(warnings) == 0


class TestExportTaskErrors:
    """导出任务错误处理测试"""
    
    def test_export_single_image_task_with_invalid_data(self):
        """Test invalid data export"""
        # Create mock task data
        task_data = {
            'row': pd.Series(["test"], index=["File_name"]),
            'index': 0,
            'psd_object': Mock(),
            'psd_file_name': "test.psd",
            'excel_file_path': "test.xlsx",
            'output_path': "output",
            'image_format': "jpg",
            'text_font': "test.ttf",
            'quality': 95,
            'optimize': False,
            'current_datetime': "20240101_120000"
        }
        
        # Mock PSD object has issues
        task_data['psd_object'].size = (100, 100)
        task_data['psd_object'].is_group.return_value = False
        task_data['psd_object'].is_visible.return_value = True
        task_data['psd_object'].topil.side_effect = Exception("PSD error")
        # Ensure Mock object is iterable
        task_data['psd_object'].__iter__ = Mock(return_value=iter([]))
        
        original_cwd = os.getcwd()
        tmp_dir = None
        try:
            tmp_dir = tempfile.mkdtemp()
            os.chdir(tmp_dir)
            
            # Should handle errors gracefully - either by returning an error result or raising an exception
            try:
                result = export_single_image_task(task_data)
                # If no exception, should return an error indicator (None or error string)
                assert result is None or isinstance(result, str), f"Expected error result, got: {result}"
            except Exception as e:
                # If exception is raised, it should be meaningful
                assert "PSD" in str(e).lower() or "topil" in str(e).lower() or "error" in str(e).lower(), f"Unexpected error: {e}"
        finally:
            os.chdir(original_cwd)
            if tmp_dir and os.path.exists(tmp_dir):
                try:
                    shutil.rmtree(tmp_dir)
                except (PermissionError, OSError):
                    pass  # 忽略权限错误，Windows下可能发生
    
    def test_export_single_image_task_with_permission_error(self):
        """Test permission error handling"""
        # Create mock task data
        task_data = {
            'row': pd.Series(["test"], index=["File_name"]),
            'index': 0,
            'psd_object': Mock(),
            'psd_file_name': "test.psd",
            'excel_file_path': "test.xlsx",
            'output_path': "/root/no_permission",  # Unix系统目录，通常无写入权限
            'image_format': "jpg",
            'text_font': "test.ttf",
            'quality': 95,
            'optimize': False,
            'current_datetime': "20240101_120000"
        }
        
        task_data['psd_object'].size = (100, 100)
        task_data['psd_object'].is_group.return_value = False
        task_data['psd_object'].is_visible.return_value = True
        task_data['psd_object'].topil.return_value = Image.new('RGB', (100, 100))
        task_data['psd_object'].__iter__ = Mock(return_value=iter([]))
        
        # Mock os.makedirs to raise PermissionError
        with patch('batch_export.os.makedirs') as mock_makedirs:
            mock_makedirs.side_effect = PermissionError("Permission denied")
            
            # Should handle permission errors specifically
            with pytest.raises((PermissionError, OSError)) as exc_info:
                result = export_single_image_task(task_data)
        # Verify it's a permission-related error
        error_msg = str(exc_info.value).lower()
        assert any(term in error_msg for term in ["permission", "denied", "read-only", "access"])
        del exc_info  # Clean up exception reference


class TestBoundaryConditionErrors:
    """边界条件错误处理测试"""
    
    def test_extremely_long_text(self):
        """Test extremely long text handling"""
        mock_layer = Mock()
        mock_layer.name = "@title#t"
        mock_layer.engine_dict = {
            'StyleRun': {
                'RunArray': [{
                    'StyleSheet': {
                        'StyleSheetData': {
                            'FontSize': 20
                        }
                    }
                }]
            }
        }
        mock_layer.size = (200, 50)
        mock_layer.offset = (0, 0)
        
        mock_image = Mock()
        
        # Test extremely long text
        long_text = "A" * 10000
        
        with patch('PIL.ImageFont.truetype') as mock_font:
            with patch('PIL.ImageDraw.Draw') as mock_draw:
                update_text_layer(mock_layer, long_text, mock_image)
                # Should handle long text
                mock_font.assert_called_once()
                mock_draw.assert_called_once()
    
    def test_zero_size_layer(self):
        """Test zero size layer handling"""
        mock_layer = Mock()
        mock_layer.name = "@title#t"
        mock_layer.engine_dict = {
            'StyleRun': {
                'RunArray': [{
                    'StyleSheet': {
                        'StyleSheetData': {
                            'FontSize': 20
                        }
                    }
                }]
            }
        }
        mock_layer.size = (0, 0)  # Zero size
        mock_layer.offset = (0, 0)
        
        mock_image = Mock()
        
        with patch('PIL.ImageFont.truetype') as mock_font:
            with patch('PIL.ImageDraw.Draw') as mock_draw:
                update_text_layer(mock_layer, "Test text", mock_image)
                # Should handle zero size layer
                mock_font.assert_called_once()
                mock_draw.assert_called_once()
    
    def test_negative_offset_layer(self):
        """Test negative offset layer handling"""
        mock_layer = Mock()
        mock_layer.name = "@title#t"
        mock_layer.engine_dict = {
            'StyleRun': {
                'RunArray': [{
                    'StyleSheet': {
                        'StyleSheetData': {
                            'FontSize': 20
                        }
                    }
                }]
            }
        }
        mock_layer.size = (200, 50)
        mock_layer.offset = (-10, -20)  # Negative offset
        
        mock_image = Mock()
        
        with patch('PIL.ImageFont.truetype') as mock_font:
            with patch('PIL.ImageDraw.Draw') as mock_draw:
                update_text_layer(mock_layer, "Test text", mock_image)
                # Should handle negative offset layer
                mock_font.assert_called_once()
                mock_draw.assert_called_once()


if __name__ == "__main__":
    pytest.main([__file__, "-v"])