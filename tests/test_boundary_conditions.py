#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
真正的边界条件测试 - 发现业务代码问题
=====================================

这个测试用例真正验证边界条件，而不是为了通过而测试
"""

import os
import sys
import tempfile
import pytest
import pandas as pd
from unittest.mock import Mock, patch
from PIL import Image, ImageDraw, ImageFont

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 导入共享测试工具
from test_utils import TestEnvironment

# 使用测试环境管理器处理sys.argv依赖
test_env = TestEnvironment()
test_env.setup_psd_renderer_args('test', 'jpg')

from psd_renderer import calculate_text_position, read_excel_file

# 恢复原始环境
test_env.cleanup()


class TestRealBoundaryConditions:
    """真正的边界条件测试"""

    def test_calculate_text_position_zero_font_size_should_raise_error(self):
        """测试：零字体大小应该抛出异常"""
        text = "Test"
        layer_width = 100
        font_size = 0
        # 创建测试用的 draw 和 font
        test_img = Image.new("RGB", (layer_width, 50))
        draw = ImageDraw.Draw(test_img)
        font = ImageFont.load_default()

        # 现在业务代码应该抛出异常
        with pytest.raises(ValueError, match="字体大小必须大于0"):
            calculate_text_position(text, layer_width, font_size, "center", draw, font)

    def test_calculate_text_position_negative_width_should_raise_error(self):
        """测试：负图层宽度应该抛出异常"""
        text = "Test"
        layer_width = -100
        font_size = 16
        # 创建测试用的 draw 和 font
        test_img = Image.new("RGB", (100, 50))
        draw = ImageDraw.Draw(test_img)
        font = ImageFont.load_default()

        # 现在业务代码应该抛出异常
        with pytest.raises(ValueError, match="图层宽度不能为负数"):
            calculate_text_position(text, layer_width, font_size, "center", draw, font)

    def test_calculate_text_position_invalid_alignment_should_raise_error(self):
        """测试：无效对齐方式应该抛出异常"""
        text = "Test"
        layer_width = 100
        font_size = 16
        # 创建测试用的 draw 和 font
        test_img = Image.new("RGB", (layer_width, 50))
        draw = ImageDraw.Draw(test_img)
        font = ImageFont.load_default()

        # 测试无效对齐方式
        with pytest.raises(ValueError, match="对齐方式必须是"):
            calculate_text_position(text, layer_width, font_size, "invalid", draw, font)

    def test_calculate_text_position_empty_text_should_work_correctly(self):
        """测试：空文本应该正确计算位置"""
        text = ""
        layer_width = 100
        font_size = 16
        # 创建测试用的 draw 和 font
        test_img = Image.new("RGB", (layer_width, 50))
        draw = ImageDraw.Draw(test_img)
        font = ImageFont.load_default()

        x_pos, y_pos = calculate_text_position(text, layer_width, font_size, "center", draw, font)

        # 空文本应该返回有效位置
        assert isinstance(x_pos, (int, float)), "X位置应该是数字"
        assert isinstance(y_pos, (int, float)), "Y位置应该是数字"
        # 空文本居中应该在图层中心附近
        assert 40 < x_pos < 60, f"空文本应该居中，实际x={x_pos}"
    
    def test_read_excel_file_nonexistent_should_raise_error(self):
        """测试：读取不存在的Excel文件应该抛出异常"""
        nonexistent_file = "/path/that/does/not/exist.xlsx"
        
        # 现在业务代码应该抛出FileNotFoundError
        with pytest.raises(FileNotFoundError, match="Excel文件不存在"):
            read_excel_file(nonexistent_file)
    
    def test_read_excel_file_invalid_format_should_raise_error(self):
        """测试：读取不支持的文件格式应该抛出异常"""
        # 创建一个存在但格式不正确的文件
        tmp_path = None
        try:
            with tempfile.NamedTemporaryFile(suffix='.txt', delete=False) as tmp:
                tmp.write(b"This is not an Excel file")
                tmp.flush()
                tmp_path = tmp.name
            
            # 现在业务代码应该抛出ValueError
            with pytest.raises(ValueError, match="不支持的文件格式"):
                read_excel_file(tmp_path)
        finally:
            if tmp_path and os.path.exists(tmp_path):
                try:
                    os.unlink(tmp_path)
                except PermissionError:
                    pass  # 忽略权限错误，Windows下可能发生


