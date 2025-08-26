#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文本位置计算测试 - 发现业务代码问题
===================================

专门测试文本位置计算的准确性，发现业务代码中的算法缺陷。
"""

import pytest
import sys
import os

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from test_utils import assert_text_position_accuracy
from unittest.mock import Mock, patch
from PIL import Image, ImageDraw, ImageFont


# 导入业务代码函数
original_argv = sys.argv.copy()
sys.argv = ['batch_export.py', '1', 'test_font.ttf', 'jpg']

try:
    from batch_export import calculate_text_position
finally:
    sys.argv = original_argv


class TestTextPositionCalculation:
    """测试文本位置计算的准确性"""
    
    def test_chinese_text_position_calculation(self):
        """测试中文文本位置计算"""
        text = "中文测试"
        layer_width = 200
        font_size = 20
        
        # 测试居中对齐
        x_pos, y_pos = calculate_text_position(text, layer_width, font_size, "center")
        
        # 中文文本：每个字符宽度 = font_size
        expected_text_width = len(text) * font_size  # 4 * 20 = 80
        expected_x = (layer_width - expected_text_width) / 2 - font_size * 0.01  # (200-80)/2 - 0.2 = 59.8
        expected_y = -font_size * 0.26  # -5.2
        
        assert_text_position_accuracy(x_pos, y_pos, expected_x, expected_y, 
                                    tolerance=0.1, message="中文居中对齐")
    
    def test_english_text_position_calculation(self):
        """测试英文文本位置计算"""
        text = "Hello"
        layer_width = 200
        font_size = 20
        
        # 测试右对齐
        x_pos, y_pos = calculate_text_position(text, layer_width, font_size, "right")
        
        # 英文文本：新算法使用真实字体度量，结果可能不同
        # 原算法期望：expected_text_width = len(text) * font_size * 0.5 = 50, expected_x = 149.8
        # 新算法可能给出不同但更精确的结果，我们验证基本逻辑即可
        assert x_pos > layer_width / 2  # 右对齐应该在右半部分
        assert y_pos < 0  # Y位置应该是负数
        assert x_pos < layer_width  # X位置应该在图层范围内
    
    def test_mixed_text_position_calculation(self):
        """测试中英文混合文本位置计算"""
        text = "中文Hello混合"
        layer_width = 300
        font_size = 16
        
        # 测试左对齐
        x_pos, y_pos = calculate_text_position(text, layer_width, font_size, "left")
        
        # 新算法使用真实字体度量，验证基本逻辑而不是具体数值
        # 左对齐位置可能在负值附近（考虑字体偏移）
        assert y_pos < 0  # Y位置应该是负数
        assert x_pos < layer_width / 4  # 左对齐应该在左侧部分
    
    def test_empty_text_position_calculation(self):
        """测试空文本位置计算"""
        text = ""
        layer_width = 100
        font_size = 16
        
        x_pos, y_pos = calculate_text_position(text, layer_width, font_size, "center")
        
        # 空文本的宽度应该是0，应该在中心位置
        expected_x = layer_width / 2  # 50
        expected_y = -font_size * 0.26  # -4.16
        
        assert_text_position_accuracy(x_pos, y_pos, expected_x, expected_y,
                                    tolerance=1.0, message="空文本居中对齐")
    
    def test_special_characters_position_calculation(self):
        """测试特殊字符位置计算"""
        text = "@#$%^&*()"
        layer_width = 150
        font_size = 14
        
        # 特殊字符按英文字符处理
        x_pos, y_pos = calculate_text_position(text, layer_width, font_size, "center")
        
        # 新算法使用真实字体度量，验证基本逻辑而不是具体数值
        expected_y = -font_size * 0.26  # -3.64
        
        # 验证Y位置正确
        assert abs(y_pos - expected_y) <= 0.5, f"特殊字符Y位置应该正确: 期望{expected_y}, 实际{y_pos}"
        
        # 验证X位置在合理范围内（居中应该在中心附近，根据真实字体度量调整）
        center_range = (layer_width * 0.1, layer_width * 0.9)
        assert center_range[0] <= x_pos <= center_range[1], f"特殊字符居中X位置应该在{center_range}范围内: {x_pos}"
    
    def test_single_character_position_calculation(self):
        """测试单字符位置计算"""
        text = "A"
        layer_width = 50
        font_size = 12
        
        # 测试不同对齐方式
        # 左对齐
        x_left, y_left = calculate_text_position(text, layer_width, font_size, "left")
        assert -5 <= x_left <= 5, f"单字符左对齐X位置应该在0附近: {x_left}"
        
        # 居中对齐
        x_center, y_center = calculate_text_position(text, layer_width, font_size, "center")
        center_range = (layer_width * 0.2, layer_width * 0.8)
        assert center_range[0] <= x_center <= center_range[1], f"单字符居中X位置应该在{center_range}范围内: {x_center}"
        
        # 右对齐
        x_right, y_right = calculate_text_position(text, layer_width, font_size, "right")
        assert x_right > layer_width * 0.6, f"单字符右对齐X位置应该在右侧60%以上: {x_right} > {layer_width * 0.6}"
        assert x_right < layer_width, f"单字符右对齐不应该超出边界: {x_right} < {layer_width}"
        
        # Y位置应该相同
        assert abs(y_left - y_right) < 0.5
        assert abs(y_left - y_center) < 0.5
        
        # 验证位置关系：left < center < right
        assert x_left < x_center, f"左对齐应该在居中左边: {x_left} < {x_center}"
        assert x_center < x_right, f"居中应该在右对齐左边: {x_center} < {x_right}"


class TestTextPositionAlgorithmIssues:
    """测试文本位置算法的问题"""
    
    def test_algorithm_simplification_issues(self):
        """测试算法简化带来的问题"""
        """业务代码使用简化的字符宽度计算：
        - 中文字符：width = font_size
        - 英文字符：width = font_size * 0.5
        这种简化在真实字体中可能不准确"""
        
        # 模拟真实字体度量（使用Pillow）
        def get_real_text_width(text, font_size):
            """使用真实字体度量获取文本宽度"""
            try:
                font = ImageFont.truetype("arial.ttf", font_size)
                draw = ImageDraw.Draw(Image.new('RGB', (1, 1)))
                bbox = draw.textbbox((0, 0), text, font=font)
                return bbox[2] - bbox[0]  # 右边界 - 左边界
            except:
                # 如果字体加载失败，回退到业务代码算法
                width = 0
                for char in text:
                    if '\u4e00' <= char <= '\u9fff':
                        width += font_size
                    else:
                        width += font_size * 0.5
                return width
        
        # 测试用例
        test_cases = [
            ("Hello", 16),
            ("中文", 16),
            ("HelloWorld", 12),
            ("这是一段很长的中文文本", 14),
            ("Mixed中English文", 18)
        ]
        
        for text, font_size in test_cases:
            # 业务代码算法
            business_width = 0
            for char in text:
                if '\u4e00' <= char <= '\u9fff':
                    business_width += font_size
                else:
                    business_width += font_size * 0.5
            
            # 真实字体度量（可能）
            try:
                real_width = get_real_text_width(text, font_size)
                
                # 计算差异百分比
                if real_width > 0:
                    diff_percent = abs(business_width - real_width) / real_width * 100
                    print(f"文本 '{text}': 业务算法={business_width:.1f}, 真实度量={real_width:.1f}, 差异={diff_percent:.1f}%")
                    
                    # 如果差异超过10%，说明算法有显著问题
                    if diff_percent > 10:
                        print(f"⚠️  算法差异过大: {text}")
                
            except Exception as e:
                print(f"无法获取真实字体度量: {e}")
    
    def test_position_calculation_edge_cases(self):
        """测试位置计算的边界情况"""
        # 测试边界值验证
        with pytest.raises(ValueError, match="字体大小必须大于0"):
            calculate_text_position("test", 100, 0, "center")
        
        with pytest.raises(ValueError, match="图层宽度不能为负数"):
            calculate_text_position("test", -100, 16, "center")
        
        with pytest.raises(ValueError, match="对齐方式必须是"):
            calculate_text_position("test", 100, 16, "invalid")
    
    def test_long_text_position_calculation(self):
        """测试长文本位置计算"""
        text = "这是一个很长的中文文本，用于测试长文本的位置计算准确性"
        layer_width = 400
        font_size = 14
        
        # 测试右对齐
        x_pos, y_pos = calculate_text_position(text, layer_width, font_size, "right")
        
        # 计算期望位置
        chinese_count = sum(1 for char in text if '\u4e00' <= char <= '\u9fff')
        expected_text_width = chinese_count * font_size
        expected_x = layer_width - expected_text_width - font_size * 0.01
        
        # 长文本可能超出图层宽度，x_pos可能为负数
        assert x_pos <= expected_x + 0.1
        assert y_pos == -font_size * 0.26
    
    def test_text_position_calculation_consistency(self):
        """测试文本位置计算的一致性"""
        text = "Test"
        layer_width = 100
        font_size = 16
        
        # 多次调用应该返回相同结果
        positions = []
        for _ in range(10):
            x_pos, y_pos = calculate_text_position(text, layer_width, font_size, "center")
            positions.append((x_pos, y_pos))
        
        # 检查所有结果是否相同
        first_x, first_y = positions[0]
        for i, (x, y) in enumerate(positions[1:], 1):
            assert abs(x - first_x) < 0.001, f"第{i}次调用x位置不一致: {x} vs {first_x}"
            assert abs(y - first_y) < 0.001, f"第{i}次调用y位置不一致: {y} vs {first_y}"


if __name__ == "__main__":
    pytest.main([__file__, "-v", "-s"])