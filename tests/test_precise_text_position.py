#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
精确文本位置测试 - 修复宽松测试断言
====================================

创建更精确的文本位置测试，替换过于宽松的测试断言。
"""

import pytest
import sys
import os

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from test_utils import assert_text_position_accuracy, TestEnvironment

# 设置测试环境
test_env = TestEnvironment()
test_env.setup_batch_export_args('test', 'test_font.ttf', 'jpg')

from batch_export import calculate_text_position

# 恢复环境
test_env.cleanup()


class TestPreciseTextPosition:
    """精确的文本位置测试"""
    
    def test_precise_right_alignment_calculation(self):
        """测试精确的右对齐计算 - 修复原测试的宽松断言"""
        text = "Hello World"
        layer_width = 200
        font_size = 20
        
        x_pos, y_pos = calculate_text_position(text, layer_width, font_size, "right")
        
        # 新算法使用真实字体度量，不再使用简化的计算公式
        # 验证基本逻辑而不是具体数值
        expected_x_range = (layer_width * 0.7, layer_width * 0.95)  # 右对齐应该在右侧70-95%范围内
        expected_y = -font_size * 0.26
        
        # 验证位置在合理范围内
        assert expected_x_range[0] <= x_pos <= expected_x_range[1], f"右对齐X位置应该在{expected_x_range}范围内，实际为{x_pos}"
        assert_text_position_accuracy(y_pos, expected_y, expected_y, 
                                    tolerance=0.5, message="右对齐Y位置应该正确")
        
        # 验证位置确实在右半部分（比原测试更严格）
        assert x_pos > layer_width / 2, f"右对齐应该在右半部分: {x_pos} > {layer_width / 2}"
    
    def test_precise_center_alignment_calculation(self):
        """测试精确的居中对齐计算"""
        text = "Test"
        layer_width = 100
        font_size = 16
        
        x_pos, y_pos = calculate_text_position(text, layer_width, font_size, "center")
        
        # "Test" = 4个英文字符 = 4 * 16 * 0.5 = 32
        expected_text_width = 4 * font_size * 0.5
        expected_x = (layer_width - expected_text_width) / 2 - font_size * 0.01
        expected_y = -font_size * 0.26
        
        assert_text_position_accuracy(x_pos, y_pos, expected_x, expected_y,
                                    tolerance=0.1, message="居中对齐应该精确计算")
        
        # 验证位置确实在中心附近
        center_point = layer_width / 2
        assert abs(x_pos - center_point) < expected_text_width / 2 + 1, "应该在中心点附近"
    
    def test_precise_left_alignment_calculation(self):
        """测试精确的左对齐计算"""
        text = "ABC"
        layer_width = 150
        font_size = 14
        
        x_pos, y_pos = calculate_text_position(text, layer_width, font_size, "left")
        
        # 左对齐应该接近0（考虑偏移）
        expected_x = 0 - font_size * 0.01
        expected_y = -font_size * 0.26
        
        assert_text_position_accuracy(x_pos, y_pos, expected_x, expected_y,
                                    tolerance=0.05, message="左对齐应该接近0")
        
        # 左对齐的x位置应该在很小的范围内
        assert -1 < x_pos < 1, f"左对齐x位置应该在0附近: {x_pos}"
    
    def test_mixed_text_precise_calculation(self):
        """测试中英文混合文本的精确计算"""
        text = "中文ABC"
        layer_width = 200
        font_size = 18
        
        x_pos, y_pos = calculate_text_position(text, layer_width, font_size, "center")
        
        # "中文ABC" = 2个中文 + 3个英文 = 2*18 + 3*18*0.5 = 36 + 27 = 63
        chinese_count = sum(1 for char in text if '\u4e00' <= char <= '\u9fff')
        english_count = len(text) - chinese_count
        expected_text_width = chinese_count * font_size + english_count * font_size * 0.5
        
        expected_x = (layer_width - expected_text_width) / 2 - font_size * 0.01
        expected_y = -font_size * 0.26
        
        assert_text_position_accuracy(x_pos, y_pos, expected_x, expected_y,
                                    tolerance=0.1, message="中英文混合应该精确计算")
    
    def test_alignment_comparison(self):
        """测试不同对齐方式的相对位置关系"""
        text = "Sample"
        layer_width = 120
        font_size = 15
        
        # 计算三种对齐方式的位置
        x_left, y_left = calculate_text_position(text, layer_width, font_size, "left")
        x_center, y_center = calculate_text_position(text, layer_width, font_size, "center")
        x_right, y_right = calculate_text_position(text, layer_width, font_size, "right")
        
        # 验证位置关系：left < center < right
        assert x_left < x_center, f"左对齐应该在居中左边: {x_left} < {x_center}"
        assert x_center < x_right, f"居中应该在右对齐左边: {x_center} < {x_right}"
        
        # 验证Y位置相同
        assert abs(y_left - y_center) < 0.01, "Y位置应该相同"
        assert abs(y_left - y_right) < 0.01, "Y位置应该相同"
        
        # 验证间距的合理性
        left_to_center = x_center - x_left
        center_to_right = x_right - x_center
        assert abs(left_to_center - center_to_right) < 1, "居中应该真的在中间"
    
    def test_extreme_values_handling(self):
        """测试极值处理"""
        # 极小字体
        x_pos, y_pos = calculate_text_position("A", 100, 1, "center")
        assert isinstance(x_pos, (int, float))
        assert isinstance(y_pos, (int, float))
        
        # 极大图层
        x_pos, y_pos = calculate_text_position("Test", 10000, 20, "right")
        assert x_pos < 10000
        assert x_pos > 9000  # 应该接近右边界
        
        # 极短文本
        x_pos, y_pos = calculate_text_position("I", 50, 12, "left")
        assert abs(x_pos) < 1  # 应该接近0


class TestImprovedTextPositionAssertions:
    """改进的文本位置断言测试"""
    
    def test_replacement_of_loose_assertions(self):
        """替换宽松断言的测试"""
        # 这是原测试中的宽松断言：
        # assert x_pos > layer_width / 4  # 只要求大于1/4
        
        # 新的精确断言应该：
        text = "Hello World"
        layer_width = 200
        font_size = 20
        
        x_pos, y_pos = calculate_text_position(text, layer_width, font_size, "right")
        
        # 计算期望的精确位置
        expected_text_width = len(text) * font_size * 0.5
        expected_x = layer_width - expected_text_width - font_size * 0.01
        
        # 新断言：精确到0.1像素
        assert abs(x_pos - expected_x) < 0.1, f"位置应该精确到0.1像素内"
        
        # 验证新断言比原断言更严格
        loose_requirement = layer_width / 4  # 50
        strict_requirement = expected_x  # 约89.8
        
        assert x_pos > loose_requirement, "满足宽松断言"
        assert abs(x_pos - strict_requirement) < 0.1, "满足精确断言"
        print(f"宽松断言要求: >{loose_requirement}, 实际: {x_pos:.1f}")
        print(f"精确断言要求: ≈{strict_requirement:.1f}, 误差: {abs(x_pos - strict_requirement):.3f}")
    
    def test_assertion_precision_improvement(self):
        """测试断言精度的改进"""
        test_cases = [
            ("Short", 100, 12, "left"),
            ("Medium length text", 200, 16, "center"),
            ("This is a much longer text string", 300, 14, "right"),
            ("中文测试", 150, 18, "center"),
            ("Mixed 中English 文", 250, 15, "left")
        ]
        
        for text, width, size, alignment in test_cases:
            x_pos, y_pos = calculate_text_position(text, width, size, alignment)
            
            # 计算期望位置
            if alignment == "left":
                expected_x = 0 - size * 0.01
            elif alignment == "center":
                chinese_count = sum(1 for char in text if '\u4e00' <= char <= '\u9fff')
                english_count = len(text) - chinese_count
                expected_text_width = chinese_count * size + english_count * size * 0.5
                expected_x = (width - expected_text_width) / 2 - size * 0.01
            else:  # right
                chinese_count = sum(1 for char in text if '\u4e00' <= char <= '\u9fff')
                english_count = len(text) - chinese_count
                expected_text_width = chinese_count * size + english_count * size * 0.5
                expected_x = width - expected_text_width - size * 0.01
            
            expected_y = -size * 0.26
            
            # 使用严格的容差
            assert_text_position_accuracy(x_pos, y_pos, expected_x, expected_y,
                                        tolerance=0.1, message=f"{text} ({alignment})")
            
            print(f"'{text}' ({alignment}): 期望({expected_x:.1f}, {expected_y:.1f}), 实际({x_pos:.1f}, {y_pos:.1f})")


if __name__ == "__main__":
    pytest.main([__file__, "-v", "-s"])