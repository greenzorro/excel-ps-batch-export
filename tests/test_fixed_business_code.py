#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试修复后的业务代码
==================

验证业务代码问题修复后的正确性。
"""

import pytest
import sys
import os

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from test_utils import TestEnvironment

# 设置测试环境
test_env = TestEnvironment()
test_env.setup_batch_export_args('test', 'test_font.ttf', 'jpg')

from batch_export import set_layer_visibility
from unittest.mock import Mock

# 恢复环境
test_env.cleanup()


class TestFixedBusinessCode:
    """测试修复后的业务代码"""
    
    def test_fixed_boolean_conversion(self):
        """测试修复后的布尔值转换"""
        # 创建模拟图层
        mock_layer = Mock()
        
        # 测试各种布尔值输入
        test_cases = [
            # (输入值, 期望结果)
            (True, True),
            (False, False),
            ("True", True),
            ("true", True),
            ("TRUE", True),
            ("1", True),
            ("yes", True),
            ("YES", True),
            ("on", True),
            ("ON", True),
            ("t", True),
            ("T", True),
            ("y", True),
            ("Y", True),
            ("False", False),  # 这是修复的关键！
            ("false", False),
            ("FALSE", False),
            ("0", False),      # 这是修复的关键！
            ("no", False),
            ("NO", False),
            ("off", False),
            ("OFF", False),
            ("f", False),
            ("F", False),
            ("n", False),
            ("N", False),
            (1, True),
            (0, False),
            (42, True),
            (-1, True),
            (0.0, False),
            (1.0, True),
            ("", False),      # 空字符串
            (None, False),    # None值
        ]
        
        for input_value, expected_result in test_cases:
            mock_layer.visible = None  # 重置
            set_layer_visibility(mock_layer, input_value)
            
            print(f"输入: {input_value!r} -> 输出: {mock_layer.visible} (期望: {expected_result})")
            assert mock_layer.visible == expected_result, f"修复失败: {input_value!r} -> {mock_layer.visible}, 期望 {expected_result}"
        
        print("✅ 布尔值转换修复验证成功！")
    
    def test_fixed_vs_original_comparison(self):
        """对比修复前后代码的行为差异"""
        mock_layer = Mock()
        
        # 原始错误逻辑（用于对比）
        def original_set_layer_visibility(layer, visibility):
            """原始的错误实现"""
            if hasattr(visibility, 'item'):
                visibility = visibility.item()
            if hasattr(visibility, 'bool'):
                visibility = visibility.bool()
            visibility = bool(visibility)  # 这里有bug！
            layer.visible = visibility
        
        # 测试问题案例
        problematic_cases = [
            ("FALSE", False),
            ("false", False),
            ("0", False),
            ("no", False),
            ("off", False),
        ]
        
        for input_value, expected in problematic_cases:
            # 测试原始实现
            mock_layer.visible = None
            original_set_layer_visibility(mock_layer, input_value)
            original_result = mock_layer.visible
            
            # 测试修复后实现
            mock_layer.visible = None
            set_layer_visibility(mock_layer, input_value)
            fixed_result = mock_layer.visible
            
            print(f"输入: {input_value!r}")
            print(f"  原始实现: {original_result} {'(错误)' if original_result != expected else '(正确)'}")
            print(f"  修复实现: {fixed_result} {'(正确)' if fixed_result == expected else '(错误)'}")
            
            assert original_result != expected, f"原始实现应该有bug: {input_value} -> {original_result}"
            assert fixed_result == expected, f"修复实现应该正确: {input_value} -> {fixed_result}"
        
        print("✅ 修复效果验证成功！")
    
    def test_edge_cases_handling(self):
        """测试边界情况处理"""
        mock_layer = Mock()
        
        # 测试边界情况
        edge_cases = [
            ("  True  ", True),  # 带空格
            ("  FALSE  ", False),
            ("TRUE", True),      # 大小写混合
            ("FaLsE", False),
            ("1.0", True),      # 浮点数字符串
            ("0.0", False),
            ("  ", False),      # 只有空格
            ("unknown", True),  # 未知字符串（非空即True）
        ]
        
        for input_value, expected in edge_cases:
            mock_layer.visible = None
            set_layer_visibility(mock_layer, input_value)
            
            result = mock_layer.visible
            print(f"边界情况: {input_value!r} -> {result} (期望: {expected})")
            assert result == expected, f"边界情况处理失败: {input_value!r}"
        
        print("✅ 边界情况处理验证成功！")
    
    def test_performance_impact(self):
        """测试修复后的性能影响"""
        import time
        
        mock_layer = Mock()
        
        # 生成大量测试数据
        test_data = ["True", "False", "1", "0", "yes", "no"] * 10000
        
        start_time = time.time()
        for value in test_data:
            set_layer_visibility(mock_layer, value)
        end_time = time.time()
        
        duration = end_time - start_time
        print(f"修复后函数处理 {len(test_data)} 个值耗时: {duration:.3f}s")
        
        # 性能应该在合理范围内
        assert duration < 1.0, f"性能问题: 处理{len(test_data)}个值耗时{duration:.3f}s"
        print("✅ 性能影响验证成功！")


if __name__ == "__main__":
    pytest.main([__file__, "-v", "-s"])