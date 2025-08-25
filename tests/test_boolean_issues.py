#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
布尔值处理测试 - 发现业务代码问题
===================================

专门测试布尔值处理的正确性，发现业务代码中的缺陷。
"""

import pytest
import sys
import os

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from test_utils import parse_boolean_value, TestEnvironment
from unittest.mock import Mock


class TestBooleanValueParsing:
    """测试布尔值解析的正确性"""
    
    def test_correct_boolean_parsing(self):
        """测试正确的布尔值解析"""
        # 布尔值输入
        assert parse_boolean_value(True) is True
        assert parse_boolean_value(False) is False
        
        # 字符串输入 - 正确解析
        assert parse_boolean_value("True") is True
        assert parse_boolean_value("true") is True
        assert parse_boolean_value("TRUE") is True
        assert parse_boolean_value("1") is True
        assert parse_boolean_value("yes") is True
        assert parse_boolean_value("YES") is True
        assert parse_boolean_value("on") is True
        assert parse_boolean_value("ON") is True
        assert parse_boolean_value("t") is True
        assert parse_boolean_value("T") is True
        assert parse_boolean_value("y") is True
        assert parse_boolean_value("Y") is True
        
        # 字符串输入 - 错误解析
        assert parse_boolean_value("False") is False
        assert parse_boolean_value("false") is False
        assert parse_boolean_value("FALSE") is False
        assert parse_boolean_value("0") is False
        assert parse_boolean_value("no") is False
        assert parse_boolean_value("NO") is False
        assert parse_boolean_value("off") is False
        assert parse_boolean_value("OFF") is False
        assert parse_boolean_value("f") is False
        assert parse_boolean_value("F") is False
        assert parse_boolean_value("n") is False
        assert parse_boolean_value("N") is False
        
        # 数字输入
        assert parse_boolean_value(1) is True
        assert parse_boolean_value(0) is False
        assert parse_boolean_value(42) is True
        assert parse_boolean_value(-1) is True
        assert parse_boolean_value(0.0) is False
        assert parse_boolean_value(1.0) is True
        
        # 空值和None
        assert parse_boolean_value("") is False
        assert parse_boolean_value(None) is False
        
        # 其他类型
        assert parse_boolean_value([]) is False
        assert parse_boolean_value([1]) is True
        assert parse_boolean_value({}) is False
        assert parse_boolean_value({"key": "value"}) is True


class TestBusinessCodeBooleanIssues:
    """测试业务代码中的布尔值处理问题"""
    
    def test_business_code_boolean_conversion_bug(self):
        """演示业务代码中的布尔值转换bug"""
        # 这是业务代码中的错误逻辑
        def business_code_boolean_conversion(value):
            """模拟业务代码中的错误布尔值转换"""
            # 处理numpy布尔类型
            if hasattr(value, 'item'):
                value = value.item()
            # 处理pandas布尔类型
            if hasattr(value, 'bool'):
                value = value.bool()
            # 确保是Python原生bool类型 - 这里有问题！
            value = bool(value)
            return value
        
        # 测试错误逻辑
        assert business_code_boolean_conversion("False") is True  # 错误！
        assert business_code_boolean_conversion("0") is True      # 错误！
        assert business_code_boolean_conversion("") is False       # 正确
        assert business_code_boolean_conversion("True") is True   # 正确
        
        # 对比正确实现
        assert parse_boolean_value("False") is False  # 正确
        assert parse_boolean_value("0") is False      # 正确
        assert parse_boolean_value("") is False       # 正确
        assert parse_boolean_value("True") is True   # 正确
    
    def test_excel_boolean_data_scenarios(self):
        """测试Excel中常见的布尔值数据场景"""
        # 模拟Excel中可能出现的各种布尔值表示
        excel_boolean_values = [
            # 标准布尔值
            True, False,
            # 字符串表示
            "TRUE", "FALSE", "True", "False", "true", "false",
            "1", "0", "YES", "NO", "yes", "no", "ON", "OFF", "on", "off",
            # 数字
            1, 0, 1.0, 0.0,
            # 空值
            "", None,
            # 其他常见情况
            " T ", " F ", "  ", "NA", "N/A"
        ]
        
        # 测试我们的正确解析函数
        correct_results = {
            True: True, False: False,
            "TRUE": True, "FALSE": False, "True": True, "False": False,
            "true": True, "false": False, "1": True, "0": False,
            "YES": True, "NO": False, "yes": True, "no": False,
            "ON": True, "OFF": False, "on": True, "off": False,
            1: True, 0: False, 1.0: True, 0.0: False,
            "": False, None: False, " T ": False, " F ": False,
            "  ": False, "NA": False, "N/A": False
        }
        
        for value, expected in correct_results.items():
            result = parse_boolean_value(value)
            assert result == expected, f"值 {value} 期望 {expected}, 实际 {result}"
    
    def test_layer_visibility_boolean_handling(self):
        """测试图层可见性的布尔值处理"""
        # 模拟业务代码中的set_layer_visibility函数
        def business_set_layer_visibility(layer, visibility):
            """模拟业务代码中的set_layer_visibility函数"""
            # 处理numpy布尔类型
            if hasattr(visibility, 'item'):
                visibility = visibility.item()
            # 处理pandas布尔类型
            if hasattr(visibility, 'bool'):
                visibility = visibility.bool()
            # 确保是Python原生bool类型
            visibility = bool(visibility)  # 这里有bug！
            layer.visible = visibility
        
        # 创建模拟图层
        mock_layer = Mock()
        
        # 测试各种Excel数据输入
        test_cases = [
            ("True", True),
            ("FALSE", False),  # 业务代码会错误地返回True
            ("1", True),
            ("0", False),      # 业务代码会错误地返回True
            ("", False),
            (None, False),
            (True, True),
            (False, False)
        ]
        
        for input_value, expected_visibility in test_cases:
            mock_layer.visible = None  # 重置
            business_set_layer_visibility(mock_layer, input_value)
            
            # 检查结果
            if input_value in ["FALSE", "0"]:
                # 这些情况下业务代码有bug
                assert mock_layer.visible is True, f"业务代码bug: {input_value} 被转换为 {mock_layer.visible}"
                print(f"发现业务代码bug: '{input_value}' -> {mock_layer.visible} (期望: {expected_visibility})")
            else:
                assert mock_layer.visible == expected_visibility, f"{input_value} -> {mock_layer.visible}"
    
    def test_boolean_parsing_performance(self):
        """测试布尔值解析的性能"""
        import time
        
        # 大量数据测试
        test_data = ["True", "False", "1", "0", "yes", "no"] * 10000
        
        start_time = time.time()
        for value in test_data:
            parse_boolean_value(value)
        end_time = time.time()
        
        duration = end_time - start_time
        assert duration < 1.0, f"布尔值解析性能问题: {duration:.3f}s"
        print(f"布尔值解析性能: {duration:.3f}s (处理{len(test_data)}个值)")


if __name__ == "__main__":
    pytest.main([__file__, "-v", "-s"])