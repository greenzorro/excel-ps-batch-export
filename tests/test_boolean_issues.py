#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
布尔值处理测试 - 验证业务代码正确性
=====================================

测试 psd_renderer.py 中 set_layer_visibility 函数的布尔值处理逻辑，
确保各种来自 Excel 的字符串/数值/布尔值都能被正确解析。
"""

import pytest
import sys
import os
import time
from unittest.mock import Mock

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from psd_renderer import set_layer_visibility


def _make_mock_layer():
    """创建一个用于 set_layer_visibility 测试的 mock 图层对象"""
    layer = Mock()
    layer.visible = None
    return layer


class TestSetLayerVisibilityBooleanInputs:
    """测试 set_layer_visibility 对 Python 原生布尔值的处理"""

    def test_bool_true(self):
        layer = _make_mock_layer()
        set_layer_visibility(layer, True)
        assert layer.visible is True

    def test_bool_false(self):
        layer = _make_mock_layer()
        set_layer_visibility(layer, False)
        assert layer.visible is False


class TestSetLayerVisibilityTruthyStrings:
    """测试 set_layer_visibility 对各种真值字符串的处理"""

    @pytest.mark.parametrize("value", [
        "True", "true", "TRUE",
        "1",
        "yes", "YES", "Yes",
        "on", "ON", "On",
        "t", "T",
        "y", "Y",
    ])
    def test_truthy_strings(self, value):
        layer = _make_mock_layer()
        set_layer_visibility(layer, value)
        assert layer.visible is True, f"Expected True for input '{value}', got {layer.visible}"


class TestSetLayerVisibilityFalsyStrings:
    """测试 set_layer_visibility 对各种假值字符串的处理"""

    @pytest.mark.parametrize("value", [
        "False", "false", "FALSE",
        "0",
        "no", "NO", "No",
        "off", "OFF", "Off",
        "f", "F",
        "n", "N",
        "",       # 空字符串
        "  ",     # 仅空格
        "  \t ",  # 空白字符
    ])
    def test_falsy_strings(self, value):
        layer = _make_mock_layer()
        set_layer_visibility(layer, value)
        assert layer.visible is False, f"Expected False for input '{value}', got {layer.visible}"


class TestSetLayerVisibilityNumericInputs:
    """测试 set_layer_visibility 对数值类型输入的处理"""

    def test_integer_nonzero(self):
        layer = _make_mock_layer()
        set_layer_visibility(layer, 1)
        assert layer.visible is True

    def test_integer_zero(self):
        layer = _make_mock_layer()
        set_layer_visibility(layer, 0)
        assert layer.visible is False

    def test_float_nonzero(self):
        layer = _make_mock_layer()
        set_layer_visibility(layer, 1.0)
        assert layer.visible is True

    def test_float_zero(self):
        layer = _make_mock_layer()
        set_layer_visibility(layer, 0.0)
        assert layer.visible is False

    def test_negative_number(self):
        layer = _make_mock_layer()
        set_layer_visibility(layer, -1)
        assert layer.visible is True

    def test_large_number(self):
        layer = _make_mock_layer()
        set_layer_visibility(layer, 42)
        assert layer.visible is True


class TestSetLayerVisibilityNoneAndEdgeCases:
    """测试 set_layer_visibility 对 None 和其他边界值的处理"""

    def test_none_input(self):
        layer = _make_mock_layer()
        set_layer_visibility(layer, None)
        assert layer.visible is False

    def test_whitespace_stripped_truthy(self):
        """带空格的真值字符串应正确解析"""
        layer = _make_mock_layer()
        set_layer_visibility(layer, " True ")
        assert layer.visible is True

    def test_whitespace_stripped_falsy(self):
        """带空格的假值字符串应正确解析"""
        layer = _make_mock_layer()
        set_layer_visibility(layer, " false ")
        assert layer.visible is False

    def test_numeric_string_non_zero(self):
        """非零数字字符串应解析为 True（通过 float 回退）"""
        layer = _make_mock_layer()
        set_layer_visibility(layer, "2.5")
        assert layer.visible is True

    def test_numeric_string_zero_float(self):
        """字符串 '0.0' 应解析为 False"""
        layer = _make_mock_layer()
        set_layer_visibility(layer, "0.0")
        assert layer.visible is False

    def test_unrecognized_non_empty_string(self):
        """无法识别的非空字符串会通过 bool() 回退为 True"""
        layer = _make_mock_layer()
        set_layer_visibility(layer, "random_text")
        assert layer.visible is True


class TestSetLayerVisibilityNumpyPandasCompat:
    """测试 set_layer_visibility 对 numpy/pandas 类型的兼容处理"""

    def test_numpy_like_item(self):
        """模拟具有 .item() 方法的 numpy 布尔值"""
        layer = _make_mock_layer()
        numpy_like = Mock()
        numpy_like.item.return_value = True
        # numpy bool 不应有 .bool() 方法（排除 hasattr 判断）
        del numpy_like.bool
        set_layer_visibility(layer, numpy_like)
        assert layer.visible is True

    def test_numpy_like_item_false(self):
        layer = _make_mock_layer()
        numpy_like = Mock()
        numpy_like.item.return_value = False
        del numpy_like.bool
        set_layer_visibility(layer, numpy_like)
        assert layer.visible is False

    def test_pandas_like_bool(self):
        """模拟具有 .bool() 方法的 pandas 布尔值"""
        layer = _make_mock_layer()
        pandas_like = Mock()
        pandas_like.bool.return_value = True
        # 不触发 .item() 分支
        del pandas_like.item
        set_layer_visibility(layer, pandas_like)
        assert layer.visible is True

    def test_pandas_like_bool_false(self):
        layer = _make_mock_layer()
        pandas_like = Mock()
        pandas_like.bool.return_value = False
        del pandas_like.item
        set_layer_visibility(layer, pandas_like)
        assert layer.visible is False

    def test_numpy_string_value_via_item(self):
        """numpy 字符串值通过 .item() 还原后正确解析"""
        layer = _make_mock_layer()
        numpy_str = Mock()
        numpy_str.item.return_value = "FALSE"
        del numpy_str.bool
        set_layer_visibility(layer, numpy_str)
        assert layer.visible is False


class TestSetLayerVisibilityExcelScenarios:
    """模拟 Excel 数据中常见的布尔值场景"""

    EXCEL_TEST_CASES = [
        # (输入值, 期望的 visible 结果, 描述)
        (True, True, "Python bool True"),
        (False, False, "Python bool False"),
        ("TRUE", True, "Excel TRUE 函数输出（大写）"),
        ("FALSE", False, "Excel FALSE 函数输出（大写）"),
        ("True", True, "Excel 字符串 True"),
        ("False", False, "Excel 字符串 False"),
        ("1", True, "Excel 数字 1 作为字符串"),
        ("0", False, "Excel 数字 0 作为字符串"),
        ("yes", True, "字符串 yes"),
        ("no", False, "字符串 no"),
        ("", False, "Excel 空单元格"),
        (None, False, "pandas 读取空单元格后的 NaN（经 fillna 处理）"),
        (1, True, "Python int 1"),
        (0, False, "Python int 0"),
        (1.0, True, "Python float 1.0"),
        (0.0, False, "Python float 0.0"),
    ]

    @pytest.mark.parametrize("input_value,expected,description", EXCEL_TEST_CASES)
    def test_excel_boolean_scenarios(self, input_value, expected, description):
        layer = _make_mock_layer()
        set_layer_visibility(layer, input_value)
        assert layer.visible == expected, (
            f"{description}: 输入 {repr(input_value)} 期望 visible={expected}, "
            f"实际 visible={layer.visible}"
        )


class TestSetLayerVisibilityPerformance:
    """测试 set_layer_visibility 的批量处理性能"""

    def test_bulk_visibility_performance(self):
        """验证批量设置图层可见性的性能在合理范围内"""
        test_inputs = [
            True, False, "True", "False", "TRUE", "FALSE",
            "1", "0", "yes", "no", "on", "off",
            "", None, 1, 0, 1.0, 0.0,
            "t", "f", "y", "n", "T", "F", "Y", "N",
        ]

        iterations = 1000
        total_calls = len(test_inputs) * iterations

        start = time.perf_counter()
        for _ in range(iterations):
            for value in test_inputs:
                layer = _make_mock_layer()
                set_layer_visibility(layer, value)
        elapsed = time.perf_counter() - start

        calls_per_second = total_calls / elapsed
        print(f"Performance: {total_calls} calls in {elapsed:.3f}s "
              f"({calls_per_second:.0f} calls/s)")

        # 26000 次调用应在合理时间内完成（远低于 2 秒）
        assert elapsed < 2.0, (
            f"Performance regression: {total_calls} calls took {elapsed:.3f}s "
            f"({calls_per_second:.0f} calls/s)"
        )


if __name__ == "__main__":
    pytest.main([__file__, "-v", "-s"])
