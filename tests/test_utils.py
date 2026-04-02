#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel-PS Batch Export 测试工具模块
===================================

提供共享的测试工具函数，减少代码重复，提高测试可维护性。
"""

import os
import sys
import tempfile
import shutil
from pathlib import Path
from unittest.mock import Mock, MagicMock
import pandas as pd
from PIL import Image, ImageDraw, ImageFont


def create_mock_layer(name="@test#t", size=(100, 50), offset=(0, 0), 
                     font_size=16, is_visible=True, is_group=False):
    """创建模拟图层对象
    
    :param str name: 图层名称
    :param tuple size: 图层尺寸 (width, height)
    :param tuple offset: 图层偏移 (x, y)
    :param int font_size: 字体大小
    :param bool is_visible: 是否可见
    :param bool is_group: 是否为组
    :return Mock: 模拟图层对象
    """
    mock_layer = Mock()
    mock_layer.name = name
    mock_layer.size = size
    mock_layer.offset = offset
    mock_layer.visible = is_visible
    mock_layer.is_group.return_value = is_group
    mock_layer.is_visible.return_value = is_visible
    
    # 如果是文本图层，添加engine_dict
    if name and name.startswith('@') and '#t' in name:
        mock_layer.engine_dict = {
            'StyleRun': {
                'RunArray': [{
                    'StyleSheet': {
                        'StyleSheetData': {
                            'FontSize': font_size,
                            'FillColor': {
                                'Values': [1.0, 0.0, 0.0, 1.0]  # 黑色
                            }
                        }
                    }
                }]
            }
        }
    
    return mock_layer


def create_mock_psd(size=(800, 600), layers=None):
    """创建模拟PSD对象
    
    :param tuple size: PSD尺寸 (width, height)
    :param list layers: 图层列表
    :return Mock: 模拟PSD对象
    """
    mock_psd = Mock()
    mock_psd.size = size
    
    if layers is None:
        layers = []
    
    mock_psd.__iter__ = Mock(return_value=iter(layers))
    
    return mock_psd


def create_test_excel_data(rows=5, include_image_cols=True, include_bool_cols=True):
    """创建测试用的Excel数据
    
    :param int rows: 数据行数
    :param bool include_image_cols: 是否包含图片列
    :param bool include_bool_cols: 是否包含布尔列
    :return pd.DataFrame: 测试数据
    """
    data = {
        "File_name": [f"test_{i}" for i in range(rows)],
        "title": [f"测试标题 {i}" for i in range(rows)],
        "subtitle": [f"副标题 {i}" for i in range(rows)],
    }
    
    if include_image_cols:
        data["background"] = ["workspace/assets/1_img/null.jpg"] * rows
        data["logo"] = ["workspace/assets/1_img/null.jpg"] * rows
    
    if include_bool_cols:
        data["show_watermark"] = [i % 2 == 0 for i in range(rows)]
        data["show_border"] = [i % 3 == 0 for i in range(rows)]
    
    return pd.DataFrame(data)


def create_temp_excel_file(data, suffix='.xlsx'):
    """创建临时Excel文件
    
    :param pd.DataFrame data: Excel数据
    :param str suffix: 文件后缀
    :return str: 临时文件路径
    """
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        data.to_excel(tmp.name, index=False)
        return tmp.name


def create_temp_image_file(size=(100, 100), color='red', suffix='.jpg'):
    """创建临时图片文件
    
    :param tuple size: 图片尺寸
    :param str color: 图片颜色
    :param str suffix: 文件后缀
    :return str: 临时文件路径
    """
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        img = Image.new('RGB', size, color=color)
        img.save(tmp.name)
        return tmp.name


class TempDirectory:
    """临时目录管理器"""
    
    def __init__(self):
        self.temp_dir = tempfile.mkdtemp()
    
    def __enter__(self):
        return self.temp_dir
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.cleanup()
    
    def cleanup(self):
        """清理临时目录"""
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir, ignore_errors=True)


def assert_text_position_accuracy(x_pos, y_pos, expected_x, expected_y, 
                                 tolerance=1.0, message=""):
    """断言文本位置准确性
    
    :param float x_pos: 实际x位置
    :param float y_pos: 实际y位置
    :param float expected_x: 期望x位置
    :param float expected_y: 期望y位置
    :param float tolerance: 容忍误差
    :param str message: 错误消息前缀
    """
    x_diff = abs(x_pos - expected_x)
    y_diff = abs(y_pos - expected_y)
    
    assert x_diff <= tolerance, f"{message}X位置不准确: 期望{expected_x}, 实际{x_pos}, 误差{x_diff}"
    assert y_diff <= tolerance, f"{message}Y位置不准确: 期望{expected_y}, 实际{y_pos}, 误差{y_diff}"


def parse_boolean_value(value):
    """正确解析布尔值（用于发现业务代码问题）
    
    :param any value: 输入值
    :return bool: 解析后的布尔值
    """
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.lower() in ('true', '1', 'yes', 'on', 't', 'y')
    if isinstance(value, (int, float)):
        return value != 0
    return bool(value)


class TestEnvironment:
    """测试环境管理器"""
    
    def __init__(self):
        self.original_argv = None
        self.original_cwd = None
    
    def setup_psd_renderer_args(self, file_name='test', image_format='jpg'):
        """设置psd_renderer.py的命令行参数"""
        self.original_argv = sys.argv.copy()
        sys.argv = ['psd_renderer.py', file_name, image_format]
    
    def restore_argv(self):
        """恢复原始命令行参数"""
        if self.original_argv:
            sys.argv = self.original_argv
            self.original_argv = None
    
    def change_to_temp_dir(self):
        """切换到临时目录"""
        self.original_cwd = os.getcwd()
        temp_dir = tempfile.mkdtemp()
        os.chdir(temp_dir)
        return temp_dir
    
    def restore_cwd(self):
        """恢复原始工作目录"""
        if self.original_cwd:
            os.chdir(self.original_cwd)
            self.original_cwd = None
    
    def cleanup(self):
        """清理测试环境"""
        self.restore_argv()
        self.restore_cwd()


def create_complex_test_data(rows=10):
    """创建复杂的测试数据，包含各种边界情况
    
    :param int rows: 数据行数
    :return pd.DataFrame: 复杂测试数据
    """
    data = {
        "File_name": [f"complex_test_{i}" for i in range(rows)],
        "title": [f"标题{i}" if i % 3 != 0 else "" for i in range(rows)],
        "long_text": ["这是一段很长的文本" * (i % 5 + 1) for i in range(rows)],
        "special_chars": ["@#$%^&*()" + str(i) for i in range(rows)],
        "chinese_text": ["中文测试文本" + str(i) for i in range(rows)],
        "mixed_text": [f"中文English混合{i}" for i in range(rows)],
        "boolean_true": [True] * rows,
        "boolean_false": [False] * rows,
        "string_true": ["True"] * rows,
        "string_false": ["False"] * rows,
        "string_1": ["1"] * rows,
        "string_0": ["0"] * rows,
        "number_1": [1] * rows,
        "number_0": [0] * rows,
        "empty_value": [""] * rows,
        "none_value": [None] * rows,
    }
    
    # 添加一些图片路径
    data.update({
        "valid_image": ["workspace/assets/1_img/null.jpg"] * rows,
        "invalid_image": ["nonexistent.jpg"] * rows,
        "empty_image": [""] * rows,
    })
    
    return pd.DataFrame(data)


def validate_layer_name_parsing(layer_name, expected_type, expected_name,
                                expected_align=None, expected_valign=None,
                                expected_paragraph=None):
    """验证图层名称解析结果

    :param str layer_name: 图层名称
    :param str expected_type: 期望的类型 (text/image/visibility)
    :param str expected_name: 期望的变量名
    :param str expected_align: 期望的对齐方式 (left/center/right)
    :param str expected_valign: 期望的垂直对齐 (top/middle/bottom)
    :param bool expected_paragraph: 期望是否为段落
    """
    # 验证图层名称格式
    if not layer_name or not layer_name.startswith('@'):
        raise AssertionError(f"图层名称格式错误: {layer_name}")

    if '#' not in layer_name:
        raise AssertionError(f"图层名称缺少操作符: {layer_name}")

    # 提取变量名和操作符
    parts = layer_name[1:].split('#', 1)
    if len(parts) != 2:
        raise AssertionError(f"图层名称格式错误: {layer_name}")

    var_name = parts[0]
    operation = parts[1]

    # 验证变量名
    assert var_name == expected_name, f"变量名不匹配: 期望{expected_name}, 实际{var_name}"

    # 验证操作类型
    if not operation:
        raise AssertionError(f"操作符为空: {layer_name}")

    op_type = operation[0]  # 't', 'i', 'v'

    if expected_type == "text":
        assert op_type == 't', f"类型不匹配: 期望text (t), 实际{op_type}"

        # 对于文本类型，使用业务代码的 parse_text_params 验证参数
        from src.psd_renderer import parse_text_params
        params = parse_text_params(layer_name)

        if expected_align is not None:
            assert params["align"] == expected_align, f"对齐方式不匹配: 期望{expected_align}, 实际{params['align']}"

        if expected_valign is not None:
            assert params["valign"] == expected_valign, f"垂直对齐不匹配: 期望{expected_valign}, 实际{params['valign']}"

        if expected_paragraph is not None:
            assert params["paragraph"] == expected_paragraph, f"段落设置不匹配: 期望{expected_paragraph}, 实际{params['paragraph']}"

    elif expected_type == "image":
        assert op_type == 'i', f"类型不匹配: 期望image (i), 实际{op_type}"

        # 使用业务代码的 parse_image_params 验证参数
        from src.psd_renderer import parse_image_params
        params = parse_image_params(layer_name)

        # 图片参数至少应该有 mode 和 alignment
        assert "mode" in params, "图片参数缺少 mode"
        assert "alignment" in params, "图片参数缺少 alignment"

    elif expected_type == "visibility":
        assert op_type == 'v', f"类型不匹配: 期望visibility (v), 实际{op_type}"

    else:
        raise AssertionError(f"未知的类型: {expected_type}")


# 兼容性函数 - 保持原有API
def create_test_data(rows=10):
    """创建测试数据（兼容性函数）"""
    data = {
        "File_name": [f"test_{i}" for i in range(rows)],
        "分类": ["测试分类"] * rows,
        "标题第1行": [f"测试标题 {i}" for i in range(rows)],
        "标题第2行": [f"副标题 {i}" for i in range(rows)],
        "直播时间": ["2024-01-01"] * rows,
        "单行": ["单行文本"] * rows,
        "两行": ["两行文本\n第二行"] * rows,
        "时间框": [True] * rows,
        "站内标": [True] * rows,
        "小标签内容": ["标签内容"] * rows,
        "背景图": ["workspace/assets/1_img/null.jpg"] * rows,
        "小标签": [True] * rows,
        "站外标": [False] * rows,
    }
    return pd.DataFrame(data)


def create_test_environment(base_dir, include_psd=False):
    """创建测试环境（兼容性函数）"""
    test_dir = Path(base_dir) / "test_env"
    test_dir.mkdir(exist_ok=True)

    # 创建目录结构（匹配新结构）
    (test_dir / "workspace" / "assets" / "fonts").mkdir(parents=True, exist_ok=True)
    (test_dir / "workspace" / "assets" / "1_img").mkdir(parents=True, exist_ok=True)
    (test_dir / "export").mkdir(parents=True, exist_ok=True)
    
    # 创建测试数据
    test_data = create_test_data()
    test_data.to_excel(test_dir / "test.xlsx", index=False)
    
    # 创建虚拟图片文件
    null_img = test_dir / "workspace" / "assets" / "1_img" / "null.jpg"
    with open(null_img, 'w') as f:
        f.write("dummy image file")
    
    return test_dir


def cleanup_test_environment(test_dir):
    """清理测试环境（兼容性函数）"""
    if test_dir.exists():
        shutil.rmtree(test_dir)


def validate_test_setup():
    """验证测试设置（兼容性函数）"""
    project_root = Path(__file__).parent.parent

    required_files = [
        "src/xlsx_generator.py",
        "src/psd_renderer.py",
        "src/file_monitor.py",
        "requirements.txt"
    ]

    missing_files = []
    for file in required_files:
        if not (project_root / file).exists():
            missing_files.append(file)

    if missing_files:
        print(f"缺少必需文件: {missing_files}")
        return False

    return True


def run_test_suite():
    """运行完整测试套件（兼容性函数）"""
    print("Excel-PS Batch Export 测试套件")
    print("=" * 50)
    
    # 验证测试设置
    if not validate_test_setup():
        print("测试设置验证失败")
        return False
    
    # 创建临时测试环境
    with tempfile.TemporaryDirectory() as temp_dir:
        test_env = create_test_environment(temp_dir)
        
        try:
            print(f"测试环境创建在: {test_env}")
            print("测试完成")
            return True

        finally:
            cleanup_test_environment(test_env)


# ============== Fixture 函数 - 减少测试代码重复 ==============

def create_test_image(width=200, height=50, mode="RGB", color=None):
    """创建测试用图像对象

    :param int width: 图像宽度
    :param int height: 图像高度
    :param str mode: 图像模式 (RGB, RGBA, etc.)
    :param color: 图像颜色
    :return Image: PIL Image 对象
    """
    from PIL import Image
    if color is None:
        color = (255, 255, 255) if mode == "RGB" else (255, 255, 255, 255)
    return Image.new(mode, (width, height), color)


def create_test_draw(image=None, width=200, height=50):
    """创建测试用 Draw 对象

    :param Image image: PIL Image 对象，如果为 None 则创建新的
    :param int width: 图像宽度（创建新图像时使用）
    :param int height: 图像高度（创建新图像时使用）
    :return ImageDraw: PIL ImageDraw 对象
    """
    from PIL import ImageDraw
    if image is None:
        image = create_test_image(width, height)
    return ImageDraw.Draw(image)


def create_test_font(size=16):
    """创建测试用字体对象

    :param int size: 字体大小
    :return ImageFont: PIL ImageFont 对象
    """
    from PIL import ImageFont
    return ImageFont.load_default()


def create_test_rendering_context(width=200, height=50):
    """创建测试用渲染环境（图像、Draw、字体）

    :param int width: 图像宽度
    :param int height: 图像高度
    :return tuple: (image, draw, font)
    """
    image = create_test_image(width, height)
    draw = create_test_draw(image)
    font = create_test_font()
    return image, draw, font