#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel-PS Batch Export 图片参数测试
====================================

测试图片缩放模式和对齐参数功能。
"""

import os
import sys
import tempfile
import shutil
from pathlib import Path
from unittest.mock import Mock, MagicMock
from PIL import Image, ImageDraw
import pytest


# 添加项目根目录到路径
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))


class TestParseImageParams:
    """测试图片参数解析函数"""

    def test_default_params(self):
        """测试默认参数"""
        from psd_renderer import parse_image_params

        result = parse_image_params("@产品图#i")
        assert result["mode"] == "cover"
        assert result["alignment"] == "cm"

    def test_cover_mode(self):
        """测试 cover 模式解析"""
        from psd_renderer import parse_image_params

        result = parse_image_params("@产品图#i_cover")
        assert result["mode"] == "cover"
        assert result["alignment"] == "cm"

    def test_contain_mode(self):
        """测试 contain 模式解析"""
        from psd_renderer import parse_image_params

        result = parse_image_params("@产品图#i_contain")
        assert result["mode"] == "contain"
        assert result["alignment"] == "cm"

    def test_alignment_lt(self):
        """测试左上对齐"""
        from psd_renderer import parse_image_params

        result = parse_image_params("@产品图#i_cover_lt")
        assert result["alignment"] == "lt"

    def test_alignment_cm(self):
        """测试居中对齐"""
        from psd_renderer import parse_image_params

        result = parse_image_params("@产品图#i_cover_cm")
        assert result["alignment"] == "cm"

    def test_alignment_rb(self):
        """测试右下对齐"""
        from psd_renderer import parse_image_params

        result = parse_image_params("@产品图#i_contain_rb")
        assert result["alignment"] == "rb"

    def test_all_alignments(self):
        """测试所有九宫格对齐"""
        from psd_renderer import parse_image_params

        alignments = ["lt", "ct", "rt", "lm", "cm", "rm", "lb", "cb", "rb"]
        for align in alignments:
            result = parse_image_params(f"@产品图#i_cover_{align}")
            assert result["alignment"] == align

    def test_combined_params(self):
        """测试组合参数"""
        from psd_renderer import parse_image_params

        result = parse_image_params("@产品图#i_contain_lt")
        assert result["mode"] == "contain"
        assert result["alignment"] == "lt"

    def test_invalid_layer_name(self):
        """测试无效图层名"""
        from psd_renderer import parse_image_params

        result = parse_image_params("产品图")  # 没有 @ 和 #
        assert result["mode"] == "cover"
        assert result["alignment"] == "cm"


class TestScaleImageByMode:
    """测试按模式缩放图片函数"""

    @pytest.fixture
    def sample_images(self):
        """创建测试图片"""
        temp_dir = tempfile.mkdtemp()

        # 创建横向图 (16:9)
        landscape = Image.new("RGB", (1600, 900), color="red")
        landscape_path = os.path.join(temp_dir, "landscape.png")
        landscape.save(landscape_path)

        # 创建纵向图 (9:16)
        portrait = Image.new("RGB", (900, 1600), color="green")
        portrait_path = os.path.join(temp_dir, "portrait.png")
        portrait.save(portrait_path)

        # 创建正方形图 (1:1)
        square = Image.new("RGB", (1000, 1000), color="blue")
        square_path = os.path.join(temp_dir, "square.png")
        square.save(square_path)

        yield {
            "landscape": landscape_path,
            "portrait": portrait_path,
            "square": square_path,
            "temp_dir": temp_dir,
        }

        # 清理
        shutil.rmtree(temp_dir)

    def test_cover_mode_landscape(self, sample_images):
        """测试 cover 模式 + 横向图"""
        from psd_renderer import scale_image_by_mode

        image = Image.open(sample_images["landscape"])  # 1600x900
        result = scale_image_by_mode(image, (1000, 1000), mode="cover", alignment="cm")

        # cover 模式：按高度缩放，裁剪左右
        assert result.size == (1000, 1000)
        # 验证颜色（应该能看到红色）
        assert result.getpixel((500, 500)) == (255, 0, 0)

    def test_cover_mode_portrait(self, sample_images):
        """测试 cover 模式 + 纵向图"""
        from psd_renderer import scale_image_by_mode

        image = Image.open(sample_images["portrait"])  # 900x1600
        result = scale_image_by_mode(image, (1000, 1000), mode="cover", alignment="cm")

        # cover 模式：按宽度缩放，裁剪上下
        assert result.size == (1000, 1000)
        # 验证颜色（应该能看到绿色）
        pixel = result.getpixel((500, 500))
        assert pixel[0] == 0  # R通道为0
        assert pixel[2] == 0  # B通道为0
        assert pixel[1] > 0  # G通道大于0（绿色）

    def test_contain_mode_landscape(self, sample_images):
        """测试 contain 模式 + 横向图"""
        from psd_renderer import scale_image_by_mode

        image = Image.open(sample_images["landscape"])  # 1600x900
        result = scale_image_by_mode(
            image, (1000, 1000), mode="contain", alignment="cm"
        )

        # contain 模式：按宽度缩放，上下留白
        assert result.size == (1000, 1000)
        # 验证有透明区域（应该是 RGBA 模式）
        assert result.mode == "RGBA"
        # 中心应该有红色
        assert result.getpixel((500, 500)) == (255, 0, 0, 255)

    def test_contain_mode_portrait(self, sample_images):
        """测试 contain 模式 + 纵向图"""
        from psd_renderer import scale_image_by_mode

        image = Image.open(sample_images["portrait"])  # 900x1600
        result = scale_image_by_mode(
            image, (1000, 1000), mode="contain", alignment="cm"
        )

        # contain 模式：按高度缩放，左右留白
        assert result.size == (1000, 1000)
        # 验证有透明区域（应该是 RGBA 模式）
        assert result.mode == "RGBA"
        # 中心应该有绿色
        pixel = result.getpixel((500, 500))
        assert pixel[0] == 0  # R通道为0
        assert pixel[2] == 0  # B通道为0
        assert pixel[1] > 0  # G通道大于0（绿色）
        assert pixel[3] == 255  # Alpha通道为255（不透明）

    def test_cover_alignment_left(self, sample_images):
        """测试 cover 模式 + 左对齐（横向图）"""
        from psd_renderer import scale_image_by_mode

        image = Image.open(sample_images["landscape"])  # 1600x900
        result = scale_image_by_mode(image, (1000, 1000), mode="cover", alignment="lm")

        assert result.size == (1000, 1000)
        # 左侧应该有红色（左对齐裁剪右侧）
        assert result.getpixel((100, 500)) == (255, 0, 0)

    def test_cover_alignment_right(self, sample_images):
        """测试 cover 模式 + 右对齐（横向图）"""
        from psd_renderer import scale_image_by_mode

        image = Image.open(sample_images["landscape"])  # 1600x900
        result = scale_image_by_mode(image, (1000, 1000), mode="cover", alignment="rm")

        assert result.size == (1000, 1000)
        # 右侧应该有红色（右对齐裁剪左侧）
        assert result.getpixel((900, 500)) == (255, 0, 0)

    def test_contain_alignment_top(self, sample_images):
        """测试 contain 模式 + 上对齐（横向图）"""
        from psd_renderer import scale_image_by_mode

        image = Image.open(sample_images["landscape"])  # 1600x900
        result = scale_image_by_mode(
            image, (1000, 1000), mode="contain", alignment="lt"
        )

        assert result.size == (1000, 1000)
        assert result.mode == "RGBA"
        # 上方应该有红色（上对齐）
        assert result.getpixel((500, 100)) == (255, 0, 0, 255)
        # 下方应该是透明
        assert result.getpixel((500, 900))[3] == 0  # Alpha 通道为 0

    def test_contain_alignment_bottom(self, sample_images):
        """测试 contain 模式 + 下对齐（横向图）"""
        from psd_renderer import scale_image_by_mode

        image = Image.open(sample_images["landscape"])  # 1600x900
        result = scale_image_by_mode(
            image, (1000, 1000), mode="contain", alignment="lb"
        )

        assert result.size == (1000, 1000)
        assert result.mode == "RGBA"
        # 下方应该有红色（下对齐）
        assert result.getpixel((500, 900)) == (255, 0, 0, 255)
        # 上方应该是透明
        assert result.getpixel((500, 100))[3] == 0  # Alpha 通道为 0

    def test_square_image(self, sample_images):
        """测试正方形图"""
        from psd_renderer import scale_image_by_mode

        image = Image.open(sample_images["square"])  # 1000x1000

        # cover 模式
        result_cover = scale_image_by_mode(
            image, (800, 800), mode="cover", alignment="cm"
        )
        assert result_cover.size == (800, 800)
        assert result_cover.getpixel((400, 400)) == (0, 0, 255)

        # contain 模式
        result_contain = scale_image_by_mode(
            image, (800, 800), mode="contain", alignment="cm"
        )
        assert result_contain.size == (800, 800)
        assert result_contain.getpixel((400, 400)) == (0, 0, 255, 255)

    def test_aspect_ratio_preserved(self, sample_images):
        """测试宽高比保持"""
        from psd_renderer import scale_image_by_mode

        # 创建测试图片（16:9）
        test_image = Image.new("RGB", (160, 90), color="yellow")

        # cover 模式：缩放后裁剪，但裁剪前保持比例
        result = scale_image_by_mode(
            test_image, (100, 100), mode="cover", alignment="cm"
        )

        # 创建一个 160x90 的图片，按高度缩放到 100x56.25，然后裁剪到 100x100
        # 验证没有拉伸变形（检查颜色均匀）
        # 如果被拉伸，像素分布会不同
        assert result.size == (100, 100)

        # contain 模式：按宽度缩放到 100x56.25，上下留白
        result = scale_image_by_mode(
            test_image, (100, 100), mode="contain", alignment="cm"
        )

        assert result.size == (100, 100)
        assert result.mode == "RGBA"
        # 检查有透明区域
        alpha = result.split()[3]
        # 应该有透明像素（上下留白）
        assert alpha.getpixel((50, 95)) == 0  # 下方透明
        # 中间应该有颜色
        assert result.getpixel((50, 50)) == (255, 255, 0, 255)


class TestUpdateImageLayer:
    """测试更新图片图层函数"""

    @pytest.fixture
    def mock_layer(self):
        """创建模拟 PSD 图层"""
        layer = Mock()
        layer.name = "@产品图#i_cover_cm"
        layer.size = (1000, 1000)
        layer.offset = (100, 100)
        return layer

    @pytest.fixture
    def sample_image(self):
        """创建测试图片"""
        temp_dir = tempfile.mkdtemp()
        image_path = os.path.join(temp_dir, "test.png")
        image = Image.new("RGB", (1600, 900), color="purple")
        image.save(image_path)

        yield image_path

        shutil.rmtree(temp_dir)

    def test_update_with_valid_image(self, mock_layer, sample_image):
        """测试使用有效图片更新图层"""
        from psd_renderer import update_image_layer

        pil_image = Image.new("RGBA", (2000, 2000), (0, 0, 0, 0))

        update_image_layer(mock_layer, sample_image, pil_image)

        # 验证图层被隐藏
        assert mock_layer.visible == False

        # 验证图片被合成到正确位置
        # 检查图层位置区域有颜色
        pixel = pil_image.getpixel((600, 600))  # offset (100, 100) + 图片中心
        assert pixel != (0, 0, 0, 0)  # 不应该是透明

    def test_update_with_invalid_path(self, mock_layer):
        """测试使用无效路径"""
        from psd_renderer import update_image_layer

        pil_image = Image.new("RGBA", (2000, 2000), (0, 0, 0, 0))
        invalid_path = "/nonexistent/path/to/image.png"

        # 不应该抛出异常，只打印警告
        update_image_layer(mock_layer, invalid_path, pil_image)

        assert mock_layer.visible == False

    def test_update_with_params(self, sample_image):
        """测试使用不同参数"""
        from psd_renderer import update_image_layer

        # 测试 contain 模式
        layer_contain = Mock()
        layer_contain.name = "@产品图#i_contain_lt"
        layer_contain.size = (1000, 1000)
        layer_contain.offset = (100, 100)

        pil_image = Image.new("RGBA", (2000, 2000), (0, 0, 0, 0))
        update_image_layer(layer_contain, sample_image, pil_image)

        assert layer_contain.visible == False
        # 检查左上角有内容（左上对齐）
        assert pil_image.getpixel((150, 150))[3] > 0  # 有内容


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
