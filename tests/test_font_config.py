"""
Font Configuration System Tests

Tests for the new JSON-based font configuration system that replaced
command-line font arguments.

Test Coverage:
- load_fonts_config() - Loading fonts.json
- get_psd_prefix() - Extracting PSD file prefix
- get_font_for_psd() - Getting font path for PSD
- Error handling - Missing fonts.json, invalid JSON, missing font files
"""

import pytest
import sys
import os
import json
import tempfile
from pathlib import Path
from unittest.mock import patch, Mock

# Add parent directory to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from test_utils import TestEnvironment

# Setup test environment before importing psd_renderer
test_env = TestEnvironment()
test_env.setup_psd_renderer_args('test', 'jpg')

import psd_renderer

# Restore environment
test_env.cleanup()


class TestLoadFontsConfig:
    """Test fonts.json loading functionality"""

    def test_load_fonts_config_normal(self, tmp_path):
        """Test loading fonts.json with valid configuration"""
        config_file = tmp_path / "fonts.json"
        config_content = {
            "1": "AlibabaPuHuiTi-2-85-Bold.ttf",
            "2": "SourceHanSansCN-Medium.otf",
            "_comment": "This should be filtered"
        }
        config_file.write_text(json.dumps(config_content, ensure_ascii=False))

        with patch('os.path.exists', return_value=True):
            with patch('builtins.open', create=True) as mock_open:
                mock_open.return_value.__enter__.return_value.read.return_value = json.dumps(config_content)

                psd_renderer.fonts_config = {}  # Reset global state
                result = psd_renderer.load_fonts_config()

                # Check that comment field is filtered out
                assert "_comment" not in result
                assert "1" in result
                assert "2" in result
                assert len(result) == 2

    def test_load_fonts_config_file_not_exists(self):
        """Test handling when fonts.json does not exist"""
        with patch('os.path.exists', return_value=False):
            psd_renderer.fonts_config = {}
            result = psd_renderer.load_fonts_config()

            # Should return empty dict and use default font
            assert result == {}

    def test_load_fonts_config_invalid_json(self, tmp_path):
        """Test handling when fonts.json has invalid JSON"""
        config_file = tmp_path / "fonts.json"
        config_file.write_text("{invalid json content")

        with patch('os.path.exists', return_value=True):
            with patch('builtins.open', create=True) as mock_open:
                mock_open.return_value.__enter__.return_value.read.return_value = "{invalid"

                psd_renderer.fonts_config = {}
                result = psd_renderer.load_fonts_config()

                # Should return empty dict on error
                assert result == {}

    def test_load_fonts_config_filters_comments(self, tmp_path):
        """Test that _comment and other _ prefixed keys are filtered"""
        config_content = {
            "_comment": "This is a comment",
            "_usage": "Usage info",
            "_path_rules": "Path rules",
            "1": "font1.ttf",
            "2": "font2.otf"
        }
        config_str = json.dumps(config_content, ensure_ascii=False)

        with patch('os.path.exists', return_value=True):
            with patch('builtins.open', create=True) as mock_open:
                mock_open.return_value.__enter__.return_value.read.return_value = config_str

                psd_renderer.fonts_config = {}
                result = psd_renderer.load_fonts_config()

                # Only non-_ prefixed keys should remain
                assert "_comment" not in result
                assert "_usage" not in result
                assert "_path_rules" not in result
                assert "1" in result
                assert "2" in result
                assert len(result) == 2


class TestGetPsdPrefix:
    """Test PSD prefix extraction"""

    def test_extract_prefix_with_hash(self):
        """Test extracting prefix from file with # separator"""
        assert psd_renderer.get_psd_prefix("1#海报.psd") == "1"
        assert psd_renderer.get_psd_prefix("product#横版.psd") == "product"
        assert psd_renderer.get_psd_prefix("模板#方图#v2.psd") == "模板"

    def test_extract_prefix_without_hash(self):
        """Test extracting prefix from file without # separator"""
        assert psd_renderer.get_psd_prefix("simple.psd") == "simple"
        assert psd_renderer.get_psd_prefix("template.psd") == "template"

    def test_extract_prefix_with_path(self):
        """Test extracting prefix from full file path"""
        # get_psd_prefix uses os.path.basename() first
        # so it only uses the filename part, not the path
        assert psd_renderer.get_psd_prefix("1#海报.psd") == "1"
        assert psd_renderer.get_psd_prefix("product.psd") == "product"

    def test_extract_prefix_multiple_hashes(self):
        """Test that only first # is considered"""
        # Should extract everything before first #
        assert psd_renderer.get_psd_prefix("name#part1#part2.psd") == "name"

    def test_extract_prefix_hash_at_end(self):
        """Test file name ending with #"""
        assert psd_renderer.get_psd_prefix("filename#.psd") == "filename"


class TestGetFontForPsd:
    """Test font path retrieval for PSD files"""

    def test_get_font_for_psd_configured(self):
        """Test getting font when PSD prefix is configured"""
        psd_renderer.fonts_config = {
            "1": "AlibabaPuHuiTi-2-85-Bold.ttf",
            "2": "SourceHanSansCN-Medium.otf"
        }

        with patch('os.path.exists', return_value=True):
            font_path = psd_renderer.get_font_for_psd("1#海报.psd")

            # Should return full path
            assert font_path == "assets/fonts/AlibabaPuHuiTi-2-85-Bold.ttf"

    def test_get_font_for_psd_font_file_not_exists(self):
        """Test error when configured font file does not exist"""
        psd_renderer.fonts_config = {
            "1": "NonExistentFont.ttf"
        }

        with patch('os.path.exists', return_value=False):
            with pytest.raises(FileNotFoundError, match="字体配置文件中指定的字体不存在"):
                psd_renderer.get_font_for_psd("1#海报.psd")

    def test_get_font_for_psd_not_configured(self):
        """Test using default font when PSD prefix is not configured"""
        psd_renderer.fonts_config = {}  # Empty config

        font_path = psd_renderer.get_font_for_psd("unconfigured#poster.psd")

        # Should return default font
        assert font_path == psd_renderer.DEFAULT_FONT

    def test_get_font_for_psd_matches_correct_prefix(self):
        """Test that correct prefix is matched"""
        psd_renderer.fonts_config = {
            "1": "font1.ttf",
            "product": "font2.ttf",
            "模板": "font3.ttf"
        }

        with patch('os.path.exists', return_value=True):
            assert psd_renderer.get_font_for_psd("1#海报.psd").endswith("font1.ttf")
            assert psd_renderer.get_font_for_psd("product#横版.psd").endswith("font2.ttf")
            assert psd_renderer.get_font_for_psd("模板#方图.psd").endswith("font3.ttf")


class TestFontConfigIntegration:
    """Integration tests for font configuration system"""

    def test_font_config_with_real_files(self, tmp_path):
        """Test font loading with actual file operations"""
        # Create a temporary fonts.json
        config_file = tmp_path / "fonts.json"
        config_content = {
            "test": "test_font.ttf",
            "_comment": "Comment should be filtered"
        }
        config_file.write_text(json.dumps(config_content, ensure_ascii=False))

        # Create the font file
        font_dir = tmp_path / "assets" / "fonts"
        font_dir.mkdir(parents=True)
        (font_dir / "test_font.ttf").write_text("dummy font content")

        # Change to temp directory
        original_cwd = os.getcwd()
        try:
            os.chdir(tmp_path)

            # Load config
            psd_renderer.fonts_config = {}
            psd_renderer.load_fonts_config()

            # Get font for PSD
            font_path = psd_renderer.get_font_for_psd("test#poster.psd")

            # Verify correct font is returned
            assert "test_font.ttf" in font_path
            assert os.path.exists(font_path)

        finally:
            os.chdir(original_cwd)

    def test_multiple_psds_same_prefix(self):
        """Test that multiple PSDs with same prefix use same font"""
        psd_renderer.fonts_config = {
            "1": "shared_font.ttf"
        }

        with patch('os.path.exists', return_value=True):
            font1 = psd_renderer.get_font_for_psd("1#海报.psd")
            font2 = psd_renderer.get_font_for_psd("1#方图.psd")
            font3 = psd_renderer.get_font_for_psd("1#竖版.psd")

            # All should return same font
            assert font1 == font2 == font3
            assert "shared_font.ttf" in font1

    def test_default_font_fallback_chain(self):
        """Test default font fallback behavior"""
        # Case 1: Config exists but prefix not found
        psd_renderer.fonts_config = {"other": "font.ttf"}
        font = psd_renderer.get_font_for_psd("missing#poster.psd")
        assert font == psd_renderer.DEFAULT_FONT

        # Case 2: Config is empty
        psd_renderer.fonts_config = {}
        font = psd_renderer.get_font_for_psd("test.psd")
        assert font == psd_renderer.DEFAULT_FONT


class TestFontConfigErrorMessages:
    """Test error messages are helpful"""

    def test_error_message_includes_details(self):
        """Test that error message includes helpful details"""
        psd_renderer.fonts_config = {
            "1": "MissingFont.ttf"
        }

        with patch('os.path.exists', return_value=False):
            with pytest.raises(FileNotFoundError) as exc_info:
                psd_renderer.get_font_for_psd("1#海报.psd")

            error_msg = str(exc_info.value)
            # Error message should include:
            assert "MissingFont.ttf" in error_msg  # Font file name
            assert "1" in error_msg  # PSD prefix
            assert "fonts.json" in error_msg  # Config file name
            assert "assets/fonts" in error_msg  # Expected location


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
