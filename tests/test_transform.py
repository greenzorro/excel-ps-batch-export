#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
transform.py 数据变换引擎测试
============================

测试全部6种变换规则类型及端到端数据一致性。
"""

import os
import sys
import json
import tempfile
import shutil
import pytest
import pandas as pd

# 添加项目根目录到 Python 路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from transform import (
    is_empty, remove_spaces,
    apply_direct, apply_conditional, apply_template,
    apply_derived, apply_derived_raw,
    transform_row, transform,
)


# ── 辅助函数 ──────────────────────────────────────────────

def make_series(data: dict) -> pd.Series:
    """快速构造测试用 Series"""
    return pd.Series(data)


# ── is_empty / remove_spaces ──────────────────────────────

class TestIsEmpty:
    def test_none(self):
        assert is_empty(None) is True

    def test_empty_string(self):
        assert is_empty("") is True

    def test_whitespace_only(self):
        assert is_empty("   ") is True

    def test_null_string(self):
        assert is_empty("NULL") is True

    def test_normal_text(self):
        assert is_empty("hello") is False

    def test_zero(self):
        assert is_empty("0") is False


class TestRemoveSpaces:
    def test_no_spaces(self):
        assert remove_spaces("abc") == "abc"

    def test_with_spaces(self):
        assert remove_spaces("a b c") == "abc"

    def test_none(self):
        assert remove_spaces(None) == ""

    def test_mixed(self):
        assert remove_spaces("通识课程 - 用户运营") == "通识课程-用户运营"


# ── 模式1: direct ─────────────────────────────────────────

class TestApplyDirect:
    def test_simple_copy(self):
        rule = {"source": "name"}
        row = make_series({"name": "测试"})
        assert apply_direct(rule, row) == "测试"

    def test_remove_spaces(self):
        rule = {"source": "name", "remove_spaces": True}
        row = make_series({"name": "通识 课程"})
        assert apply_direct(rule, row) == "通识课程"

    def test_empty_becomes_empty(self):
        rule = {"source": "name"}
        row = make_series({"name": ""})
        assert apply_direct(rule, row) == ""

    def test_null_becomes_empty(self):
        rule = {"source": "name"}
        row = make_series({"name": "NULL"})
        assert apply_direct(rule, row) == ""

    def test_missing_field(self):
        rule = {"source": "nonexistent"}
        row = make_series({"name": "测试"})
        assert apply_direct(rule, row) == ""


# ── 模式2: conditional ────────────────────────────────────

class TestApplyConditional:
    def test_guard_nonempty_source_nonempty(self):
        rule = {"source": "副标题", "depends_on": "主标题", "remove_spaces": True}
        raw = make_series({"副标题": "有 值"})
        product = {"主标题": "存在"}
        assert apply_conditional(rule, raw, product) == "有值"

    def test_guard_empty_returns_empty(self):
        rule = {"source": "副标题", "depends_on": "主标题"}
        raw = make_series({"副标题": "有值"})
        product = {"主标题": ""}
        assert apply_conditional(rule, raw, product) == ""

    def test_source_empty_returns_empty(self):
        rule = {"source": "副标题", "depends_on": "主标题"}
        raw = make_series({"副标题": ""})
        product = {"主标题": "存在"}
        assert apply_conditional(rule, raw, product) == ""

    def test_source_single_space_returns_empty(self):
        rule = {"source": "副标题", "depends_on": "主标题"}
        raw = make_series({"副标题": " "})
        product = {"主标题": "存在"}
        assert apply_conditional(rule, raw, product) == ""


# ── 模式3: template ───────────────────────────────────────

class TestApplyTemplate:
    def test_basic_template(self):
        rule = {
            "template": "{_row}-{分类}-{标题}",
            "remove_spaces": ["分类"],
        }
        row = make_series({"分类": "通识 课程", "标题": "测试"})
        result = apply_template(rule, row, 1)
        assert result == "1-通识课程-测试"

    def test_skip_if_empty(self):
        rule = {
            "template": "{_row}-{标题}{副标题}",
            "skip_if_empty": ["副标题"],
        }
        row = make_series({"标题": "主标题", "副标题": ""})
        result = apply_template(rule, row, 1)
        assert result == "1-主标题"

    def test_skip_if_empty_with_dash(self):
        """测试 skip_if_empty 移除 -{field} 模式"""
        rule = {
            "template": "{_row}-{标题}-{副标题}",
            "skip_if_empty": ["副标题"],
        }
        row = make_series({"标题": "主标题", "副标题": ""})
        result = apply_template(rule, row, 1)
        assert result == "1-主标题"

    def test_no_dashes_between_titles(self):
        """测试原始公式的拼接逻辑：标题1和标题2之间没有 -"""
        rule = {
            "template": "{_row}-{分类}-{标题第1行}{标题第2行}",
            "remove_spaces": ["分类", "标题第1行"],
            "skip_if_empty": ["标题第2行"],
        }
        row = make_series({"分类": "通识课程", "标题第1行": "大促活动", "标题第2行": "技巧"})
        result = apply_template(rule, row, 3)
        assert result == "3-通识课程-大促活动技巧"

    def test_row_variable(self):
        rule = {"template": "{_row}"}
        row = make_series({})
        assert apply_template(rule, row, 42) == "42"

    def test_empty_field_removed(self):
        rule = {"template": "{a}-{b}"}
        row = make_series({"a": "hello", "b": ""})
        result = apply_template(rule, row, 1)
        # b 为空时 {b} 被替换为空，剩 "hello-"
        assert result == "hello"


# ── 模式4: derived ────────────────────────────────────────

class TestApplyDerived:
    def test_when_empty_true_field_empty(self):
        rule = {"field": "副标题", "when_empty": True}
        product = {"副标题": ""}
        assert apply_derived(rule, product) == "True"

    def test_when_empty_true_field_nonempty(self):
        rule = {"field": "副标题", "when_empty": True}
        product = {"副标题": "有值"}
        assert apply_derived(rule, product) == "False"

    def test_when_empty_false_field_empty(self):
        rule = {"field": "直播时间", "when_empty": False}
        product = {"直播时间": ""}
        assert apply_derived(rule, product) == "False"

    def test_when_empty_false_field_nonempty(self):
        rule = {"field": "直播时间", "when_empty": False}
        product = {"直播时间": "9月24日"}
        assert apply_derived(rule, product) == "True"


# ── 模式5: derived_raw ────────────────────────────────────

class TestApplyDerivedRaw:
    def test_has_value(self):
        rule = {"source": "站内标"}
        row = make_series({"站内标": "1"})
        assert apply_derived_raw(rule, row) == "True"

    def test_empty(self):
        rule = {"source": "站内标"}
        row = make_series({"站内标": ""})
        assert apply_derived_raw(rule, row) == "False"

    def test_null(self):
        rule = {"source": "站外标"}
        row = make_series({"站外标": "NULL"})
        assert apply_derived_raw(rule, row) == "False"


# ── 模式6: lookup_template (已合并到 template) ──────────

# lookup_template 功能已合并到 template 类型，以下测试验证 template
# 也能完成路径拼接场景


# ── 行级守卫 ──────────────────────────────────────────────

class TestRowGuard:
    def test_primary_empty_skips_row(self):
        rules = {
            "primary_field": "标题第1行",
            "columns": {"标题第1行": {"type": "direct", "source": "标题第1行"}},
        }
        row = make_series({"标题第1行": ""})
        assert transform_row(row, rules, 1) == {}

    def test_primary_nonempty_processes_row(self):
        rules = {
            "primary_field": "标题第1行",
            "columns": {"标题第1行": {"type": "direct", "source": "标题第1行"}},
        }
        row = make_series({"标题第1行": "测试"})
        result = transform_row(row, rules, 1)
        assert result == {"标题第1行": "测试"}


# ── _x000D_ 透传 ─────────────────────────────────────────

class TestEncodingResiduePassthrough:
    def test_x000d_preserved_in_direct(self):
        rule = {"source": "text"}
        row = make_series({"text": "hello_x000D_world"})
        assert apply_direct(rule, row) == "hello_x000D_world"

    def test_x000d_preserved_in_conditional(self):
        rule = {"source": "text", "depends_on": "guard"}
        raw = make_series({"text": "line1_x000D_line2"})
        product = {"guard": "yes"}
        assert apply_conditional(rule, raw, product) == "line1_x000D_line2"

    def test_x000d_preserved_in_template(self):
        rule = {"template": "{text}"}
        row = make_series({"text": "a_x000D_b"})
        assert apply_template(rule, row, 1) == "a_x000D_b"


# ── 端到端集成测试 ────────────────────────────────────────

class TestEndToEnd:
    """使用临时目录运行完整的 transform → 对比流程"""

    def setup_method(self):
        self.tmpdir = tempfile.mkdtemp()
        self.workspace = os.path.join(self.tmpdir, "workspace")
        os.makedirs(self.workspace)

    def teardown_method(self):
        shutil.rmtree(self.tmpdir)

    def _write_files(self, template, json_rules, csv_data):
        """写入规则和原始数据文件"""
        json_path = os.path.join(self.workspace, f"{template}.json")
        csv_path = os.path.join(self.workspace, f"{template}_raw.csv")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(json_rules, f, ensure_ascii=False, indent=2)
        with open(csv_path, "w", encoding="utf-8") as f:
            f.write(csv_data)

    def test_full_pipeline_simple(self):
        """测试最小 ETL 管道：direct + derived"""
        rules = {
            "primary_field": "name",
            "columns": {
                "name": {"type": "direct", "source": "name"},
                "active": {"type": "derived", "field": "name", "when_empty": False},
            },
        }
        csv = "name\nAlice\nBob\n"
        self._write_files("t1", rules, csv)

        # 切换到临时目录运行
        old_cwd = os.getcwd()
        os.chdir(self.tmpdir)
        try:
            count = transform("t1")
            assert count == 2

            df = pd.read_excel(
                os.path.join(self.workspace, "t1.xlsx"), dtype=str
            )
            assert len(df) == 2
            assert df.iloc[0]["name"] == "Alice"
            assert df.iloc[0]["active"] == "True"
            assert df.iloc[1]["name"] == "Bob"
        finally:
            os.chdir(old_cwd)

    def test_full_pipeline_with_primary_guard(self):
        """测试 primary_field 行级守卫跳过空行"""
        rules = {
            "primary_field": "name",
            "columns": {
                "name": {"type": "direct", "source": "name"},
            },
        }
        csv = "name\nAlice\n\nCharlie\n"
        self._write_files("t2", rules, csv)

        old_cwd = os.getcwd()
        os.chdir(self.tmpdir)
        try:
            count = transform("t2")
            assert count == 2  # 空行被跳过
        finally:
            os.chdir(old_cwd)

    def test_full_pipeline_template_and_conditional(self):
        """测试 template + conditional 组合"""
        rules = {
            "primary_field": "title",
            "columns": {
                "File_name": {
                    "type": "template",
                    "template": "{_row}-{title}",
                    "remove_spaces": ["title"],
                },
                "title": {"type": "direct", "source": "title", "remove_spaces": True},
                "subtitle": {
                    "type": "conditional",
                    "source": "subtitle",
                    "depends_on": "title",
                },
            },
        }
        csv = "title,subtitle\nMain Title,Sub\nAnother,\n"
        self._write_files("t3", rules, csv)

        old_cwd = os.getcwd()
        os.chdir(self.tmpdir)
        try:
            count = transform("t3")
            assert count == 2

            df = pd.read_excel(
                os.path.join(self.workspace, "t3.xlsx"), dtype=str
            )
            assert df.iloc[0]["File_name"] == "1-MainTitle"
            assert df.iloc[0]["subtitle"] == "Sub"
            assert str(df.iloc[1]["subtitle"]).strip() in ("", "nan")
        finally:
            os.chdir(old_cwd)

    def test_template_for_path_generation(self):
        """测试 template 生成路径（原 lookup_template 场景）"""
        rules = {
            "primary_field": "cat",
            "columns": {
                "cat": {"type": "direct", "source": "cat"},
                "bg": {
                    "type": "template",
                    "template": "img/{cat}.png",
                    "remove_spaces": ["cat"],
                },
            },
        }
        csv = "cat\nA B\nC D\n"
        self._write_files("t4", rules, csv)

        old_cwd = os.getcwd()
        os.chdir(self.tmpdir)
        try:
            transform("t4")
            df = pd.read_excel(
                os.path.join(self.workspace, "t4.xlsx"), dtype=str
            )
            assert df.iloc[0]["bg"] == "img/AB.png"
            assert df.iloc[1]["bg"] == "img/CD.png"
        finally:
            os.chdir(old_cwd)

    def test_derived_raw_boolean(self):
        """测试 derived_raw 从原始字段推导布尔值"""
        rules = {
            "primary_field": "name",
            "columns": {
                "name": {"type": "direct", "source": "name"},
                "flag": {"type": "derived_raw", "source": "flag"},
            },
        }
        csv = "name,flag\nA,1\nB,\n"
        self._write_files("t5", rules, csv)

        old_cwd = os.getcwd()
        os.chdir(self.tmpdir)
        try:
            transform("t5")
            df = pd.read_excel(
                os.path.join(self.workspace, "t5.xlsx"), dtype=str
            )
            assert df.iloc[0]["flag"] == "True"
            assert df.iloc[1]["flag"] == "False"
        finally:
            os.chdir(old_cwd)


# ── 文件缺失处理 ──────────────────────────────────────────

class TestFileErrors:
    def test_missing_rules_file(self):
        with pytest.raises(FileNotFoundError, match="规则文件不存在"):
            from transform import load_rules
            load_rules("nonexistent")

    def test_missing_raw_data(self):
        with pytest.raises(FileNotFoundError, match="原始数据文件不存在"):
            from transform import load_raw_data
            load_raw_data("nonexistent")
