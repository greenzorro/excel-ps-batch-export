#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
transform.py 数据管道真实场景测试
================================

基于 transform.py 的实际数据管道功能编写有意义的集成测试：
- Excel I/O 性能基准
- 资源使用监控
- 串行执行验证
"""

import os
import sys
import json
import time
import tempfile
import shutil
import pytest
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from transform import (
    transform,
    transform_row,
    load_rules,
    load_raw_data,
    is_empty,
    remove_spaces,
)


class TestExcelIOBenchmark:
    """Excel I/O 性能基准测试：验证 transform 在不同数据规模下的读写性能"""

    def setup_method(self):
        self.tmpdir = tempfile.mkdtemp()
        self.workspace = os.path.join(self.tmpdir, "workspace")
        os.makedirs(self.workspace)

    def teardown_method(self):
        shutil.rmtree(self.tmpdir)

    def _write_files(self, template, json_rules, csv_data):
        json_path = os.path.join(self.workspace, f"{template}.json")
        csv_path = os.path.join(self.workspace, f"{template}_raw.csv")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(json_rules, f, ensure_ascii=False, indent=2)
        with open(csv_path, "w", encoding="utf-8") as f:
            f.write(csv_data)

    def _make_csv(self, rows, fields):
        header = ",".join(fields)
        lines = [header]
        for i in range(rows):
            line = ",".join(f"val_{field}_{i}" for field in fields)
            lines.append(line)
        return "\n".join(lines) + "\n"

    def test_write_100_rows_benchmark(self):
        """100 行数据写入基准"""
        rules = {
            "primary_field": "name",
            "columns": {
                "name": {"type": "direct", "source": "name"},
                "tag": {"type": "derived_raw", "source": "tag"},
            },
        }
        csv = "name,tag\n" + "\n".join(
            f"item_{i},{i % 2}" for i in range(100)
        ) + "\n"
        self._write_files("bench100", rules, csv)

        old_cwd = os.getcwd()
        os.chdir(self.tmpdir)
        try:
            start = time.time()
            count = transform("bench100")
            elapsed = time.time() - start

            assert count == 100
            assert elapsed < 10.0, f"100 rows took {elapsed:.2f}s, too slow"

            df = pd.read_excel(
                os.path.join(self.workspace, "bench100.xlsx"), dtype=str
            )
            assert len(df) == 100
            assert df.iloc[0]["name"] == "item_0"
            assert df.iloc[0]["tag"] == "True"
        finally:
            os.chdir(old_cwd)

    def test_write_1000_rows_benchmark(self):
        """1000 行数据写入基准"""
        rules = {
            "primary_field": "name",
            "columns": {
                "name": {"type": "direct", "source": "name"},
                "display": {
                    "type": "template",
                    "template": "{_row}-{name}",
                },
                "active": {"type": "derived", "field": "name", "when_empty": False},
            },
        }
        csv = "name\n" + "\n".join(f"row_{i}" for i in range(1000)) + "\n"
        self._write_files("bench1k", rules, csv)

        old_cwd = os.getcwd()
        os.chdir(self.tmpdir)
        try:
            start = time.time()
            count = transform("bench1k")
            elapsed = time.time() - start

            assert count == 1000
            assert elapsed < 30.0, f"1000 rows took {elapsed:.2f}s, too slow"

            df = pd.read_excel(
                os.path.join(self.workspace, "bench1k.xlsx"), dtype=str
            )
            assert len(df) == 1000
            # Verify template output for first and last rows
            assert df.iloc[0]["display"] == "1-row_0"
            assert df.iloc[999]["display"] == "1000-row_999"
            assert df.iloc[0]["active"] == "True"
        finally:
            os.chdir(old_cwd)

    def test_read_back_data_integrity(self):
        """写入后回读，验证数据完整性无丢失"""
        rules = {
            "primary_field": "id",
            "columns": {
                "id": {"type": "direct", "source": "id"},
                "chinese": {"type": "direct", "source": "chinese"},
                "number": {"type": "direct", "source": "number"},
                "has_chinese": {"type": "derived_raw", "source": "chinese"},
            },
        }
        csv_lines = ["id,chinese,number"]
        for i in range(50):
            csv_lines.append(f"ID{i},中文测试{i},{i * 100}")
        csv = "\n".join(csv_lines) + "\n"
        self._write_files("integrity", rules, csv)

        old_cwd = os.getcwd()
        os.chdir(self.tmpdir)
        try:
            transform("integrity")
            df = pd.read_excel(
                os.path.join(self.workspace, "integrity.xlsx"), dtype=str
            )

            assert len(df) == 50
            for i in range(50):
                assert df.iloc[i]["id"] == f"ID{i}"
                assert df.iloc[i]["chinese"] == f"中文测试{i}"
                assert df.iloc[i]["number"] == str(i * 100)
                assert df.iloc[i]["has_chinese"] == "True"
        finally:
            os.chdir(old_cwd)


class TestResourceUsageMonitoring:
    """资源使用监控：验证 transform 处理过程中不会产生异常的数据膨胀"""

    def setup_method(self):
        self.tmpdir = tempfile.mkdtemp()
        self.workspace = os.path.join(self.tmpdir, "workspace")
        os.makedirs(self.workspace)

    def teardown_method(self):
        shutil.rmtree(self.tmpdir)

    def _write_files(self, template, json_rules, csv_data):
        json_path = os.path.join(self.workspace, f"{template}.json")
        csv_path = os.path.join(self.workspace, f"{template}_raw.csv")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(json_rules, f, ensure_ascii=False, indent=2)
        with open(csv_path, "w", encoding="utf-8") as f:
            f.write(csv_data)

    def test_output_size_proportional_to_input(self):
        """输出文件大小应与输入行数成合理比例"""
        rules = {
            "primary_field": "text",
            "columns": {
                "text": {"type": "direct", "source": "text"},
            },
        }

        sizes_and_outputs = []

        old_cwd = os.getcwd()
        os.chdir(self.tmpdir)
        try:
            for row_count in [10, 100, 500]:
                csv = "text\n" + "\n".join(
                    f"line_{i}_" + "x" * 50 for i in range(row_count)
                ) + "\n"
                template_name = f"prop_{row_count}"
                self._write_files(template_name, rules, csv)

                transform(template_name)

                xlsx_path = os.path.join(self.workspace, f"{template_name}.xlsx")
                output_size = os.path.getsize(xlsx_path)
                sizes_and_outputs.append((row_count, output_size))

            # Output should grow roughly proportionally (not exponentially)
            # 500 rows should not be more than 100x the size of 10 rows
            ratio_10_to_500 = sizes_and_outputs[2][1] / sizes_and_outputs[0][1]
            assert ratio_10_to_500 < 100, (
                f"Output size grows too fast: "
                f"500 rows is {ratio_10_to_500:.1f}x the size of 10 rows"
            )
        finally:
            os.chdir(old_cwd)

    def test_column_count_matches_rules(self):
        """输出列数应严格等于规则定义的列数"""
        rules = {
            "primary_field": "a",
            "columns": {
                "col_a": {"type": "direct", "source": "a"},
                "col_b": {"type": "direct", "source": "b"},
                "col_flag": {"type": "derived_raw", "source": "c"},
                "col_label": {
                    "type": "template",
                    "template": "{a}-{b}",
                },
            },
        }
        csv = "a,b,c\nx,y,1\nz,w,\n"
        self._write_files("cols", rules, csv)

        old_cwd = os.getcwd()
        os.chdir(self.tmpdir)
        try:
            transform("cols")
            df = pd.read_excel(
                os.path.join(self.workspace, "cols.xlsx"), dtype=str
            )
            assert len(df.columns) == 4
            assert set(df.columns) == {"col_a", "col_b", "col_flag", "col_label"}
        finally:
            os.chdir(old_cwd)

    def test_empty_rows_excluded_from_output(self):
        """主字段为空的行应被排除，输出行数应正确"""
        rules = {
            "primary_field": "name",
            "columns": {
                "name": {"type": "direct", "source": "name"},
                "val": {"type": "direct", "source": "val"},
            },
        }
        # 5 rows total, 2 have empty primary field
        csv = "name,val\nAlice,1\n,Bob\nCharlie,2\n,Dave\nEve,3\n"
        self._write_files("guard", rules, csv)

        old_cwd = os.getcwd()
        os.chdir(self.tmpdir)
        try:
            count = transform("guard")
            assert count == 3, f"Expected 3 non-empty rows, got {count}"

            df = pd.read_excel(
                os.path.join(self.workspace, "guard.xlsx"), dtype=str
            )
            assert len(df) == 3
            assert set(df["name"].tolist()) == {"Alice", "Charlie", "Eve"}
        finally:
            os.chdir(old_cwd)


class TestSerialExecutionVerification:
    """串行执行验证：确保连续多次 transform 调用的结果一致且互不干扰"""

    def setup_method(self):
        self.tmpdir = tempfile.mkdtemp()
        self.workspace = os.path.join(self.tmpdir, "workspace")
        os.makedirs(self.workspace)

    def teardown_method(self):
        shutil.rmtree(self.tmpdir)

    def _write_files(self, template, json_rules, csv_data):
        json_path = os.path.join(self.workspace, f"{template}.json")
        csv_path = os.path.join(self.workspace, f"{template}_raw.csv")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(json_rules, f, ensure_ascii=False, indent=2)
        with open(csv_path, "w", encoding="utf-8") as f:
            f.write(csv_data)

    def test_sequential_transforms_independent(self):
        """连续执行不同模板，结果应互不影响"""
        rules_a = {
            "primary_field": "x",
            "columns": {
                "x": {"type": "direct", "source": "x"},
            },
        }
        rules_b = {
            "primary_field": "y",
            "columns": {
                "y": {"type": "direct", "source": "y"},
            },
        }
        self._write_files("seq_a", rules_a, "x\nalpha\nbeta\n")
        self._write_files("seq_b", rules_b, "y\none\ntwo\nthree\n")

        old_cwd = os.getcwd()
        os.chdir(self.tmpdir)
        try:
            count_a = transform("seq_a")
            count_b = transform("seq_b")

            assert count_a == 2
            assert count_b == 3

            df_a = pd.read_excel(
                os.path.join(self.workspace, "seq_a.xlsx"), dtype=str
            )
            df_b = pd.read_excel(
                os.path.join(self.workspace, "seq_b.xlsx"), dtype=str
            )

            assert list(df_a.columns) == ["x"]
            assert list(df_b.columns) == ["y"]
            assert df_a["x"].tolist() == ["alpha", "beta"]
            assert df_b["y"].tolist() == ["one", "two", "three"]
        finally:
            os.chdir(old_cwd)

    def test_same_template_run_twice_produces_identical_output(self):
        """同一模板执行两次，输出应完全一致（幂等性）"""
        rules = {
            "primary_field": "name",
            "columns": {
                "name": {"type": "direct", "source": "name"},
                "file": {
                    "type": "template",
                    "template": "{_row}-{name}",
                },
                "has_name": {"type": "derived", "field": "name", "when_empty": False},
            },
        }
        csv = "name\nAlice\nBob\nCharlie\n"
        self._write_files("idempotent", rules, csv)

        old_cwd = os.getcwd()
        os.chdir(self.tmpdir)
        try:
            transform("idempotent")
            df_first = pd.read_excel(
                os.path.join(self.workspace, "idempotent.xlsx"), dtype=str
            )

            transform("idempotent")
            df_second = pd.read_excel(
                os.path.join(self.workspace, "idempotent.xlsx"), dtype=str
            )

            assert df_first.equals(df_second), (
                "Two runs of the same template produced different results"
            )
        finally:
            os.chdir(old_cwd)

    def test_three_templates_serial_all_succeed(self):
        """三个模板串行执行，全部成功且数据正确"""
        templates = {}
        for idx, (tpl, src_col, values) in enumerate([
            ("s1", "city", "Beijing,Shanghai,Guangzhou"),
            ("s2", "color", "Red,Green,Blue,Yellow"),
            ("s3", "animal", "Cat,Dog"),
        ]):
            rules = {
                "primary_field": src_col,
                "columns": {
                    src_col: {"type": "direct", "source": src_col},
                    "idx": {
                        "type": "template",
                        "template": f"{idx}-{{_row}}",
                    },
                },
            }
            csv = src_col + "\n" + "\n".join(values.split(",")) + "\n"
            self._write_files(tpl, rules, csv)
            templates[tpl] = (src_col, len(values.split(",")))

        old_cwd = os.getcwd()
        os.chdir(self.tmpdir)
        try:
            for tpl, (src_col, expected_count) in templates.items():
                count = transform(tpl)
                assert count == expected_count, (
                    f"Template {tpl}: expected {expected_count} rows, got {count}"
                )

                df = pd.read_excel(
                    os.path.join(self.workspace, f"{tpl}.xlsx"), dtype=str
                )
                assert len(df) == expected_count
                assert src_col in df.columns
                assert "idx" in df.columns
        finally:
            os.chdir(old_cwd)


if __name__ == "__main__":
    pytest.main([__file__, "-v", "-s"])
