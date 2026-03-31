#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel-PS Batch Export Performance Tests
=====================================

This test module validates project performance including:
- Large file processing performance
- Serial processing efficiency
- Memory usage
"""

import os
import sys
import time
import tempfile
import shutil
import pytest
from pathlib import Path
import pandas as pd

# 条件导入psutil，如果没有则使用内置模块替代
try:
    import psutil
    HAS_PSUTIL = True
except ImportError:
    HAS_PSUTIL = False
    # 使用内置模块获取进程信息
    import resource

    class MockPsutilProcess:
        """模拟psutil.Process用于测试"""
        def __init__(self):
            pass

        def memory_info(self):
            """获取内存信息"""
            usage = resource.getrusage(resource.RUSAGE_SELF)
            # rss: 实际物理内存使用（字节）
            return type('obj', (object,), {'rss': usage.ru_maxrss * 1024})()

        def cpu_percent(self, interval=None):
            """获取CPU使用率"""
            return 0.0

    psutil = type('obj', (object,), {'Process': MockPsutilProcess})()

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

class PerformanceTestConfig:
    """性能测试配置"""
    
    def __init__(self):
        self.project_root = Path(__file__).parent.parent
        self.temp_dir = Path(tempfile.mkdtemp())
        self.results_dir = self.temp_dir / "results"


@pytest.fixture
def perf_config():
    """提供性能测试配置"""
    return PerformanceTestConfig()


@pytest.fixture
def temp_workspace(perf_config):
    """提供临时工作空间"""
    perf_config.results_dir.mkdir(exist_ok=True)
    yield perf_config.temp_dir
    
    # 清理临时文件
    if perf_config.temp_dir.exists():
        shutil.rmtree(perf_config.temp_dir)


class TestPerformance:
    """性能测试类"""
    
    def test_excel_reading_performance(self, perf_config, temp_workspace):
        """测试Excel读取性能"""
        print("\n=== Excel Reading Performance Test ===")
        
        # 创建大型测试数据集
        large_data = {
            "File_name": [f"test_{i}" for i in range(1000)],
            "分类": ["测试分类"] * 1000,
            "标题第1行": ["测试标题"] * 1000,
            "标题第2行": ["副标题"] * 1000,
            "背景图": ["assets/1_img/null.jpg"] * 1000,
            "小标签": [True] * 1000,
        }
        
        df = pd.DataFrame(large_data)
        excel_file = temp_workspace / "large_test.xlsx"
        df.to_excel(excel_file, index=False)
        
        # 测试读取性能
        start_time = time.time()
        df_loaded = pd.read_excel(excel_file)
        end_time = time.time()
        
        duration = end_time - start_time
        print(f"Excel reading time: {duration:.3f}s")
        print(f"Data rows: {len(df_loaded)}")
        
        assert duration < 2.0, f"Excel reading timetoo long: {duration:.3f}s"
        assert len(df_loaded) == 1000, "Data rows不正确"
    
    def test_excel_writing_performance(self, perf_config, temp_workspace):
        """测试Excel写入性能"""
        print("\n=== Excel Writing Performance Test ===")
        
        # 创建大型测试数据集
        large_data = {
            "File_name": [f"test_{i}" for i in range(5000)],
            "分类": [f"分类_{i%10}" for i in range(5000)],
            "标题第1行": [f"测试标题_{i}" for i in range(5000)],
            "标题第2行": [f"副标题_{i}" for i in range(5000)],
            "背景图": [f"assets/1_img/img_{i%10}.jpg" for i in range(5000)],
            "小标签": [i % 2 == 0 for i in range(5000)],
        }
        
        df = pd.DataFrame(large_data)
        excel_file = temp_workspace / "large_write_test.xlsx"
        
        # 测试写入性能
        start_time = time.time()
        df.to_excel(excel_file, index=False)
        end_time = time.time()
        
        duration = end_time - start_time
        print(f"Excel writing time: {duration:.3f}s")
        print(f"Data rows: {len(df)}")
        
        assert duration < 5.0, f"Excel writing timetoo long: {duration:.3f}s"
        assert excel_file.exists(), "Excel文件未创建"
    
    def test_memory_usage(self, perf_config, temp_workspace):
        """测试Memory usage情况"""
        print("\n=== Memory Usage Test ===")
        
        process = psutil.Process()
        initial_memory = process.memory_info().rss / 1024 / 1024  # MB
        
        # 模拟内存密集型操作
        large_list = []
        for i in range(10000):
            large_list.append({"data": i, "text": "测试文本" * 100})
        
        peak_memory = process.memory_info().rss / 1024 / 1024  # MB
        
        # 清理内存
        del large_list
        
        final_memory = process.memory_info().rss / 1024 / 1024  # MB
        
        print(f"Initial memory: {initial_memory:.2f}MB")
        print(f"Peak memory: {peak_memory:.2f}MB")
        print(f"Final memory: {final_memory:.2f}MB")
        print(f"Memory growth: {peak_memory - initial_memory:.2f}MB")
        
        # Memory usage应该合理
        assert peak_memory - initial_memory < 20, f"Memory usage增长过多: {peak_memory - initial_memory:.2f}MB"
    
    def test_psd_file_simulation(self, perf_config, temp_workspace):
        """测试transform.py数据变换管线性能（100+行）"""
        print("\n=== Transform Pipeline Performance Test ===")

        from transform import transform_row

        # 构造100+行原始数据和规则
        num_rows = 150
        raw_data = pd.DataFrame({
            "File_name": [f"product_{i}" for i in range(num_rows)],
            "分类": [f"分类_{i % 5}" for i in range(num_rows)],
            "标题第1行": [f"标题_{i}" for i in range(num_rows)],
            "标题第2行": [f"副标题_{i}" for i in range(num_rows)],
            "背景图": [f"assets/1_img/img_{i % 3}.jpg" for i in range(num_rows)],
            "小标签": ["TRUE" if i % 2 == 0 else "" for i in range(num_rows)],
        })

        rules = {
            "primary_field": "File_name",
            "columns": {
                "File_name": {"type": "direct", "source": "File_name"},
                "分类": {"type": "direct", "source": "分类"},
                "标题第1行": {"type": "direct", "source": "标题第1行"},
                "标题第2行": {"type": "direct", "source": "标题第2行"},
                "文件名拼接": {
                    "type": "template",
                    "template": "{分类}-{标题第1行}",
                },
                "小标签": {"type": "direct", "source": "小标签"},
                "有小标签": {"type": "derived_raw", "source": "小标签"},
            },
        }

        start_time = time.time()
        results = []
        for idx in range(num_rows):
            row = transform_row(raw_data.iloc[idx], rules, idx + 1)
            results.append(row)
        duration = time.time() - start_time

        non_empty_results = [r for r in results if r]
        print(f"Transform pipeline: {duration:.3f}s for {num_rows} rows")
        print(f"Non-empty results: {len(non_empty_results)}")

        assert duration < 2.0, f"Transform pipeline too slow: {duration:.3f}s for {num_rows} rows"
        assert len(non_empty_results) == num_rows, f"Expected {num_rows} results, got {len(non_empty_results)}"

        # 验证结果正确性
        for r in non_empty_results[:5]:
            assert "File_name" in r, "Result should contain File_name"
            assert "文件名拼接" in r, "Result should contain template column"
    
    def test_processing_simulation(self, perf_config, temp_workspace):
        """测试批量transform_row操作性能"""
        print("\n=== Batch transform_row Performance Test ===")

        from transform import transform_row

        num_rows = 100
        raw_data = pd.DataFrame({
            "name": [f"item_{i}" for i in range(num_rows)],
            "value": [str(i * 10) for i in range(num_rows)],
            "category": [f"cat_{i % 10}" for i in range(num_rows)],
            "extra": [f"备注_{i}" if i % 3 == 0 else "" for i in range(num_rows)],
        })

        # 使用多种规则类型的规则集
        rules = {
            "primary_field": "name",
            "columns": {
                "name": {"type": "direct", "source": "name"},
                "value": {"type": "direct", "source": "value", "remove_spaces": True},
                "display": {
                    "type": "template",
                    "template": "{category}-{name}",
                },
                "extra": {"type": "conditional", "source": "extra", "depends_on": "value"},
                "has_extra": {"type": "derived_raw", "source": "extra"},
            },
        }

        start_time = time.time()
        results = []
        for idx in range(num_rows):
            row_result = transform_row(raw_data.iloc[idx], rules, idx + 1)
            results.append(row_result)
        duration = time.time() - start_time

        non_empty = [r for r in results if r]
        print(f"Batch transform_row: {duration:.3f}s for {num_rows} rows")
        print(f"Non-empty results: {len(non_empty)}")

        assert duration < 1.0, f"Batch transform_row too slow: {duration:.3f}s"
        assert len(non_empty) == num_rows

        # 验证不同规则类型都被正确执行
        sample = non_empty[0]
        assert "name" in sample, "direct rule result missing"
        assert "display" in sample, "template rule result missing"
        assert "has_extra" in sample, "derived_raw rule result missing"


if __name__ == "__main__":
    pytest.main([__file__, "-v", "-s"])