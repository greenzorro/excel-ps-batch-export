#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel-PS Batch Export Performance Tests
=====================================

This test module validates project performance including:
- Large file processing performance
- Parallel processing efficiency
- Memory usage
"""

import os
import sys
import time
import tempfile
import shutil
import psutil
import pytest
from pathlib import Path
import pandas as pd

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
        """测试PSD文件处理模拟"""
        print("\n=== PSD File Processing Simulation Test ===")
        
        # 创建模拟PSD文件
        mock_psd_file = temp_workspace / "mock.psd"
        with open(mock_psd_file, 'wb') as f:
            f.write(b'Mock PSD content' * 1000)
        
        start_time = time.time()
        start_memory = psutil.Process().memory_info().rss / 1024 / 1024
        
        # 模拟PSD处理操作
        for i in range(100):
            # 模拟图层解析
            layers = [
                f"@标题{i}#t",
                f"@背景{i}#i", 
                f"@水印{i}#v"
            ]
            for layer in layers:
                # 模拟解析操作
                if layer.startswith('@') and '#' in layer:
                    var_name = layer[1:].split('#')[0]
                    operation = layer.split('#')[1]
        
        end_time = time.time()
        end_memory = psutil.Process().memory_info().rss / 1024 / 1024
        
        duration = end_time - start_time
        memory_used = end_memory - start_memory
        
        print(f"PSD simulation processing time: {duration:.3f}s")
        print(f"Memory usage: {memory_used:.2f}MB")
        
        assert duration < 0.5, f"PSD simulation processing timetoo long: {duration:.3f}s"
        assert memory_used < 5, f"PSD模拟处理Memory usage过多: {memory_used:.2f}MB"
    
    def test_concurrent_simulation(self, perf_config, temp_workspace):
        """测试并发处理模拟"""
        print("\n=== Concurrent Processing Simulation Test ===")
        
        import threading
        import time
        
        def worker(worker_id, results):
            """工作线程"""
            start_time = time.time()
            
            # 模拟工作负载
            for i in range(1000):
                _ = i * i + worker_id
            
            end_time = time.time()
            results[worker_id] = end_time - start_time
        
        # 创建多个线程
        num_threads = 4
        threads = []
        results = {}
        
        start_time = time.time()
        
        for i in range(num_threads):
            thread = threading.Thread(target=worker, args=(i, results))
            threads.append(thread)
            thread.start()
        
        # 等待所有线程完成
        for thread in threads:
            thread.join()
        
        end_time = time.time()
        total_time = end_time - start_time
        
        print(f"Concurrent processing time: {total_time:.3f}s")
        print(f"Thread count: {num_threads}")
        print(f"Each thread processing time: {results}")
        
        assert total_time < 1.0, f"Concurrent processing timetoo long: {total_time:.3f}s"
        assert len(results) == num_threads, "不是所有线程都完成了"


if __name__ == "__main__":
    pytest.main([__file__, "-v", "-s"])