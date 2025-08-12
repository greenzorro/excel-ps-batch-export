#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel-PS Batch Export Test Runner
=================================

Unified test runner script supporting multiple test modes and report generation.
"""

import subprocess
import sys
import os
from pathlib import Path

def run_command(cmd, description):
    """Run command and display results"""
    print(f"\n{'='*60}")
    print(f"🚀 {description}")
    print(f"{'='*60}")
    print(f"Command: {' '.join(cmd)}")
    print(f"{'-'*60}")
    
    result = subprocess.run(cmd, capture_output=True, text=True)
    
    if result.stdout:
        print(result.stdout)
    if result.stderr:
        print("Error output:")
        print(result.stderr)
    
    return result.returncode == 0

def run_basic_tests():
    """Run basic functionality tests"""
    return run_command([
        sys.executable, "-m", "pytest", 
        "test_simple.py", "-v"
    ], "Running basic functionality tests")

def run_business_logic_tests():
    """Run business logic tests"""
    return run_command([
        sys.executable, "-m", "pytest", 
        "test_business_logic.py", "-v"
    ], "Running business logic tests")

def run_error_handling_tests():
    """Run error handling tests"""
    return run_command([
        sys.executable, "-m", "pytest", 
        "test_error_handling.py", "-v"
    ], "Running error handling tests")

def run_boundary_condition_tests():
    """Run boundary condition tests"""
    return run_command([
        sys.executable, "-m", "pytest", 
        "test_boundary_conditions.py", "-v"
    ], "Running boundary condition tests")

def run_performance_tests():
    """Run performance tests"""
    return run_command([
        sys.executable, "-m", "pytest", 
        "test_performance.py", "-v", "-s"
    ], "Running performance tests")

def run_all_tests():
    """Run all tests"""
    return run_command([
        sys.executable, "-m", "pytest", 
        "test_simple.py", "test_performance.py", 
        "test_business_logic.py", "test_error_handling.py", 
        "test_boundary_conditions.py", "-v"
    ], "Running all tests")

def run_with_html_report():
    """Generate HTML report"""
    return run_command([
        sys.executable, "-m", "pytest", 
        "test_simple.py", "test_performance.py", 
        "test_business_logic.py", "test_error_handling.py", 
        "test_boundary_conditions.py",
        "--html=test_report.html", "--self-contained-html"
    ], "Generating HTML test report")


def main():
    """Main function"""
    print("🧪 Excel-PS Batch Export Test Suite")
    print("=" * 60)
    
    if len(sys.argv) < 2:
        print("Usage:")
        print("  python run_tests.py <mode>")
        print("")
        print("Available modes:")
        print("  basic     - Basic functionality tests")
        print("  business  - Business logic tests")
        print("  error     - Error handling tests")
        print("  boundary  - Boundary condition tests")
        print("  perf      - Performance tests")
        print("  all       - All tests")
        print("  html      - Generate HTML report")
        print("")
        print("Examples:")
        print("  python run_tests.py all")
        print("  python run_tests.py html")
        print("  python run_tests.py business")
        print("  python run_tests.py error")
        return 1
    
    mode = sys.argv[1].lower()
    
    # 切换到脚本所在目录
    script_dir = Path(__file__).parent
    os.chdir(script_dir)
    
    success = False
    
    if mode == "basic":
        success = run_basic_tests()
    elif mode == "business":
        success = run_business_logic_tests()
    elif mode == "error":
        success = run_error_handling_tests()
    elif mode == "boundary":
        success = run_boundary_condition_tests()
    elif mode == "perf":
        success = run_performance_tests()
    elif mode == "all":
        success = run_all_tests()
    elif mode == "html":
        success = run_with_html_report()
    else:
        print(f"❌ Unknown mode: {mode}")
        return 1
    
    print(f"\n{'='*60}")
    if success:
        print("✅ Tests completed!")
        return 0
    else:
        print("❌ Tests failed!")
        return 1

if __name__ == "__main__":
    sys.exit(main())