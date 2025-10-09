# Excel-PS Batch Export Test Suite

This directory contains the comprehensive test suite for the excel-ps-batch-export project.

## 🎯 Test Suite Status

**✅ All 140 tests passing (100% success rate)**

**Last Updated**: 2025-10-09
**Status**: All tests passing, clipboard importer first-row data handling verified

## Test Files Description

### Core Test Files

- **test_simple.py** - Core functionality tests (recommended)
  - Layer name parsing tests
  - Excel operation validation tests
  - File structure integrity tests
  - Dependency package functionality tests

- **test_business_logic.py** - Business logic tests
  - Advanced layer parsing tests
  - Excel data validation tests
  - Text rendering functionality tests
  - Image layer handling tests
  - Validation reporting tests

- **test_error_handling.py** - Error handling tests
  - Excel file error handling tests
  - Image data error handling tests
  - Text rendering error handling tests
  - PSD template error handling tests
  - Validation error handling tests
  - Export task error handling tests
  - Boundary condition error handling tests

- **test_boundary_conditions.py** - Boundary conditions tests
  - Extreme data conditions tests
  - Excel extreme conditions tests
  - Image extreme conditions tests
  - Text extreme conditions tests
  - Validation extreme conditions tests
  - File system extreme conditions tests
  - Memory extreme conditions tests

- **test_performance.py** - Performance tests
  - Excel processing performance tests
  - Memory usage monitoring tests
  - Concurrent processing capability tests
  - PSD file processing simulation tests

- **test_utils.py** - Test utility functions
  - Test data generation tools
  - Test environment management tools

### Specialized Test Files

- **test_boolean_issues.py** - Boolean value handling tests
  - Correct boolean parsing tests
  - Business code boolean conversion bug detection
  - Excel boolean data scenarios
  - Layer visibility boolean handling
  - Boolean parsing performance tests


- **test_integration.py** - Integration tests
  - Program startup basic tests
  - Program initialization tests
  - Datetime format handling tests
  - Command line argument parsing tests
  - File path handling tests
  - Error handling startup tests
  - Batch export script existence tests
  - Required dependencies tests
  - Program structure tests
  - Main function logic tests

- **test_platform_compatibility.py** - Platform compatibility tests
  - Windows console encoding tests
  - Chinese file path handling tests
  - Special characters in output tests
  - File path with spaces tests
  - Long file paths tests
  - Different locale settings tests
  - Error message encoding tests
  - Progress display encoding tests
  - System info detection tests
  - Environment variable handling tests
  - Unicode in Excel data tests
  - Console output buffering tests

- **test_precise_text_position.py** - Precise text position tests
  - Precise right alignment calculation tests
  - Precise center alignment calculation tests
  - Precise left alignment calculation tests
  - Mixed text precise calculation tests
  - Alignment comparison tests
  - Extreme values handling tests
  - Improved text position assertions tests

- **test_real_scenarios.py** - Real scenario tests
  - Real execution with actual files tests
  - Large dataset simulation tests
  - Concurrent execution simulation tests
  - Error recovery and continuation tests
  - Resource usage monitoring tests
  - User workflow simulation tests
  - Performance benchmark tests
  - Memory leak detection tests


- **test_clipboard_importer.py** - Clipboard importer functionality tests
  - Clipboard data parsing tests (tab/comma separated, first row preserved)
  - Excel file selection and target sheet detection tests
  - Data writing with proper positioning tests
  - Xlwings formula recalculation tests
  - Error handling and user interaction tests

- **test_logging_functionality.py** - Logging functionality tests
  - Log export activity basic functionality tests
  - Duplicate record prevention tests
  - Zero count handling tests
  - File format integrity tests
  - Cross-platform compatibility tests
  - Concurrent scenario simulation tests

### Test Coverage

#### Functionality Tests (test_simple.py)
- [x] Layer name parsing (text variables, image variables, visibility variables)
- [x] Excel file reading and data validation
- [x] Project file structure integrity checks
- [x] Required dependency package installation checks
- [x] psd-tools library functionality verification

#### Business Logic Tests (test_business_logic.py)
- [x] Advanced layer name parsing with complex parameters
- [x] Excel data validation with various scenarios
- [x] Text rendering with different alignments and languages
- [x] Image layer handling and processing
- [x] Validation reporting functionality

#### Error Handling Tests (test_error_handling.py)
- [x] Excel file error handling (missing, corrupted, empty files)
- [x] Image data error handling (invalid paths, formats, corrupted files)
- [x] Text rendering error handling (invalid fonts, empty text, special characters)
- [x] PSD template error handling (missing files, corrupted files)
- [x] Validation error handling (missing columns, invalid data)
- [x] Export task error handling (invalid data, permission errors)
- [x] Boundary condition error handling (extreme values, edge cases)

#### Boundary Conditions Tests (test_boundary_conditions.py)
- [x] Extreme data conditions (Unicode characters, maximum lengths)
- [x] Excel extreme conditions (maximum rows/columns, mixed data types)
- [x] Image extreme conditions (extreme dimensions, offsets)
- [x] Text extreme conditions (extreme font sizes, layer sizes, multilingual text)
- [x] Validation extreme conditions (large datasets, null values)
- [x] File system extreme conditions (special filenames, deep directory structures)
- [x] Memory extreme conditions (memory usage simulation, concurrent operations)

#### Performance Tests (test_performance.py)
- [x] Excel processing performance benchmarks
- [x] Memory usage monitoring and optimization
- [x] Concurrent processing efficiency tests
- [x] Large data processing capability verification

#### Test Results
- **Total Tests**: 140
- **Passed**: 140 (100%)
- **Failed**: 0
- **Performance**: Excellent
- **Languages**: All test output in English
- **Coverage**: Complete coverage of core functionality, error handling, boundary conditions, edge cases, integration, platform compatibility, real scenarios, clipboard import functionality, and logging functionality
- **Path Handling**: Verified no hardcoded absolute paths in business code
- **Cross-platform**: Compatible with Windows and macOS
- **Specialized Testing**: Boolean value handling, text position precision, platform compatibility, real-world scenario testing, clipboard import functionality with xlwings formula recalculation, and comprehensive logging functionality testing

## Running Tests

### Prerequisites
1. Install project dependencies:
```bash
pip install -r requirements.txt
```

2. Install test dependencies:
```bash
pip install pytest psutil pytest-html
```

### Running Options

#### Using Unified Test Script (Recommended)
```bash
# Run from test directory
cd test
python run_tests.py <mode>
```

Available modes:
- `basic` - Basic functionality tests
- `business` - Business logic tests
- `error` - Error handling tests
- `boundary` - Boundary conditions tests
- `perf` - Performance tests
- `all` - All tests

Examples:
```bash
# Run all tests
python run_tests.py all

# Run only performance tests
python run_tests.py perf

# Run only error handling tests
python run_tests.py error

# Run only boundary conditions tests
python run_tests.py boundary
```

#### Direct pytest Usage
```bash
# Run core functionality tests
python -m pytest test_simple.py -v

# Run business logic tests
python -m pytest test_business_logic.py -v

# Run error handling tests
python -m pytest test_error_handling.py -v

# Run boundary conditions tests
python -m pytest test_boundary_conditions.py -v

# Run performance tests
python -m pytest test_performance.py -v -s

# Run all tests
python -m pytest test/ -v

# Generate HTML report
python -m pytest test/ -v --html=test_report.html --self-contained-html
```

## Test File Structure

```
test/
├── README.md              # This document
├── run_tests.py           # Unified test runner script
├── test_simple.py         # Core functionality tests
├── test_business_logic.py # Business logic tests
├── test_error_handling.py # Error handling tests
├── test_boundary_conditions.py # Boundary conditions tests
├── test_performance.py    # Performance tests
├── test_clipboard_importer.py # Clipboard importer tests
├── test_logging_functionality.py # Logging functionality tests
└── test_utils.py          # Test utility functions
```

## Test Data

Tests use actual files from the project root directory:
- `1.psd`, `1.xlsx` - Basic functionality tests
- `3#1.psd`, `3#2.psd`, `3.xlsx` - Multi-template tests
- `assets/` - Resource files directory

## Test Configuration

Tests automatically create temporary workspaces without affecting project files. Temporary files are automatically cleaned up after tests complete.

## Notes

1. **Recommended** to use `test_simple.py` and `test_performance.py` for testing
2. Tests require PSD and Excel files from the project root directory
3. Performance tests require the psutil package
4. Some tests may be skipped due to missing actual resources
5. **Windows Note**: All Windows temporary file permission issues have been fixed. Tests now run at 100% pass rate.

## Test Results

- **Total Tests**: 140
- **Pass Rate**: 100%
- **Failed Tests**: 0
- **Test Distribution**:
  - Core functionality (17 tests) - 100% pass
  - Business logic (25 tests) - 100% pass
  - Error handling (21 tests) - 100% pass
  - Boundary conditions (6 tests) - 100% pass
  - Performance testing (5 tests) - 100% pass
  - Boolean value handling (9 tests) - 100% pass
  - Integration testing (11 tests) - 100% pass
  - Platform compatibility (13 tests) - 100% pass
  - Precise text position (7 tests) - 100% pass
  - Real scenarios (9 tests) - 100% pass
  - Clipboard importer (18 tests) - 100% pass
  - Logging functionality (6 tests) - 100% pass
- **Last Updated**: 2025-10-09
- **Status**: All tests passing, clipboard importer first-row data handling verified

## Troubleshooting

### Common Issues

1. **ImportError**: Ensure all dependency packages are installed
2. **FileNotFoundError**: Ensure test files exist in the project root directory
3. **PermissionError**: Ensure write permissions

### Debug Mode
Use `-s` parameter to run tests with detailed output:
```bash
python -m pytest -v -s
```

### Skip Slow Tests
Use `-m "not slow"` to skip tests marked as slow:
```bash
python -m pytest -v -m "not slow"
```