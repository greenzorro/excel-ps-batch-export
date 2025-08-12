# Excel-PS Batch Export Test Suite

This directory contains the comprehensive test suite for the excel-ps-batch-export project.

## Test Files Description

### Main Test Files

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
- **Total Tests**: 67
- **Passed**: 67 (100%)
- **Performance**: Excellent (comprehensive test coverage across all modules)
- **Languages**: All test output in English
- **Coverage**: Complete coverage of core functionality, error handling, boundary conditions, and edge cases

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
- `html` - Generate HTML report

Examples:
```bash
# Run all tests
python run_tests.py all

# Generate HTML report
python run_tests.py html

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
├── test_report.html       # HTML test report (latest)
├── report.html            # HTML test report (legacy)
├── test_simple.py         # Core functionality tests
├── test_business_logic.py # Business logic tests
├── test_error_handling.py # Error handling tests
├── test_boundary_conditions.py # Boundary conditions tests
├── test_performance.py    # Performance tests
├── test_utils.py          # Test utility functions
└── assets/                # Test assets directory
    └── style.css          # HTML report styles
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

## Test Results

- **Total Tests**: 67
- **Pass Rate**: 100%
- **Test Distribution**: 
  - Core functionality (13 tests)
  - Business logic (25 tests)
  - Error handling (20 tests)
  - Boundary conditions (6 tests)
  - Performance testing (5 tests)

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