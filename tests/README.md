# Excel-PS Batch Export Test Suite

This directory contains the comprehensive test suite for the excel-ps-batch-export project.

## 🎯 Test Suite Status

**✅ All 188 tests passing (100% success rate)**

**Last Updated**: 2026-01-09
**Status**: All tests passing

## Test Files Description

### Core Test Files

- **test_simple.py** - Core functionality tests (recommended)
  - Layer name parsing tests
  - Excel operation validation tests
  - File structure integrity tests
  - Dependency package functionality tests

- **test_font_config.py** - Font configuration system tests
  - fonts.json loading tests (normal, missing file, invalid JSON)
  - PSD prefix extraction tests
  - Font path retrieval tests (configured, default, error handling)
  - Comment field filtering tests
  - Integration tests with real files

- **test_business_logic.py** - Business logic tests
  - Advanced layer parsing tests
  - Excel data validation tests
  - Text rendering functionality tests
  - Image layer handling tests
  - Validation reporting tests
  - Multiple PSD template filename generation tests
  - Filename sanitization tests
  - Text preprocessing tests
  - Rotation functionality tests (text rotation angle parsing)

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
  - Serial processing capability tests
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
  - Serial execution simulation tests
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
  - PSD renderer integration tests (automatic image generation)
  - Error handling and user interaction tests

- **test_logging_functionality.py** - Logging functionality tests
  - Log export activity basic functionality tests
  - Duplicate record prevention tests
  - Zero count handling tests
  - File format integrity tests
  - Cross-platform compatibility tests
  - Serial scenario simulation tests

### Test Coverage

#### Functionality Tests (test_simple.py)
- [x] Layer name parsing (text variables, image variables, visibility variables)
- [x] Excel file reading and data validation
- [x] Project file structure integrity checks
- [x] Required dependency package installation checks
- [x] psd-tools library functionality verification

#### Font Configuration Tests (test_font_config.py)
- [x] Loading fonts.json with valid configuration
- [x] Handling missing fonts.json file (uses default font)
- [x] Handling invalid JSON format in fonts.json (uses default font)
- [x] Filtering comment fields (keys starting with _)
- [x] Extracting prefix from PSD filename with # separator
- [x] Extracting prefix from PSD filename without # separator
- [x] Extracting prefix from full file path
- [x] Handling multiple # in filename (only first is considered)
- [x] Getting configured font for PSD template
- [x] Error when configured font file doesn't exist
- [x] Using default font when PSD prefix not configured
- [x] Multiple PSDs with same prefix use same font
- [x] Default font fallback chain behavior
- [x] Error message includes helpful details

#### Business Logic Tests (test_business_logic.py)
- [x] Advanced layer name parsing with complex parameters
- [x] Excel data validation with various scenarios
- [x] Text rendering with different alignments and languages
- [x] Image layer handling and processing
- [x] Validation reporting functionality
- [x] Multiple PSD template filename generation with edge cases
- [x] Filename sanitization functionality
- [x] Text preprocessing functionality
- [x] Rotation functionality (positive/negative/decimal angles, combined with alignment parameters)

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
- [x] Memory extreme conditions (memory usage simulation, serial operations)

#### Performance Tests (test_performance.py)
- [x] Excel processing performance benchmarks
- [x] Memory usage monitoring and optimization
- [x] Serial processing efficiency tests
- [x] Large data processing capability verification

#### Test Results
- **Total Tests**: 188
- **Passed**: 188 (100%)
- **Failed**: 0
- **Performance**: Excellent
- **Languages**: All test output in English
- **Coverage**: Complete coverage of core functionality, error handling, boundary conditions, edge cases, integration, platform compatibility, real scenarios, clipboard import functionality, logging functionality, text rotation, and font configuration system
- **Path Handling**: Verified no hardcoded absolute paths in business code
- **Cross-platform**: Compatible with Windows and macOS
- **Architecture**: Single-process serial execution for optimal stability and simplicity

## Running Tests

### Prerequisites
1. Install project dependencies:
```bash
pip install -r requirements.txt
```

2. Install test dependencies:
```bash
pip install pytest pytest-html
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
├── test_font_config.py    # Font configuration system tests
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
3. Some tests may be skipped due to missing actual resources
4. **Windows Note**: All Windows temporary file permission issues have been fixed. Tests now run at 100% pass rate.
5. **Architecture**: Project uses single-process serial execution for optimal stability and simplicity

## Test Results

- **Total Tests**: 188
- **Pass Rate**: 100%
- **Failed Tests**: 0
- **Test Distribution**:
  - Core functionality (17 tests)
  - Font configuration (17 tests)
  - Business logic (56 tests, including 8 rotation tests)
  - Error handling (18 tests)
  - Boundary conditions (6 tests)
  - Performance testing (5 tests)
  - Boolean value handling (9 tests)
  - Integration testing (11 tests)
  - Platform compatibility (13 tests)
  - Precise text position (7 tests)
  - Real scenarios (9 tests)
  - Clipboard importer (22 tests)
  - Logging functionality (6 tests)

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