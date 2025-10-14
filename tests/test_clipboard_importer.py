"""
Test file for clipboard_importer.py functionality

This test module covers the clipboard import functionality including:
- Clipboard data reading and parsing
- Excel file selection and target sheet detection
- Data writing to Excel with proper positioning
- Error handling and user interaction
"""

import os
import sys
import tempfile
import pandas as pd
from unittest.mock import Mock, patch, MagicMock
import pytest

# Add parent directory to path to import clipboard_importer
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Import the module to test
import clipboard_importer


class TestClipboardDataParsing:
    """Test clipboard data parsing functionality"""

    def test_parse_tab_separated_data(self):
        """Test parsing tab-separated clipboard data"""
        test_data = "姓名\t年龄\t城市\n张三\t25\t北京\n李四\t30\t上海"
        df = clipboard_importer.parse_clipboard_data(test_data)

        # After fix: all rows are data, first row is not treated as header
        assert len(df) == 3
        assert len(df.columns) == 3
        assert df.iloc[0, 0] == "姓名"  # First row is preserved
        assert df.iloc[1, 0] == "张三"
        assert df.iloc[1, 1] == "25"
        assert df.iloc[1, 2] == "北京"

    def test_parse_comma_separated_data(self):
        """Test parsing comma-separated clipboard data"""
        test_data = "姓名,年龄,城市\n张三,25,北京\n李四,30,上海"
        df = clipboard_importer.parse_clipboard_data(test_data)

        # After fix: all rows are data, first row is not treated as header
        assert len(df) == 3
        assert len(df.columns) == 3
        assert df.iloc[0, 0] == "姓名"  # First row is preserved
        assert df.iloc[1, 0] == "张三"
        assert df.iloc[1, 1] == "25"
        assert df.iloc[1, 2] == "北京"

    def test_parse_single_line_data(self):
        """Test parsing single line clipboard data"""
        test_data = "张三\t25\t北京"
        df = clipboard_importer.parse_clipboard_data(test_data)

        assert len(df) == 1
        assert len(df.columns) == 3
        assert df.iloc[0, 0] == "张三"

    def test_parse_empty_data(self):
        """Test parsing empty clipboard data"""
        with pytest.raises(ValueError, match="剪贴板数据为空"):
            clipboard_importer.parse_clipboard_data("")


class TestExcelFileSelection:
    """Test Excel file selection functionality"""

    @patch('os.listdir')
    def test_find_excel_files(self, mock_listdir):
        """Test finding Excel files in directory"""
        mock_listdir.return_value = ['test1.xlsx', 'test2.xls', 'not_excel.txt', 'test3.xlsx']

        with patch('builtins.input', return_value='1'):
            result = clipboard_importer.find_target_excel_file()
            assert result == 'test1.xlsx'

    @patch('os.listdir')
    def test_no_excel_files_found(self, mock_listdir):
        """Test when no Excel files are found"""
        mock_listdir.return_value = ['file1.txt', 'file2.pdf']

        with pytest.raises(FileNotFoundError, match="当前目录未找到Excel文件"):
            clipboard_importer.find_target_excel_file()

    @patch('os.listdir')
    def test_user_exit_selection(self, mock_listdir):
        """Test user exit during file selection"""
        mock_listdir.return_value = ['test1.xlsx', 'test2.xlsx']

        with patch('builtins.input', return_value='q'):
            with pytest.raises(SystemExit) as exc_info:
                clipboard_importer.find_target_excel_file()
            assert exc_info.value.code == 0


class TestTargetSheetDetection:
    """Test target sheet detection functionality"""

    def test_get_target_sheet_with_paste_sheet(self):
        """Test getting target sheet when '粘贴' sheet exists"""
        mock_workbook = Mock()
        mock_workbook.sheetnames = ['Sheet1', '粘贴', 'Sheet2']

        result = clipboard_importer.get_target_sheet(mock_workbook)
        assert result == "粘贴"

    def test_get_target_sheet_without_paste_sheet(self):
        """Test getting target sheet when no '粘贴' sheet exists"""
        mock_workbook = Mock()
        mock_workbook.sheetnames = ['Sheet1', 'Sheet2', 'Sheet3']

        result = clipboard_importer.get_target_sheet(mock_workbook)
        assert result == "Sheet2"

    def test_get_target_sheet_single_sheet(self):
        """Test getting target sheet when only one sheet exists"""
        mock_workbook = Mock()
        mock_workbook.sheetnames = ['Sheet1']

        result = clipboard_importer.get_target_sheet(mock_workbook)
        assert result == "Sheet1"


class TestExcelWriting:
    """Test Excel writing functionality"""

    def test_write_to_excel_clears_target_area(self):
        """Test that writing clears B2:Z1000 area before writing"""
        # Create a temporary Excel file for testing
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_file:
            tmp_path = tmp_file.name

        try:
            # Create a simple Excel file with some data
            df_test = pd.DataFrame({
                'A': ['Test1', 'Test2'],
                'B': ['Data1', 'Data2']
            })
            df_test.to_excel(tmp_path, index=False)

            # Create test data to write
            test_df = pd.DataFrame({
                'Column1': ['Value1', 'Value2'],
                'Column2': ['Data1', 'Data2']
            })

            # Mock the clipboard_importer functions
            with patch('clipboard_importer.get_target_sheet', return_value='Sheet1'):
                with patch('clipboard_importer.safe_print_message'):
                    result = clipboard_importer.write_to_excel(tmp_path, test_df)

            # Verify the function returns expected values
            assert result[0] == 'Sheet1'  # sheet name
            assert result[1] == 2  # start row (B2)
            assert result[2] == 2  # row count

        finally:
            # Clean up
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)


class TestXlwingsRecalculation:
    """Test xlwings formula recalculation functionality"""

    @patch('clipboard_importer.xw.App')
    @patch('clipboard_importer.safe_print_message')
    def test_xlwings_recalculation_success(self, mock_print, mock_app):
        """Test successful xlwings formula recalculation"""
        # Mock xlwings objects
        mock_app_instance = Mock()
        mock_wb = Mock()
        mock_app.return_value = mock_app_instance
        mock_app_instance.books.open.return_value = mock_wb
        mock_wb.app.calculate.return_value = None

        # Create a real Excel file for testing
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_file:
            tmp_path = tmp_file.name

        try:
            # Create a simple Excel file
            df_test = pd.DataFrame({'A': ['Test1'], 'B': ['Test2']})
            df_test.to_excel(tmp_path, index=False)

            # Create test data
            test_df = pd.DataFrame({'Column1': ['Value1'], 'Column2': ['Value2']})

            # Mock only the parts that need to be mocked
            with patch('clipboard_importer.get_target_sheet', return_value='Sheet1'):
                # This will trigger the xlwings recalculation
                clipboard_importer.write_to_excel(tmp_path, test_df)

            # Verify xlwings was called
            mock_app.assert_called_once_with(visible=False)
            mock_app_instance.books.open.assert_called_once_with(tmp_path)
            mock_wb.app.calculate.assert_called_once()
            mock_wb.save.assert_called_once()
            mock_wb.close.assert_called_once()
            mock_app_instance.quit.assert_called_once()

        finally:
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)

    @patch('clipboard_importer.xw.App')
    @patch('clipboard_importer.safe_print_message')
    def test_xlwings_recalculation_error_handling(self, mock_print, mock_app):
        """Test xlwings recalculation error handling"""
        # Mock xlwings to raise an exception
        mock_app.side_effect = Exception("xlwings error")

        # Create a real Excel file for testing
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_file:
            tmp_path = tmp_file.name

        try:
            # Create a simple Excel file
            df_test = pd.DataFrame({'A': ['Test1'], 'B': ['Test2']})
            df_test.to_excel(tmp_path, index=False)

            # Create test data
            test_df = pd.DataFrame({'Column1': ['Value1'], 'Column2': ['Value2']})

            # Mock only the parts that need to be mocked
            with patch('clipboard_importer.get_target_sheet', return_value='Sheet1'):
                # This should handle the xlwings error gracefully
                clipboard_importer.write_to_excel(tmp_path, test_df)

            # Verify error message was printed
            mock_print.assert_any_call("警告: 无法重新计算公式: xlwings error")
            mock_print.assert_any_call("第一个sheet的数据可能需要手动刷新")

        finally:
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)


class TestErrorHandling:
    """Test error handling functionality"""

    def test_get_clipboard_data_empty(self):
        """Test handling empty clipboard data"""
        with patch('pyperclip.paste', return_value=''):
            with pytest.raises(Exception, match="剪贴板为空"):
                clipboard_importer.get_clipboard_data()

    def test_get_clipboard_data_error(self):
        """Test handling clipboard reading errors"""
        with patch('pyperclip.paste', side_effect=Exception("Clipboard error")):
            with pytest.raises(Exception, match="无法读取剪贴板"):
                clipboard_importer.get_clipboard_data()

    def test_write_to_excel_file_not_found(self):
        """Test handling non-existent Excel file"""
        test_df = pd.DataFrame({'A': ['Test']})

        with pytest.raises(Exception, match="写入Excel文件失败"):
            clipboard_importer.write_to_excel('nonexistent.xlsx', test_df)


class TestPSDRendererIntegration:
    """Test PSD renderer integration functionality"""

    @patch('clipboard_importer.subprocess.run')
    @patch('clipboard_importer.os.path.exists')
    @patch('clipboard_importer.os.listdir')
    @patch('clipboard_importer.safe_print_message')
    def test_run_psd_renderer_success(self, mock_print, mock_listdir, mock_exists, mock_run):
        """Test successful PSD renderer execution"""
        # Setup mocks
        mock_exists.return_value = True
        # Mock listdir to return both font files and PSD files
        def listdir_side_effect(path):
            if path == 'assets/fonts':
                return ['alibaba_font.ttf', 'other_font.ttf']
            elif path == '.':
                return ['test#1.psd', 'test#2.psd']
            else:
                return []
        mock_listdir.side_effect = listdir_side_effect
        mock_run.return_value = Mock(returncode=0, stdout='', stderr='')

        # Test function
        result = clipboard_importer.run_psd_renderer('test.xlsx')

        # Verify result
        assert result == True
        mock_run.assert_called_once()
        mock_print.assert_called()

    @patch('clipboard_importer.os.path.exists')
    @patch('clipboard_importer.os.listdir')
    @patch('clipboard_importer.safe_print_message')
    def test_run_psd_renderer_no_fonts_dir(self, mock_print, mock_listdir, mock_exists):
        """Test PSD renderer when fonts directory doesn't exist"""
        # Setup mocks
        def exists_side_effect(path):
            if path == 'assets/fonts':
                return False
            else:
                return True
        mock_exists.side_effect = exists_side_effect
        # Mock listdir to return PSD files
        mock_listdir.return_value = ['test#1.psd', 'test#2.psd']

        # Test function
        result = clipboard_importer.run_psd_renderer('test.xlsx')

        # Verify result
        assert result == False
        mock_print.assert_called_with("警告: 字体目录不存在: assets/fonts")

    @patch('clipboard_importer.os.path.exists')
    @patch('clipboard_importer.os.listdir')
    @patch('clipboard_importer.safe_print_message')
    def test_run_psd_renderer_no_font_files(self, mock_print, mock_listdir, mock_exists):
        """Test PSD renderer when no font files are found"""
        # Setup mocks
        mock_exists.return_value = True
        # Mock listdir to return PSD files but no font files
        def listdir_side_effect(path):
            if path == 'assets/fonts':
                return []
            elif path == '.':
                return ['test#1.psd', 'test#2.psd']
            else:
                return []
        mock_listdir.side_effect = listdir_side_effect

        # Test function
        result = clipboard_importer.run_psd_renderer('test.xlsx')

        # Verify result
        assert result == False
        mock_print.assert_called_with("警告: 未找到字体文件")

    @patch('clipboard_importer.subprocess.run')
    @patch('clipboard_importer.os.path.exists')
    @patch('clipboard_importer.os.listdir')
    @patch('clipboard_importer.safe_print_message')
    def test_run_psd_renderer_failure(self, mock_print, mock_listdir, mock_exists, mock_run):
        """Test PSD renderer execution failure"""
        # Setup mocks
        mock_exists.return_value = True
        # Mock listdir to return both font files and PSD files
        def listdir_side_effect(path):
            if path == 'assets/fonts':
                return ['alibaba_font.ttf']
            elif path == '.':
                return ['test#1.psd', 'test#2.psd']
            else:
                return []
        mock_listdir.side_effect = listdir_side_effect
        mock_run.return_value = Mock(returncode=1, stdout='Error output', stderr='Error details')

        # Test function
        result = clipboard_importer.run_psd_renderer('test.xlsx')

        # Verify result
        assert result == False
        mock_run.assert_called_once()
        # Check that failure message was printed (not necessarily as the last call)
        mock_print.assert_any_call("\n✗ 图片渲染失败:")


class TestMainFunction:
    """Test main function execution"""

    @patch('clipboard_importer.get_clipboard_data')
    @patch('clipboard_importer.parse_clipboard_data')
    @patch('clipboard_importer.find_target_excel_file')
    @patch('clipboard_importer.write_to_excel')
    @patch('clipboard_importer.run_psd_renderer')
    @patch('clipboard_importer.safe_print_message')
    def test_main_success(self, mock_print, mock_run_psd, mock_write, mock_find, mock_parse, mock_get):
        """Test successful main function execution with PSD rendering"""
        # Setup mocks
        mock_get.return_value = "姓名\t年龄\n张三\t25"
        mock_parse.return_value = pd.DataFrame({'姓名': ['张三'], '年龄': ['25']})
        mock_find.return_value = 'test.xlsx'
        mock_write.return_value = ('Sheet1', 2, 1)
        mock_run_psd.return_value = True

        # Run main function
        result = clipboard_importer.main()

        # Verify result
        assert result == 0
        mock_write.assert_called_once()
        mock_run_psd.assert_called_once_with('test.xlsx')

    @patch('clipboard_importer.get_clipboard_data')
    @patch('clipboard_importer.safe_print_message')
    def test_main_error_handling(self, mock_print, mock_get):
        """Test main function error handling"""
        # Setup mock to raise exception
        mock_get.side_effect = Exception("Test error")

        # Run main function
        result = clipboard_importer.main()

        # Verify error handling
        assert result == 1
        mock_print.assert_called()


if __name__ == "__main__":
    pytest.main([__file__, "-v"])