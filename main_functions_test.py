import openpyxl
import unittest
from excel_project import FILE1,FILE2, range_letter


class AppX_test(unittest.TestCase):

    def test_date_record_no_cell_source_file(self):
        """
        test:write value in no cell
        """
        wb = openpyxl.load_workbook(FILE1)
        sheet = wb.active
        no_value = sheet['A2'].value
        self.assertIsNotNone(no_value)

    def test_date_record_degree_cell_source_file(self):
        """
        test:write value in temperature cell
        """
        wb = openpyxl.load_workbook(FILE1)
        sheet = wb.active
        degree_value = sheet['B2'].value
        self.assertIsNotNone(degree_value)

    def test_Load_button(self):
        """
        test:Load button(source_file)
        """
        file1_path_split = FILE1.split('/')
        actual_name_file_1 = file1_path_split[-1]
        expected = 'source_file.xlsx'
        self.assertEqual(expected, actual_name_file_1)

    def test_Report_button(self):
        """
        test:Report file button(final_report)
        """
        file2_path_split = FILE2.split('/')
        actual_name_file_2 = file2_path_split[-1]
        expected = 'final_report.xlsx'
        self.assertEqual(expected, actual_name_file_2)

    def test_range_letter(self):
        """
        :test: range letter
        """
        start = 'A'
        stop = 'F'
        lista = [i for i in range_letter(start, stop)]
        assert lista == ['A', 'B', 'C', 'D', 'E', 'F']


if __name__ == '__main__':
    unittest.main()
