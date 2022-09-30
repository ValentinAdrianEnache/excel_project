import openpyxl
import unittest
from excel_project import FILE1, FILE3


class AppX_test(unittest.TestCase):

    def test_reset_data_source_file(self):
        """
        test:delete cells source_file
        """
        wb = openpyxl.load_workbook(FILE1)
        sheet = wb.active
        no_value = sheet['A2'].value
        degree_value = sheet['B2'].value
        self.assertIsNone(no_value)
        self.assertIsNone(degree_value)

    def test_reset_data_date_hour(self):
        """
        test:delete cells date_hour
        """
        wb = openpyxl.load_workbook(FILE3)
        sheet = wb.active
        cells = ['A2', 'B2', 'C2', 'D2']
        for i in cells:
            self.assertIsNone(sheet[i].value)


if __name__ == '__main__':
    unittest.main()




