###############################################################################
#
# Tests for Excel's Strict XML date.
#

import unittest
import xlrd
from .base import from_this_dir


class TestStrictDates(unittest.TestCase):
    # Test Excel files with dates in Excel Strict XML format. The files used
    # have the number value in the first column and the equivalent date in the
    # second column.

    def test_strict_date_1900_epoch(self):
        # Test for dates and times in the Excel standard 1900 epoch.

        workbook = xlrd.open_workbook(from_this_dir('strict_date_1900.xlsx'))
        worksheet = workbook.sheet_by_name('Sheet1')

        for row in range(19):
            num_cell = worksheet.cell(row, 0)
            date_cell = worksheet.cell(row, 1)
            self.assertAlmostEqual(num_cell.value, date_cell.value)

    def test_strict_date_1904_epoch(self):
        # Test for dates and times in the Excel for Mac 1904 epoch.

        workbook = xlrd.open_workbook(from_this_dir('strict_date_1904.xlsx'))
        worksheet = workbook.sheet_by_name('Sheet1')

        for row in range(19):
            num_cell = worksheet.cell(row, 0)
            date_cell = worksheet.cell(row, 1)
            self.assertAlmostEqual(num_cell.value, date_cell.value)
