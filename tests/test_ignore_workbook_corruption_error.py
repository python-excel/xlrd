from unittest import TestCase

import xlrd

from .helpers import from_sample


class TestIgnoreWorkbookCorruption(TestCase):

    def test_not_corrupted(self):
        with self.assertRaises(Exception) as context:
            xlrd.open_workbook(from_sample('corrupted_error.xls'))
        self.assertTrue('Workbook corruption' in str(context.exception))

        xlrd.open_workbook(from_sample('corrupted_error.xls'), ignore_workbook_corruption=True)
