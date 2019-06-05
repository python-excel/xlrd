from unittest import TestCase

from xlrd.compdoc import CompDocError

import xlrd
from .base import from_this_dir


class TestFormulas(TestCase):

    def test_not_corrupted(self):
        with self.assertRaises(Exception) as context:
            xlrd.open_workbook(from_this_dir('corrupted_error.xls'))
        self.assertTrue('Workbook corruption' in str(context.exception))

        try:
            xlrd.open_workbook(from_this_dir('corrupted_error.xls'), ignore_workbook_corruption_error=True)
        except CompDocError:
            self.fail('ignore_workbook_corruption_error=True, but the exception is raised.')
