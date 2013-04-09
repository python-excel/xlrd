from unittest import TestCase

import os

from xlrd import open_workbook

from .base import from_this_dir

class TestOpen(TestCase):
    # test different uses of open_workbook

    def test_names_demo(self):
        # For now, we just check this doesn't raise an error.
        open_workbook(
            from_this_dir(os.path.join('..','xlrd','examples','namesdemo.xls'))
            )

    def test_ragged_rows_tidied_with_formatting(self):
        # For now, we just check this doesn't raise an error.
        open_workbook(from_this_dir('issue20.xls'),
                      formatting_info=True)
        
    def test_BYTES_X00(self):
        # For now, we just check this doesn't raise an error.
        open_workbook(from_this_dir('picture_in_cell.xls'),
                      formatting_info=True)
        
    def test_xlsx_simple(self):
        # For now, we just check this doesn't raise an error.
        open_workbook(from_this_dir('text_bar.xlsx'))
        # we should make assertions here that data has been
        # correctly processed.
        
    def test_xlsx(self):
        # For now, we just check this doesn't raise an error.
        open_workbook(from_this_dir('reveng1.xlsx'))
        # we should make assertions here that data has been
        # correctly processed.
