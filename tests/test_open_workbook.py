import os
from unittest import TestCase

from xlrd import open_workbook

from .base import from_home_dir, from_this_dir


class TestOpen(TestCase):
    # test different uses of open_workbook

    def test_names_demo(self):
        # For now, we just check this doesn't raise an error.
        open_workbook(
            from_this_dir(os.path.join('..','examples','namesdemo.xls')),
        )

    def test_tilde_path_expansion(self):
        # For now, we just check this doesn't raise an error.
        from_home_dir('text_bar.xlsx', lambda file_path: open_workbook(file_path))

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


    def test_err_cell_empty(self):
        # For cell with type "e" (error) but without inner 'val' tags
        open_workbook(from_this_dir('err_cell_empty.xlsx'))
