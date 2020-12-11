import os
import shutil
import tempfile
from unittest import TestCase

from xlrd import open_workbook

from .helpers import from_sample


class TestOpen(TestCase):
    # test different uses of open_workbook

    def test_names_demo(self):
        # For now, we just check this doesn't raise an error.
        open_workbook(
            from_sample(from_sample('namesdemo.xls')),
        )

    def test_ragged_rows_tidied_with_formatting(self):
        # For now, we just check this doesn't raise an error.
        open_workbook(from_sample('issue20.xls'),
                      formatting_info=True)

    def test_BYTES_X00(self):
        # For now, we just check this doesn't raise an error.
        open_workbook(from_sample('picture_in_cell.xls'),
                      formatting_info=True)
