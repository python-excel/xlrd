from unittest import TestCase

import os

from xlrd import open_workbook

from .base import from_this_dir

class TestOpen(TestCase):
    # test different uses of open_workbook

    def test_open(self):
        # For now, we just check this doesn't raise an error.
        open_workbook(
            from_this_dir(os.path.join('..','xlrd','examples','namesdemo.xls'))
            )
