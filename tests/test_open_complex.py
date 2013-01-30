import sys
import os
import unittest

import xlrd

def from_this_dir(filename):
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)

class TestOpening(unittest.TestCase):
    """This opens the complex file in the examples directory, to
    exercise some more of the code.
    """
    def test_open(self):
        fname = from_this_dir(os.path.join('..','xlrd','examples','namesdemo.xls'))
        # For now, we just check this doesn't raise an error.
        xlrd.open_workbook(fname)
