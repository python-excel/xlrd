# Portions Copyright (C) 2010, Manfred Moitzi under a BSD licence

import sys
import os
import unittest

import xlrd

from .base import from_this_dir

class TestCell(unittest.TestCase):

    def setUp(self):
        self.book = xlrd.open_workbook(from_this_dir('profiles.xls'), formatting_info=True)
        self.sheet = self.book.sheet_by_name('PROFILEDEF')
        
    def test_string_cell(self):
        cell = self.sheet.cell(0, 0)
        self.assertEqual(cell.ctype, xlrd.book.XL_CELL_TEXT)
        self.assertEqual(cell.value, 'PROFIL')
        self.assertTrue(cell.xf_index > 0)

    def test_number_cell(self):
        cell = self.sheet.cell(1, 1)
        self.assertEqual(cell.ctype, xlrd.book.XL_CELL_NUMBER)
        self.assertEqual(cell.value, 100)
        self.assertTrue(cell.xf_index > 0)

    def test_calculated_cell(self):
        sheet2 = self.book.sheet_by_name('PROFILELEVELS')
        cell = sheet2.cell(1, 3)
        self.assertEqual(cell.ctype, xlrd.book.XL_CELL_NUMBER)
        self.assertAlmostEqual(cell.value, 265.131, places=3)
        self.assertTrue(cell.xf_index > 0)

    def test_merged_cells(self):
        book = xlrd.open_workbook(from_this_dir('xf_class.xls'), formatting_info=True)
        sheet3 = book.sheet_by_name('table2')
        row_lo, row_hi, col_lo, col_hi = sheet3.merged_cells[0]
        self.assertEqual(sheet3.cell(row_lo, col_lo).value, 'MERGED')
        self.assertEqual((row_lo, row_hi, col_lo, col_hi), (3, 7, 2, 5))
