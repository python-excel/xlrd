#!/usr/bin/env python
# Author:  mozman <mozman@gmx.at>
# Purpose: test cell functions
# Created: 03.12.2010
# Copyright (C) 2010, Manfred Moitzi
# License: GPLv3

import sys
import os
import unittest

import xlrd

def from_tests_dir(filename):
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)

book = xlrd.open_workbook(from_tests_dir('profiles.xls'), formatting_info=True)
sheet = book.sheet_by_name('PROFILEDEF')

class TestCell(unittest.TestCase):
    def test_string_cell(self):
        cell = sheet.cell(0, 0)
        self.assertEqual(cell.ctype, xlrd.book.XL_CELL_TEXT)
        self.assertEqual(cell.value, 'PROFIL')
        self.assertTrue(cell.xf_index > 0)

    def test_number_cell(self):
        cell = sheet.cell(1, 1)
        self.assertEqual(cell.ctype, xlrd.book.XL_CELL_NUMBER)
        self.assertEqual(cell.value, 100)
        self.assertTrue(cell.xf_index > 0)

    def test_calculated_cell(self):
        sheet2 = book.sheet_by_name('PROFILELEVELS')
        cell = sheet2.cell(1, 3)
        self.assertEqual(cell.ctype, xlrd.book.XL_CELL_NUMBER)
        self.assertAlmostEqual(cell.value, 265.131, places=3)
        self.assertTrue(cell.xf_index > 0)

    def test_merged_cells(self):
        book = xlrd.open_workbook(from_tests_dir('xf_class.xls'), formatting_info=True)
        sheet3 = book.sheet_by_name('table2')
        row_lo, row_hi, col_lo, col_hi = sheet3.merged_cells[0]
        self.assertEqual(sheet3.cell(row_lo, col_lo).value, 'MERGED')
        self.assertEqual((row_lo, row_hi, col_lo, col_hi), (3, 7, 2, 5))

if __name__=='__main__':
    unittest.main()
