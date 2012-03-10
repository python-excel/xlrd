#!/usr/bin/env python
# encoding: utf-8
# Author:  mozman <mozman@gmx.at>
# Purpose: test cell functions
# Created: 03.12.2010
# Copyright (C) 2010, Manfred Moitzi
# License: GPLv3

import sys
import os
import unittest
from datetime import datetime, date, time

import xlrd
from xlrd import xfconst

def from_tests_dir(filename):
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)

book = xlrd.open_workbook(from_tests_dir('profiles.xls'),
                           formatting_info=True)
sheet = book.sheet_by_name('PROFILEDEF')

class TestCellValues(unittest.TestCase):
    def test_string_cell(self):
        cell = sheet.cell(0, 0)
        self.assertEqual(cell.ctype, xlrd3.XL_CELL_TEXT)
        self.assertEqual(cell.value, 'PROFIL')
        self.assertTrue(cell.has_xf)

    def test_number_cell(self):
        cell = sheet.cell(1, 1)
        self.assertEqual(cell.ctype, xlrd3.XL_CELL_NUMBER)
        self.assertEqual(cell.value, 100)
        self.assertTrue(cell.has_xf)

    def test_calculated_cell(self):
        sheet2 = book.sheet_by_name('PROFILELEVELS')
        cell = sheet2.cell(1, 3)
        self.assertEqual(cell.ctype, xlrd3.XL_CELL_NUMBER)
        self.assertAlmostEqual(cell.value, 265.131, places=3)
        self.assertTrue(cell.has_xf)

    def test_date(self):
        book = xlrd.open_workbook(from_tests_dir('Formate.xls'),
                                   formatting_info=True)
        sheet = book.sheet_by_name('Blätt1')
        for row, y, m, d in [(0, 1907, 7, 3), (1, 2005, 2, 23), (2, 1988, 5, 3)]:
            cell = sheet.cell(row, 1)
            self.assertEqual(cell.date(), date(y, m, d))

    def test_time(self):
        book = xlrd.open_workbook(from_tests_dir('Formate.xls'),
                                   formatting_info=True)
        sheet = book.sheet_by_name('Blätt1')
        for row, h, m, s in [(3, 6, 34, 0), (4, 12, 56, 0), (5, 17, 47, 13)]:
            cell = sheet.cell(row, 1)
            self.assertEqual(cell.time(), time(h, m, s))

xf_book = xlrd.open_workbook(from_tests_dir('xf_class.xls'),
                              formatting_info=True)

class TestXFCellProperties(unittest.TestCase):
    sheet = xf_book.sheet_by_name('table1')

    def cell(self, row, col):
        return self.sheet.cell(row, col)

    def test_red_background_color(self):
        # bgcolor should be 'red' #ff0000, but is not! (also xlrd 0.7.1)
        # but pattern_color is 'red'
        self.assertEqual(self.cell(0, 0).background_color(), (153, 51, 0))
        self.assertEqual(self.cell(0, 0).pattern_color(), (255, 0, 0))
        self.assertEqual(self.cell(0, 0).fill_pattern(), 1)

    def test_green_background_color(self):
        # bgcolor should be 'green' #008000, but is not! (also xlrd 0.7.1)
        # but pattern_color is 'green'
        self.assertEqual(self.cell(1, 0).background_color(), (0, 128, 128))
        self.assertEqual(self.cell(1, 0).pattern_color(), (0, 128, 0))
        self.assertEqual(self.cell(1, 0).fill_pattern(), 1)

    def test_blue_background_color(self):
        # bgcolor should be 'green' #0000ff, but is not! (also xlrd 0.7.1)
        # and pattern_color is 'blue'
        self.assertEqual(self.cell(2, 0).background_color(), (0, 128, 128))
        self.assertEqual(self.cell(2, 0).pattern_color(), (0, 102, 204))
        self.assertEqual(self.cell(2, 0).fill_pattern(), 1)

    def test_font_color(self):
        self.assertEqual(self.cell(0, 1).font_color(), (255, 0, 0))

    def test_format_str(self):
        self.assertEqual(self.cell(0, 0).format_str().upper(), "GENERAL")

    def test_horiz_alignment(self):
        self.assertEqual(self.cell(3, 0).alignment.hor_align, xfconst.HOR_ALIGN_LEFT)
        self.assertEqual(self.cell(3, 1).alignment.hor_align, xfconst.HOR_ALIGN_CENTRED)
        self.assertEqual(self.cell(3, 2).alignment.hor_align, xfconst.HOR_ALIGN_RIGHT)

    def test_vert_alignment(self):
        self.assertEqual(self.cell(4, 0).alignment.vert_align, xfconst.VERT_ALIGN_TOP)
        self.assertEqual(self.cell(4, 1).alignment.vert_align, xfconst.VERT_ALIGN_CENTRED)
        self.assertEqual(self.cell(4, 2).alignment.vert_align, xfconst.VERT_ALIGN_BOTTOM)

    def test_other_alignment(self):
        self.assertEqual(self.cell(4, 0).alignment.rotation, 0)
        self.assertEqual(self.cell(4, 0).alignment.text_wrapped, 0)
        self.assertEqual(self.cell(4, 0).alignment.indent_level, 0)
        self.assertEqual(self.cell(4, 0).alignment.shrink_to_fit, 0)
        self.assertEqual(self.cell(4, 0).alignment.text_direction, 0)

    def test_borderstyles(self):
        cell = self.cell(9, 0)
        self.assertEqual(cell.value, 'borderstyle')
        styles = cell.borderstyles()
        self.assertEqual(styles['left'], xfconst.LS_THIN)
        self.assertEqual(styles['right'], xfconst.LS_THIN)
        self.assertEqual(styles['top'], xfconst.LS_MEDIUM)
        self.assertEqual(styles['bottom'], xfconst.LS_MEDIUM)
        self.assertEqual(styles['diag'], xfconst.LS_THIN)

    def test_bordercolors(self):
        cell = self.cell(9, 0)
        self.assertEqual(cell.value, 'borderstyle')
        colors = cell.bordercolors()
        self.assertEqual(colors['left'], (0, 128, 0))
        self.assertEqual(colors['right'], (0, 128, 0))
        self.assertEqual(colors['top'], (255, 0, 0))
        self.assertEqual(colors['bottom'], (255, 0, 0))
        self.assertEqual(colors['diag'], None) # None is default color ???

    def test_diagline(self):
        cell = self.cell(9, 0)
        self.assertEqual(cell.value, 'borderstyle')
        self.assertTrue(cell.has_up_diag)
        self.assertTrue(cell.has_down_diag)

    def test_protection(self):
        cell = self.cell(0, 0)
        # 'cell_locked' and 'formula_hidden' but only if sheet is protected
        self.assertTrue(cell.is_cell_locked)
        self.assertTrue(cell.is_formula_hidden)

    def test_repr_with_formatting_info(self):
        cell = self.cell(0, 0)
        self.assertEqual(repr(cell), "text:'RED' (XF:62)")
        self.assertNotEqual(cell.xf_index, None)

    def test_repr_without_formatting_info(self):
        book = xlrd.open_workbook(from_tests_dir('xf_class.xls'),
                           formatting_info=False)
        sheet = book.sheet_by_name('table1')
        cell = sheet.cell(0, 0)
        self.assertEqual(repr(cell), "text:'RED'")
        self.assertEqual(cell.xf_index, None)

if __name__=='__main__':
    unittest.main()
