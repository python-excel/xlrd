#!/usr/bin/env python
# encoding: utf-8
# Author:  mozman <mozman@gmx.at>
# Purpose: test cell formats
# Created: 03.12.2010
# Copyright (C) 2010, Manfred Moitzi
# License: GPLv3

import sys
import os
import unittest

import xlrd

def from_tests_dir(filename):
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)

book = xlrd.open_workbook(from_tests_dir('Formate.xls'), formatting_info=True)

class TestCellContent(unittest.TestCase):
    def test_text_cells(self):
        sheet = book.sheet_by_name('Blätt1')
        for row, name in enumerate(['Huber', 'Äcker', 'Öcker']):
            cell = sheet.cell(row, 0)
            self.assertEqual(cell.ctype, xlrd.XL_CELL_TEXT)
            self.assertEqual(cell.value, name)
            self.assertTrue(cell.xf_index > 0)

    def test_date_cells(self):
        sheet = book.sheet_by_name('Blätt1')
        # see also 'Dates in Excel spreadsheets' in the documentation
        # convert: xldate_as_tuple(float, book.datemode) -> (year, month,
        # day, hour, minutes, seconds)
        for row, date in [(0, 2741.), (1, 38406.), (2, 32266.)]:
            cell = sheet.cell(row, 1)
            self.assertEqual(cell.ctype, xlrd.XL_CELL_DATE)
            self.assertEqual(cell.value, date)
            self.assertTrue(cell.xf_index > 0)

    def test_time_cells(self):
        sheet = book.sheet_by_name('Blätt1')
        # see also 'Dates in Excel spreadsheets' in the documentation
        # convert: xldate_as_tuple(float, book.datemode) -> (year, month,
        # day, hour, minutes, seconds)
        for row, time in [(3, .273611), (4, .538889), (5, .741123)]:
            cell = sheet.cell(row, 1)
            self.assertEqual(cell.ctype, xlrd.XL_CELL_DATE)
            self.assertAlmostEqual(cell.value, time, places=6)
            self.assertTrue(cell.xf_index > 0)

    def test_percent_cells(self):
        sheet = book.sheet_by_name('Blätt1')
        for row, time in [(6, .974), (7, .124)]:
            cell = sheet.cell(row, 1)
            self.assertEqual(cell.ctype, xlrd.XL_CELL_NUMBER)
            self.assertAlmostEqual(cell.value, time, places=3)
            self.assertTrue(cell.xf_index > 0)

    def test_currency_cells(self):
        sheet = book.sheet_by_name('Blätt1')
        for row, time in [(8, 1000.30), (9, 1.20)]:
            cell = sheet.cell(row, 1)
            self.assertEqual(cell.ctype, xlrd.XL_CELL_NUMBER)
            self.assertAlmostEqual(cell.value, time, places=2)
            self.assertTrue(cell.xf_index > 0)

    def test_get_from_merged_cell(self):
        sheet = book.sheet_by_name('ÖÄÜ')
        cell = sheet.cell(2, 2)
        self.assertEqual(cell.ctype, xlrd.XL_CELL_TEXT)
        self.assertEqual(cell.value, 'MERGED CELLS')
        self.assertTrue(cell.xf_index > 0)

    def test_ignore_diagram(self):
        sheet = book.sheet_by_name('Blätt3')
        cell = sheet.cell(0, 0)
        self.assertEqual(cell.ctype, xlrd.XL_CELL_NUMBER)
        self.assertEqual(cell.value, 100)
        self.assertTrue(cell.xf_index > 0)

if __name__=='__main__':
    unittest.main()
