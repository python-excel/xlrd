#!/usr/bin/env python
# Author:  mozman <mozman@gmx.at>
# Purpose: test xlrd basic functions
# Created: 03.12.2010
# Copyright (C) 2010, Manfred Moitzi
# License: GPLv3

import sys
import os
import unittest

import xlrd

def from_tests_dir(filename):
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)

class TestOpenWorkbook(unittest.TestCase):
    def test_open_workbook(self):
        book = xlrd.open_workbook(from_tests_dir('profiles.xls'))

class TestReadWorkbook(unittest.TestCase):
    book = xlrd.open_workbook(from_tests_dir('profiles.xls'))
    sheetnames = ['PROFILEDEF', 'AXISDEF', 'TRAVERSALCHAINAGE', 'AXISDATUMLEVELS', 'PROFILELEVELS']

    def test_nsheets(self):
        self.assertEqual(self.book.nsheets, 5)

    def test_sheet_by_name(self):
        for name in self.sheetnames:
            sheet = self.book.sheet_by_name(name)
            self.assertTrue(sheet)

    def test_sheet_by_index(self):
        for index in range(5):
            sheet = self.book.sheet_by_index(index)
            self.assertEqual(sheet.name, self.sheetnames[index])

    def test_sheets(self):
        sheets = self.book.sheets()
        for index, sheet in enumerate(sheets):
            self.assertEqual(sheet.name, self.sheetnames[index])

    def test_sheet_names(self):
        names = self.book.sheet_names()
        self.assertEqual(self.sheetnames, names)



if __name__=='__main__':
    unittest.main()
