#!/usr/bin/env python
#coding:utf-8
# Author:  mozman -- <mozman@gmx.at>
# Purpose: test formula (inspired by sjmachin)
# Created: 21.01.2011
# Copyright (C) , Manfred Moitzi
# License: BSD licence

import os
import sys
import unittest

import xlrd

def from_this_dir(filename):
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)

try:
    ascii
except NameError:
    # For Python 2
    def ascii(s):
        a = repr(s)
        if a.startswith(('u"', "u'")):
            a = a[1:]
        return a

book = xlrd.open_workbook(from_this_dir('formula_test_sjmachin.xls'))
sheet = book.sheet_by_index(0)

class TestFormulas(unittest.TestCase):

    def get_value(self, col, row):
        return ascii(sheet.col_values(col)[row])

    def test_is_opened(self):
        self.assertTrue(book is not None)

    def test_cell_B2(self):
        self.assertEqual(self.get_value(1, 1), r"'\u041c\u041e\u0421\u041a\u0412\u0410 \u041c\u043e\u0441\u043a\u0432\u0430'")

    def test_cell_B3(self):
        self.assertEqual(self.get_value(1, 2), '0.14285714285714285')

    def test_cell_B4(self):
        self.assertEqual(self.get_value(1, 3), "'ABCDEF'")

    def test_cell_B5(self):
        self.assertEqual(self.get_value(1, 4), "''")

    def test_cell_B6(self):
        self.assertEqual(self.get_value(1, 5), '1')

    def test_cell_B7(self):
        self.assertEqual(self.get_value(1, 6), '7')

    def test_cell_B8(self):
        self.assertEqual(self.get_value(1, 7), r"'\u041c\u041e\u0421\u041a\u0412\u0410 \u041c\u043e\u0441\u043a\u0432\u0430'")

# Adding tests in a new file, because I don't want to modify the existing test
# file using Libreoffice, in case it introduces subtle changes. -TK, Jan 2013.
book2 = xlrd.open_workbook(from_this_dir('formula_test_names.xls'))
sheet2 = book2.sheet_by_index(0)

class TestNameFormulas(unittest.TestCase):
    
    def get_value(self, col, row):
        return ascii(sheet2.col_values(col)[row])
    
    def test_is_opened(self):
        assert book is not None
    
    def test_unaryop(self):
        self.assertEqual(self.get_value(1, 1), '-7.0')
    
    def test_attrsum(self):
        self.assertEqual(self.get_value(1, 2), '4.0')
    
    def test_func(self):
        self.assertEqual(self.get_value(1, 3), '6.0')
    
    def test_func_var_args(self):
        self.assertEqual(self.get_value(1, 4), '3.0')
    
    def test_if(self):
        self.assertEqual(self.get_value(1, 5), "'b'")
    
    def test_choose(self):
        self.assertEqual(self.get_value(1, 6), "'C'")
    
    #~ def test_cell(self):
        #~ self.assertEqual(self.get_value(1, 7), self.get_value(1, 2))
    
    #~ def test_area(self):
        #~ # SUM(B2:B4) = -7 + 4 + 6 = 3
        #~ self.assertEqual(self.get_value(1, 7), "3.0")

if __name__=='__main__':
    unittest.main()
