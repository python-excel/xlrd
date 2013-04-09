from unittest import TestCase

from xlrd import open_workbook
from xlrd.book import Book
from .base import from_this_dir

class TestWorkbook(TestCase):
    def setUp(self):
        self.book = open_workbook(from_this_dir('xlsxsample.xlsx'))
        
    def test_open_workbook(self):
        assert isinstance(self.book, Book)
