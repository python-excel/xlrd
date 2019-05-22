# Portions Copyright (C) 2010, Manfred Moitzi under a BSD licence

from unittest import TestCase

from xlrd import open_workbook
from xlrd.book import Book

from .base import from_this_dir


class TestSharedStringsAltLocation(TestCase):

    def setUp(self):
        self.book = open_workbook(from_this_dir('sharedstrings_alt_location.xlsx'))

    def test_open_workbook(self):
        # Without the handling of the alternate location for the sharedStrings.xml file, this would pop.
        self.assertTrue(isinstance(self.book, Book))
