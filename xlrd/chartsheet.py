# -*- coding: utf-8 -*-
# Copyright (c) 2005-2013 Stephen John Machin, Lingfo Pty Ltd
# This module is part of the xlrd package, which is released under a
# BSD-style licence.

from __future__ import print_function

from struct import unpack

from .biffh import *

DEBUG = 0
OBJ_MSO_DEBUG = 0



class ChartSheet(BaseObject):
    """
    Contains the data for one chart sheet.

    Only supports reading codename for the chart sheet at the moment.

    .. warning::

      You don't instantiate this class yourself. You access :class:`ChartSheet`
      objects via the :class:`~xlrd.book.Book` object that
      was returned when you called :func:`xlrd.open_workbook`.
    """

    #: Name of chart sheet.
    name = ''

    #: A reference to the :class:`~xlrd.book.Book` object to which this sheet
    #: belongs.
    #:
    #: Example usage: ``some_sheet.book.datemode``
    book = None

    def __init__(self, book, position, name):
        self.book = book
        self.biff_version = book.biff_version
        self._position = position
        self.logfile = book.logfile
        self.name = name
        self.verbosity = book.verbosity
        self.codename = None

    def read(self, bk):
        global rc_stats
        DEBUG = 0
        oldpos = bk._position
        bk._position = self._position
        local_unpack = unpack
        bk_get_record_parts = bk.get_record_parts
        bv = self.biff_version
        eof_found = 0
        while 1:
            # if DEBUG: print "SHEET.READ: about to read from position %d" % bk._position
            rc, data_len, data = bk_get_record_parts()
            # if rc in rc_stats:
            #     rc_stats[rc] += 1
            # else:
            #     rc_stats[rc] = 1
            # if DEBUG: print "SHEET.READ: op 0x%04x, %d bytes %r" % (rc, data_len, data)
            if rc == XL_EOF:
                DEBUG = 0
                if DEBUG: print("SHEET.READ: EOF", file=self.logfile)
                eof_found = 1
                break
            elif rc == XL_CODENAME:
                if bv < BIFF_FIRST_UNICODE:
                    self.codename = unpack_string(data, 0, bk.encoding or bk.derive_encoding(), lenlen=2)
                else:
                    self.codename = unpack_unicode(data, 0, lenlen=2)
            else:
                # if DEBUG: print "SHEET.READ: Unhandled record type %02x %d bytes %r" % (rc, data_len, data)
                pass
        if not eof_found:
            raise XLRDError("ChartSheet %r missing EOF record"
                % (self.name))
        bk._position = oldpos
        return 1
