[![Build Status](https://travis-ci.org/python-excel/xlrd.svg?branch=master)](https://travis-ci.org/python-excel/xlrd)
[![Coverage Status](https://coveralls.io/repos/github/python-excel/xlrd/badge.svg?branch=master)](https://coveralls.io/github/python-excel/xlrd?branch=master)
[![Documentation Status](https://readthedocs.org/projects/xlrd/badge/?version=latest)](http://xlrd.readthedocs.io/en/latest/?badge=latest)
[![PyPI version](https://badge.fury.io/py/xlrd.svg)](https://badge.fury.io/py/xlrd)

### xlrd

PLEASE NOTE: This library currently has no active maintainers. You are advised to use [OpenPyXL](https://openpyxl.readthedocs.io/en/stable/) instead. If you absolutely have to read .xls files, then
xlrd will probably still work for you, but please do not submit issues complaining that this library
will not read your corrupted or non-standard file. Just because Excel or some other piece of software opens your
file does not mean it is a valid xls file.

For more background to this: https://groups.google.com/d/msg/python-excel/P6TjJgFVjMI/g8d0eWxTBQAJ

**Purpose**: Provide a library for developers to use to extract data from Microsoft Excel (tm) spreadsheet files. It is not an end-user tool.

**Author**: John Machin

**Licence**: BSD-style (see licences.py)

**Versions of Python supported**: 2.7, 3.4+.

**Outside scope**: xlrd will safely and reliably ignore any of these if present in the file:

*   Charts, Macros, Pictures, any other embedded object. WARNING: currently this includes embedded worksheets.
*   VBA modules
*   Formulas (results of formula calculations are extracted, of course).
*   Comments
*   Hyperlinks
*   Autofilters, advanced filters, pivot tables, conditional formatting, data validation
*   Handling password-protected (encrypted) files.

**Quick start**:

```python
import xlrd
book = xlrd.open_workbook("myfile.xls")
print("The number of worksheets is {0}".format(book.nsheets))
print("Worksheet name(s): {0}".format(book.sheet_names()))
sh = book.sheet_by_index(0)
print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
print("Cell D30 is {0}".format(sh.cell_value(rowx=29, colx=3)))
for rx in range(sh.nrows):
    print(sh.row(rx))
```

**Another quick start**: This will show the first, second and last rows of each sheet in each file:

    python PYDIR/scripts/runxlrd.py 3rows *blah*.xls

**Acknowledgements**:

*   This package started life as a translation from C into Python of parts of a utility called "xlreader" developed by David Giffin. "This product includes software developed by David Giffin <david@giffin.org>."
*   OpenOffice.org has truly excellent documentation of the Microsoft Excel file formats and Compound Document file format, authored by Daniel Rentz. See http://sc.openoffice.org
*   U+5F20 U+654F: over a decade of inspiration, support, and interesting decoding opportunities.
*   Ksenia Marasanova: sample Macintosh and non-Latin1 files, alpha testing
*   Backporting to Python 2.1 was partially funded by Journyx - provider of timesheet and project accounting solutions (http://journyx.com/).
*   Provision of formatting information in version 0.6.1 was funded by Simplistix Ltd (http://www.simplistix.co.uk/)
