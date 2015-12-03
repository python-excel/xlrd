### Python package "xlrd"

**Purpose**: Provide a library for developers to use to extract data from Microsoft Excel (tm) spreadsheet files. It is not an end-user tool.

**Author**: John Machin, Lingfo Pty Ltd (sjmachin@lexicon.net)

**Licence**: BSD-style (see licences.py)

**Versions of Python supported**: 2.6, 2.7, 3.2+.

**External modules required**:

<dl>

<dd>The package itself is pure Python with no dependencies on modules or packages outside the standard Python distribution.</dd>

</dl>

**Outside the current scope**: xlrd will safely and reliably ignore any of these if present in the file:

*   Charts, Macros, Pictures, any other embedded object. WARNING: currently this includes embedded worksheets.
*   VBA modules
*   Formulas (results of formula calculations are extracted, of course).
*   Comments
*   Hyperlinks
*   Autofilters, advanced filters, pivot tables, conditional formatting, data validation

**Unlikely to be done**:

*   Handling password-protected (encrypted) files.

**Particular emphasis (refer docs for details)**:

*   Operability across OS, regions, platforms
*   Handling Excel's date problems, including the Windows / Macintosh four-year differential.
*   Providing access to named constants and named groups of cells (from version 0.6.0)
*   Providing access to "visual" information: font, "number format", background, border, alignment and protection for cells, height/width etc for rows/columns (from version 0.6.1)

**Quick start**:

```
    import xlrd
    book = xlrd.open_workbook("myfile.xls")
    print "The number of worksheets is", book.nsheets
    print "Worksheet name(s):", book.sheet_names()
    sh = book.sheet_by_index(0)
    print sh.name, sh.nrows, sh.ncols
    print "Cell D30 is", sh.cell_value(rowx=29, colx=3)
    for rx in range(sh.nrows):
        print sh.row(rx)
    # Refer to docs for more details.
    # Feedback on API is welcomed.

```

**Another quick start**: This will show the first, second and last rows of each sheet in each file:

```
    OS-prompt>python PYDIR/scripts/runxlrd.py 3rows *blah*.xls
```

**Installation**:

*   On Windows: use the installer.
*   Any OS: Unzip the .zip file into a suitable directory, chdir to that directory, then do "python setup.py install".
*   If PYDIR is your Python installation directory: the main files are in PYDIR/Lib/site-packages/xlrd the docs are in the doc subdirectory, and there's a sample script: PYDIR/Scripts/runxlrd.py
*   If os.sep != "/": make the appropriate adjustments.

**Download URLs**:

*   http://pypi.python.org/pypi/xlrd
*   http://www.lexicon.net/sjmachin/xlrd.htm

**Acknowledgements**:

*   This package started life as a translation from C into Python of parts of a utility called "xlreader" developed by David Giffin. "This product includes software developed by David Giffin <david@giffin.org>."
*   OpenOffice.org has truly excellent documentation of the Microsoft Excel file formats and Compound Document file format, authored by Daniel Rentz. See http://sc.openoffice.org
*   U+5F20 U+654F: over a decade of inspiration, support, and interesting decoding opportunities.
*   Ksenia Marasanova: sample Macintosh and non-Latin1 files, alpha testing
*   Backporting to Python 2.1 was partially funded by Journyx - provider of timesheet and project accounting solutions (http://journyx.com/).
*   Provision of formatting information in version 0.6.1 was funded by Simplistix Ltd (http://www.simplistix.co.uk/)
*   << a growing list of names; see HISTORY.html >>: feedback, testing, test files, ...
