Python package "xlrd"
---------------------

Purpose:

    Provide a library for developers to use to extract data
    from Microsoft Excel (tm) spreadsheet files.
    
    It is not an end-user tool.

Author: John Machin, Lingfo Pty Ltd (sjmachin@lexicon.net)

Licence: BSD-style (see licences.py)

Version of xlrd: 0.5.2

Version of Python required: 2.1 or later.
    
External modules required:
    The package itself is pure Python with no dependencies on modules or packages
    outside the standard Python distribution. To run the demo script runxlrd.py with
    Python 2.2 or 2.1 requires the Optik module (version 1.4.1 or later) from 
    http://optik.sourceforge.net/
  
Versions of Excel supported:
    2004, 2002, XP, 2000, 97, 95, 5.0, 4.0, 3.0.
    2.x could be done readily enough if any demand.
    
Outside the current scope: xlrd will safely and reliably ignore any of these
if present in the file:
    * Anything to do with the on-screen presentation of the data (fonts, panes,
      column widths, row heights, ...)
    * Charts, Macros, Pictures, any other embedded object. WARNING: currently
      this includes embedded worksheets.
    * VBA modules
    * Formulas (results of formula calculations are extracted, of course).
    * Comments
    * Hyperlinks

Unlikely to be done:
    * Handling password-protected (encrypted) files.
    
Particular emphasis (refer docs for details):

    * Operability across OS, regions, platforms
      
    * Handling Excel's date problems, including the Windows / Macintosh
      four-year differential.
    
Quick start:

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

Installation:

    * On Windows: use the installer.

    * Any OS: Starting with either the .zip file or the .tar.gz file, unzip into a suitable directory,
    chdir to that directory, then do "python setup.py install".

Where did it go?

    If <PD> is your Python installation directory:
    the main files are in <PD>/Lib/site-packages/xlrd
    (except for Python 2.1 where they will be in <PD>/xlrd),
    the docs are in the doc subdirectory,
    and there's a sample script: <PD>/Scripts/runxlrd.py
    
    If os.sep != "/": make the appropriate adjustments.
    
Where did it come from?

    http://www.lexicon.net/sjmachin/xlrd.htm
    
Another quick start: This will show the first, second and last rows of each
    sheet in each file:

    OS-prompt>python <PD>/scripts/runxlrd.py 3rows *blah*.xls

Acknowledgements:

* This package started life as a translation from C into Python
of parts of a utility called "xlreader" developed by David Giffin.
"This product includes software developed by David Giffin <david@giffin.org>."

* OpenOffice.org has truly excellent documentation of the Microsoft Excel file formats
and Compound Document file format, authored by Daniel Rentz. See http://sc.openoffice.org

* U+5F20 U+654F: over a decade of inspiration, support, and interesting decoding opportunities.

* Ksenia Marasanova: sample Macintosh and non-Latin1 files, alpha testing

* Backporting to Python 2.1 was partially funded by Journyx - provider of
timesheet and project accounting solutions (http://journyx.com/).

* << a growing list of names; see HISTORY.txt >>: feedback, testing, test files, ...