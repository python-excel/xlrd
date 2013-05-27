Developers Guide
====================================

All the examples shown below can be found in the ``xlrd`` directory of the course material.

Opening Excel Files 
-------------------------------------

Workbooks can be loaded either from a :py:func:`file`, an :py:class:`~mmap.mmap` object or from a string:

::

  from mmap import mmap,ACCESS_READ
  from xlrd import open_workbook

  print open_workbook('simple.xls')

  with open('simple.xls','rb') as f:
      print open_workbook(
          file_contents=mmap(f.fileno(),0,access=ACCESS_READ)
          )

  aString = open('simple.xls','rb').read()
  print open_workbook(file_contents=aString)

Navigating a Workbook
---------------------

Here is a simple example of workbook navigation:

::

  from xlrd import open_workbook
  
  wb = open_workbook('simple.xls')
  
  for s in wb.sheets():
      print 'Sheet:',s.name
      for row in range(s.nrows):
          values = []
          for col in range(s.ncols):
              values.append(s.cell(row,col).value)
          print ','.join(values)
      print

The next few sections will cover the navigation of workbooks in more detail.

Introspecting a Book
~~~~~~~~~~~~~~~~~~~~

The :py:class:`~xlrd.book.Book` object returned by :py:func:`~xlrd.open_workbook` contains 
all information to do with the workbook and can be used to retrieve individual sheets within the workbook.

*   The :py:attr:`~xlrd.book.Book.nsheets` attribute is an integer containing the number of sheets 
    in a workbook. This attribute, in combination with the :py:meth:`~xlrd.book.Book.sheet_by_index` method, 
    is the most common way of retrieving individual sheets.
*   The :py:meth:`~xlrd.book.Book.sheet_names` method returns a list of unicodes containing the 
    names of all sheets in the workbook. Individual sheets can be retrieved using 
    these names by way of the :py:meth:`~xlrd.book.Book.sheet_by_name` function.
*   The results of the :py:meth:`~xlrd.book.Book.sheets` method can be iterated 
    over to retrieve each of the sheets in the workbook.

The <introspect_book.py> demonstrates some of these methods and attributes:

::

  from xlrd import open_workbook
  
  book = open_workbook('simple.xls')
  
  print book.nsheets
  
  for sheet_index in range(book.nsheets):
      print book.sheet_by_index(sheet_index)
  
  print book.sheet_names()
  for sheet_name in book.sheet_names():
      print book.sheet_by_name(sheet_name)
  
  for sheet in book.sheets():
      print sheet


:py:class:`~xlrd.book.Book` objects have other attributes relating to the content 
of the workbook that are only rarely useful:

* :py:class:`~xlrd.book.Book.codepage`
* :py:class:`~xlrd.book.Book.countries`
* :py:class:`~xlrd.book.Book.user_name`

Introspecting a Sheet
~~~~~~~~~~~~~~~~~~~~~~

The :py:class:`~xlrd.sheet.Sheet` object returned by any of 
the methods described above contain all the information to do with a worksheet and its contents.

*   The :py:attr:`~xlrd.sheet.Sheet.name` attribute is a unicode representing the name of the worksheet.
*   The :py:attr:`~xlrd.sheet.Sheet.nrows` and :py:attr:`~xlrd.sheet.Sheet.ncols` attributes 
    contain the number of rows and the number of columns, respectively, in the worksheet.

The following example shows how these can be used to iterate over and display the contents of one worksheet:

::

  from xlrd import open_workbook, cellname
  
  book = open_workbook('odd.xls')
  sheet = book.sheet_by_index(0)
  
  print sheet.name
  
  print sheet.nrows
  print sheet.ncols
  
  for row_index in range(sheet.nrows):
      for col_index in range(sheet.ncols):
          print cellname(row_index,col_index),'-',
          print sheet.cell(row_index,col_index).value
  introspect_sheet.py

:py:class:`~xlrd.sheet.Sheet` objects have other attributes relating to the content of the worksheet 
that are only rarely useful:

* :py:attr:`~xlrd.sheet.Sheet.col_label_ranges`
* :py:attr:`~xlrd.sheet.Sheet.row_label_ranges`
* :py:attr:`~xlrd.sheet.Sheet.visibility`


Getting a particular Cell
~~~~~~~~~~~~~~~~~~~~~~~~~

*   The :py:meth:`~xlrd.sheet.Sheet.cell` method of a :py:class:`~xlrd.sheet.Sheet` object 
    can be used to return the contents of a particular cell.
*   The :py:meth:`~xlrd.sheet.Sheet.cell` method returns 
    a :py:class:`~xlrd.sheet.Cell` object. 
*   These objects have very few  attributes, of which  :py:attr:`~xlrd.sheet.Cell.value` contains 
    the actual value of the cell and :py:attr:`~xlrd.sheet.Cell.ctype` contains the type of the cell.

In addition, :py:class:`~xlrd.sheet.Sheet` objects have two methods for returning these two types of data. 

*   The :py:meth:`~xlrd.sheet.Sheet.cell_value` method returns the value for a particular cell
*   The :py:meth:`~xlrd.sheet.Sheet.cell_type` method returns the type of a particular cell. 
*   These methods can be quicker to execute than retrieving the :py:meth:`~xlrd.sheet.Cell` object.

Cell types are covered in more detail later. The following example shows the methods, attributes and classes in action:

::

  from xlrd import open_workbook,XL_CELL_TEXT
  
  book = open_workbook('odd.xls')
  sheet = book.sheet_by_index(1)
  
  cell = sheet.cell(0,0)
  print cell
  print cell.value
  print cell.ctype==XL_CELL_TEXT
  
  for i in range(sheet.ncols):
      print sheet.cell_type(1,i),sheet.cell_value(1,i)
  cell_access.py

Iterating over the contents of a Sheet
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

We've already seen how to iterate over the contents of a worksheet and retrieve the resulting individual cells. However, there are methods to retrieve groups of cells more easily. 

There are a symmetrical set of methods that retrieve groups of cell information either by row or by column.

*   The :py:meth:`~xlrd.sheet.Sheet.row` 
    and :py:meth:`~xlrd.sheet.Sheet.col` methods return all 
    the :py:class:`~xlrd.sheet.Sheet.Cell` objects for a whole row or column, respectively.

*   The :py:meth:`~xlrd.sheet.Sheet.row_slice` and :py:meth:`~xlrd.sheet.Sheet.col_slice` methods 
    return a list of :py:class:`~xlrd.sheet.Sheet.Cell` objects in a row 
    or column, respectively, bounded by and start index and an optional end index.
*   The :py:meth:`~xlrd.sheet.Sheet.row_types` 
    and :py:meth:`~xlrd.sheet.Sheet.col_types` methods 
    return a list of integers representing the cell types in a row 
    or column, respectively, bounded by and start index and an optional end index.
*   The :py:meth:`~xlrd.sheet.Sheet.row_values` and :py:meth:`~xlrd.sheet.Sheet.col_values` methods 
    return a list of objects representing the cell values in a the row or 
    column, bounded by a start index and an optional end index.

The following examples from sheet_iteration.py demonstrates all of the sheet iteration methods:

::

  from xlrd import open_workbook
  
  book = open_workbook('odd.xls')
  sheet0 = book.sheet_by_index(0)
  sheet1 = book.sheet_by_index(1)
  
  print sheet0.row(0)
  print sheet0.col(0)
  print
  print sheet0.row_slice(0,1)
  print sheet0.row_slice(0,1,2)
  print sheet0.row_values(0,1)
  print sheet0.row_values(0,1,2)
  print sheet0.row_types(0,1)
  print sheet0.row_types(0,1,2)
  print
  print sheet1.col_slice(0,1)
  print sheet0.col_slice(0,1,2)
  print sheet1.col_values(0,1)
  print sheet0.col_values(0,1,2)
  print sheet1.col_types(0,1)
  print sheet0.col_types(0,1,2)
  
  

Utility Functions
~~~~~~~~~~~~~~~~~

When navigating around a :py:meth:`~xlrd.book.Book`, it's often useful 
to be able to convert between row and column indexes and 
the Excel cell references that users may be used to seeing. 

The following functions are provided to help with this:

*   The :py:func:`~xlrd.formula.cellname` function turns a row and column index into 
    a relative Excel cell reference.
*   The :py:func:`~xlrd.formula.cellnameabs` function turns a row and column 
    index into an absolute Excel cell reference.
*   The :py:func:`~xlrd.formula.colname` function turns a column index into an Excel column name.

These three functions are demonstrated in the following example:

::

  from xlrd import cellname, cellnameabs, colname
  
  print cellname(0,0),cellname(10,10),cellname(100,100)
  print cellnameabs(3,1),cellnameabs(41,59),cellnameabs(265,358)
  print colname(0),colname(10),colname(100)
  utility.py

Unicode
-------

All text attributes and values used by :py:mod:`xlrd` will be either 
:func:`unicode` objects or, in rare cases, ascii strings.

Each piece of text in an Excel file written by Microsoft Excel is encoded into one of the following:

*   Latin1, if it fits
*   UTF_16_LE, if it doesn't fit into Latin1
*   In older files, by an encoding specified by an MS codepage. These are mapped 
    to Python encodings by :mod:`xlrd` and still result in :func:`unicode` objects.

In rare cases, other software has been known to write no codepage or the wrong codepage 
into Excel files. In this case, the correct **encoding** may need to 
be specified to :py:func:`xlrd.open_workbook`.

::

  from xlrd import open_workbook
  book = open_workbook('dodgy.xls', encoding='cp1252')


Types of Cell
-------------

We have already seen the cell type expressed as an integer. This integer corresponds to 
a set of constants in xlrd that identify the type of the cell. The full set of possible 
cell types is listed in the following sections.

Text
~~~~

*   These are represented by the :py:const:`xlrd.XL_CELL_TEXT` constant.
*   Cells of this type will have values that are :py:func:`unicode` objects.

Number
~~~~~~
*   These are represented by the :py:const:`xlrd.XL_CELL_NUMBER` constant.
*   Cells of this type will have values that are :func:`float` objects.

Date
~~~~
.. note::
    
    Dates don't really exist in Excel files, they are merely Numbers with a particular number formatting.

*   These are represented by the :py:const:`xlrd.XL_CELL_DATE` constant.
*   :py:mod:`xlrd` will return :py:const:`xlrd.XL_CELL_DATE` as the cell type 
    if the number format string looks like a date.
*   The :py:func:`xlrd.xldate_as_tuple` method is provided for 
    turning the :func:`float` in a Date cell into a :func:`tuple` suitable for 
    instantiating various :py:mod:`datetime` objects. 

This dates.py example shows how to use it:

::

  from datetime import date,datetime,time
  from xlrd import open_workbook,xldate_as_tuple
  
  book = open_workbook('types.xls')
  sheet = book.sheet_by_index(0)
  
  date_value = xldate_as_tuple(sheet.cell(3,2).value,book.datemode)
  print datetime(*date_value),date(*date_value[:3])
  datetime_value = xldate_as_tuple(sheet.cell(3,3).value,book.datemode)
  print datetime(*datetime_value)
  time_value = xldate_as_tuple(sheet.cell(3,4).value,book.datemode)
  print time(*time_value[3:])
  print datetime(*time_value)
  

Caveats:

*   Excel files have two possible date modes
    *   one for files originally created on Windows 
    *   and one for files originally created on an Apple machine. 
*   This is expressed as the :py:attr:`~xlrd.book.Book.datemode` 
    attribute of :py:class:`~xlrd.book.Book` objects 
    and **must** be passed to :func:`~xlrd.xldate.xldate_as_tuple`.

* The Excel file format has various problems with dates before 3 Jan 1904 that can cause date ambiguities that can result in :func:`~xlrd.xldate.xldate_as_tuple` raising an :class:`~xlrd.xldate.XLDateError`.

* The Excel formula function ``DATE()`` can return unexpected dates in certain circumstances.

Boolean
~~~~~~~

*   These are represented by the :py:const:`xlrd.XL_CELL_BOOLEAN` constant.
*   Cells of this type will have values that are :func:`bool` objects.

Error
~~~~~

*   These are represented by the :py:data:`xlrd.XL_CELL_ERROR` constant.
*   Cells of this type will have values that are integers representing specific error codes.
*   The :py:attr:`error_text_from_code` dictionary can be used to turn error codes into error messages:

::

  from xlrd import open_workbook, error_text_from_code
  
  book = open_workbook('types.xls')
  sheet = book.sheet_by_index(0)
  
  print error_text_from_code[sheet.cell(5,2).value]
  print error_text_from_code[sheet.cell(5,3).value]


For a simpler way of sensibly displaying all cell types, see :py:func:`xlutils.display`.

Empty / Blank
~~~~~~~~~~~~~

Excel files only store cells that either have information in them or have formatting 
applied to them. However, :py:mod:`xlrd` presents sheets as rectangular grids of cells.

Cells where no information is present in the Excel file are represented by 
the :py:data:`xlrd.XL_CELL_EMPTY` constant. In addition, there is only 
one “empty cell”, whose value is an empty string, used by :py:mod:`xlrd`, so empty cells 
may be checked using a Python identity check.

Cells where only formatting information is present in the Excel file are 
represented by the :py:attr:`xlrd.XL_CELL_BLANK` constant and their value will always be an empty string.

::

  from xlrd import open_workbook,empty_cell
  
  print empty_cell.value
  
  book = open_workbook('types.xls')
  sheet = book.sheet_by_index(0)
  empty = sheet.cell(6,2)
  blank = sheet.cell(7,2)
  print empty is blank, empty is empty_cell, blank is empty_cell
  
  book = open_workbook('types.xls',formatting_info=True)
  sheet = book.sheet_by_index(0)
  empty = sheet.cell(6,2)
  blank = sheet.cell(7,2)
  print empty.ctype,repr(empty.value)
  print blank.ctype,repr(blank.value)
  
  #emptyblank.py

The following example brings all of the above cell types together and shows examples of their use:

::

  from xlrd import open_workbook
  
  def cell_contents(sheet,row_x):
      result = []
      for col_x in range(2,sheet.ncols):
          cell = sheet.cell(row_x,col_x)
          result.append((cell.ctype,cell,cell.value))
      return result
  
  sheet = open_workbook('types.xls').sheet_by_index(0)
  
  print 'XL_CELL_TEXT',cell_contents(sheet,1)
  print 'XL_CELL_NUMBER',cell_contents(sheet,2)
  print 'XL_CELL_DATE',cell_contents(sheet,3)
  print 'XL_CELL_BOOLEAN',cell_contents(sheet,4)
  print 'XL_CELL_ERROR',cell_contents(sheet,5)
  print 'XL_CELL_BLANK',cell_contents(sheet,6)
  print 'XL_CELL_EMPTY',cell_contents(sheet,7)
  
  print
  sheet = open_workbook(
              'types.xls',formatting_info=True
              ).sheet_by_index(0)
  
  print 'XL_CELL_TEXT',cell_contents(sheet,1)
  print 'XL_CELL_NUMBER',cell_contents(sheet,2)
  print 'XL_CELL_DATE',cell_contents(sheet,3)
  print 'XL_CELL_BOOLEAN',cell_contents(sheet,4)
  print 'XL_CELL_ERROR',cell_contents(sheet,5)
  print 'XL_CELL_BLANK',cell_contents(sheet,6)
  print 'XL_CELL_EMPTY',cell_contents(sheet,7)
  
  cell_types.py

Names
-----

These are an infrequently used but powerful way of abstracting commonly used information found within Excel files.

They have many uses, and :py:mod:`xlrd` can extract information from many of them. A 
notable exception are names that refer to sheet and VBA macros, which are extracted but should be ignored.

Names are created in Excel by navigating to ``Insert > Name > Define``. If you plan 
to use :py:mod:`xlrd` to extract information from Names, familiarity with the definition 
and use of names in your chosen spreadsheet application is a good idea.

Types
~~~~~

A Name can refer to:

* A constant

  * ``CurrentInterestRate = 0.015``

  * ``NameOfPHB = “Attila T. Hun”``

* An absolute (i.e. not relative) cell reference

  * ``CurrentInterestRate = Sheet1!$B$4``

* Absolute reference to a 1D, 2D, or 3D block of cells

  * ``MonthlySalesByRegion = Sheet2:Sheet5!$A$2:$M$100``

* A list of absolute references

  * ``Print_Titles = [row_header_ref, col_header_ref])``

Constants can be extracted.

The coordinates of an absolute reference can be extracted so that you can then extract the corresponding data from the relevant sheet(s).

A relative reference is useful only if you have external knowledge of what cells can be used as the origin. Many formulas found in Excel files include function calls and multiple references and are not useful, and can be too hard to evaluate.

A full calculation engine is not included in :py:mod:`xlrd`.

Scope
~~~~~

The scope of a Name can be global, or it may be specific to a particular sheet. A Name's 
identifier may be re-used in different scopes. When there are multiple Names with 
the same identifier, the most appropriate one is used based on scope. A good example 
of this is the built-in name ``Print_Area``; each worksheet may have one of these.

Examples:

``name=rate, scope=Sheet1, formula=0.015``

``name=rate, scope=Sheet2, formula=0.023``

``name=rate, scope=``*global*``, formula=0.040``

A cell formula ``(1+rate)^20`` is equivalent to ``1.015^20`` if it appears in ``Sheet1`` but equivalent to ``1.023^20`` if it appears in ``Sheet2``, and ``1.040^20`` if it appears in any other sheet.

Usage
~~~~~

Common reasons for using names include:

* Assigning textual names to values that may occur in many places within a workbook

  * eg: ``RATE = 0.015``

* Assigning textual names to complex formulae that may be easily mis-copied

  * eg: ``SALES_RESULTS = $A$10:$M$999``

Here's an example real-world use case: reporting to head office. A company's head office makes up a template workbook. Each department gets a copy to fill in. The various ranges of data to be provided all have defined names. When the files come back, a script is used to
validate that the department hasn't trashed the workbook and the names are used to extract the data for further processing. Using names decouples any artistic repositioning of the ranges, by either head office template-designing user or by departmental users who are filling in the template, from the script which only has to know what the names of the ranges are.

In the examples directory of the ``xlrd`` distribution you will find ``namesdemo.xls`` which has examples of most of the non-macro varieties of defined names. There is also ``xlrdnamesAPIdemo.py`` which shows how to use the name lookup dictionaries, and how to extract constants and references and the data that references point to.

Formatting
----------

We've already seen that :py:func:`~xlrd.open_workbook` has a parameter to load formatting information 
from Excel files. When this is done, all the formatting information is available, 
but the details of how it is presented are beyond the scope of this tutorial.

If you wish to copy existing formatted data to a new Excel file, see 
:py:func:`xlutils.copy` and :py:func:`xlutils.filter`.

If you do wish to inspect formatting information, you'll need to consult the following attributes of the following classes:

* :py:class:`xlrd.book.Book`
    * :py:attr:`~xlrd.book.Book.colour_map`
    * :py:attr:`~xlrd.book.Book.font_list`
    * :py:attr:`~xlrd.book.Book.format_list`
    * :py:attr:`~xlrd.book.Book.format_map`
    * :py:attr:`~xlrd.book.Book.palette_record`
    * :py:attr:`~xlrd.book.Book.style_name_map`
    * :py:attr:`~xlrd.book.Book.xf_list`

* :py:class:`xlrd.sheet.Sheet`
    * :py:attr:`~xlrd.sheet.Sheet.cell_xf_index`
    * :py:attr:`~xlrd.sheet.Sheet.rowinfo_map`
    * :py:attr:`~xlrd.sheet.Sheet.colinfo_map`
    * :py:attr:`~xlrd.sheet.Sheet.computed_column_width`
    * :py:attr:`~xlrd.sheet.Sheet.default_additional_space_above`
    * :py:attr:`~xlrd.sheet.Sheet.default_additional_space_below`
    * :py:attr:`~xlrd.sheet.Sheet.default_row_height`
    * :py:attr:`~xlrd.sheet.Sheet.default_row_height_mismatch`
    * :py:attr:`~xlrd.sheet.Sheet.default_row_hidden`
    * :py:attr:`~xlrd.sheet.Sheet.defcolwidth`
    * :py:attr:`~xlrd.sheet.Sheet.gcw`
    * :py:attr:`~xlrd.sheet.Sheet.merged_cells`
    * :py:attr:`~xlrd.sheet.Sheet.standardwidth`

* :py:class:`xlrd.sheet.Cell`
    * :py:attr:`~xlrd.sheet.Cell.xf_index`

In addition, the following classes are used solely to represent formatting information:

* :py:class:`xlrd.sheet.Rowinfo`
* :py:class:`xlrd.sheet.Colinfo`
* :py:class:`xlrd.formatting.Font`
* :py:class:`xlrd.formatting.Format`
* :py:class:`xlrd.formatting.XF`
* :py:class:`xlrd.formatting.XFAlignment`
* :py:class:`xlrd.formatting.XFBackground`
* :py:class:`xlrd.formatting.XFBorder`
* :py:class:`xlrd.formatting.XFProtection`

Working with large Excel files
------------------------------

If you are working with particularly large Excel files, then there are two 
features of :py:mod:`xlrd` that you should be aware of:

*   The **on_demand** parameter can be passed as *True* to :py:func:`xlrd.open_workbook` resulting 
    in worksheets only being loaded into memory when they are requested.
*   :py:class:`~xlrd.book.Book` objects have an :py:meth:`~xlrd.book.Book.unload_sheet` method 
    that will unload a worksheet, specified by either sheet index or sheet name, from memory.

The following example shows how a large workbook could be iterated over when only sheets 
matching a certain pattern need to be inspected, and where only one of those sheets ends 
up in memory at any one time:

::

  from xlrd import open_workbook
  
  book = open_workbook('simple.xls',on_demand=True)
  
  for name in book.sheet_names():
      if name.endswith('2'):
          sheet = book.sheet_by_name(name)
          print sheet.cell_value(0,0)
          book.unload_sheet(name)
          

Introspecting Excel files with :command:`runxlrd.py`
-----------------------------------------------------

The :py:mod:`xlrd` source distribution includes a :command:`runxlrd.py` script that 
is extremely useful for introspecting Excel files without writing a single line of Python.

You are encouraged to run a variety of the commands it provides over the Excel files 
provided in the course materials.

The following gives an overview of what's available from ``runxlrd``, and can be obtained using ``python runxlrd.py –-help``:

::

  runxlrd.py [options] command [input-file-patterns]
  
  Commands:
  
  2rows           Print the contents of first and last row in each sheet
  3rows           Print the contents of first, second and last row in each sheet
  bench           Same as "show", but doesn't print -- for profiling
  biff_count[1]   Print a count of each type of BIFF record in the file
  biff_dump[1]    Print a dump (char and hex) of the BIFF records in the file
  fonts           hdr + print a dump of all font objects
  hdr             Mini-overview of file (no per-sheet information)
  hotshot         Do a hotshot profile run e.g. ... -f1 hotshot bench bigfile*.xls
  labels          Dump of sheet.col_label_ranges and ...row... for each sheet
  name_dump       Dump of each object in book.name_obj_list
  names           Print brief information for each NAME record
  ov              Overview of file
  profile         Like "hotshot", but uses cProfile
  show            Print the contents of all rows in each sheet
  version[0]      Print versions of xlrd and Python and exit
  xfc             Print "XF counts" and cell-type counts -- see code for details
  
  [0] means no file arg
  [1] means only one file arg i.e. no glob.glob pattern


  Options:
  
  -h, --help            show this help message and exit
  -l LOGFILENAME, --logfilename=LOGFILENAME
                        contains error messages
  -v VERBOSITY, --verbosity=VERBOSITY
                        level of information and diagnostics provided
  -p PICKLEABLE, --pickleable=PICKLEABLE
                        1: ensure Book object is pickleable (default); 0: don't bother
  -m MMAP, --mmap=MMAP  1: use mmap; 0: don't use mmap; -1: accept heuristic
  -e ENCODING, --encoding=ENCODING
                        encoding override
  -f FORMATTING, --formatting=FORMATTING
                        0 (default): no fmt info 1: fmt info (all cells)
  -g GC, --gc=GC        0: auto gc enabled; 1: auto gc disabled, manual collect after each file; 2: no gc
  -s ONESHEET, --onesheet=ONESHEET
                        restrict output to this sheet (name or index)
  -u, --unnumbered      omit line numbers or offsets in biff_dump

