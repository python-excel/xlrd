Changes
=======

.. currentmodule:: xlrd

1.2.0 (15 December 2018)
------------------------

- Added support for Python 3.7.
- Added optional support for defusedxml to help mitigate exploits.
- Automatically convert ``~`` in file paths to the current user's home
  directory.
- Removed ``examples`` directory from the installed package. They are still
  available in the source distribution.
- Fixed ``time.clock()`` deprecation warning.

1.1.0 (22 August 2017)
----------------------

- Fix for parsing of merged cells containing a single cell reference in xlsx
  files.

- Fix for "invalid literal for int() with base 10: 'true'" when reading some
  xlsx files.

- Make xldate_as_datetime available to import direct from xlrd.

- Build universal wheels.

- Sphinx documentation.

- Document the problem with XML vulnerabilities in xlsx files and mitigation
  measures.

- Fix :class:`NameError` on ``has_defaults is not defined``.

- Some whitespace and code style tweaks.

- Make example in README compatible with both Python 2 and 3.

- Add default value for cells containing errors that causeed parsing of some
  xlsx files to fail.

- Add Python 3.6 to the list of supported Python versions, drop 3.3 and 2.6.

- Use generator expressions to avoid unnecessary lists in memory.

- Document unicode encoding used in Excel files from Excel 97 onwards.

- Report hyperlink errors in R1C1 syntax.

Thanks to the following for their contributions to this release:

- icereval@gmail.com
- Daniel Rech
- Ville Skyttä
- Yegor Yefremov
- Maxime Lorant
- Alexandr N Zamaraev
- Zhaorong Ma
- Jon Dufresne
- Chris McIntyre
- coltleese@gmail.com
- Ivan Masá

1.0.0 (2 June 2016)
-------------------

- Official support, such as it is, is now for 2.6, 2.7, 3.3+

- Fixes a bug in looking up non-lowercase sheet filenames by ensuring that the
  sheet targets are transformed the same way as the component_names dict keys.

- Fixes a bug for ``ragged_rows=False`` when merged cells increases the number
  of columns in the sheet. This requires all rows to be extended to ensure equal
  row lengths that match the number of columns in the sheet.

- Fixes to enable reading of SAP-generated .xls files.

- support BIFF4 files with missing FORMAT records.

- support files with missing WINDOW2 record.

- Empty cells are now always unicode strings, they were a bytestring on
  Python 2 and a unicode string on Python 3.

- Fix for ``<cell>`` ``inlineStr`` attribute without ``<si>`` child.

- Fix for a zoom of ``None`` causing problems on Python 3.

- Fix parsing of bad dimensions.

- Fix xlsx sheet to comments relationship.

Thanks to the following for their contributions to this release:

- Lars-Erik Hannelius
- Deshi Xiao
- Stratos Moro
- Volker Diels-Grabsch
- John McNamara
- Ville Skyttä
- Patrick Fuller
- Dragon Dave McKee
- Gunnlaugur Þór Briem

0.9.4 (14 July 2015)
--------------------

- Automated tests are now run on Python 3.4

- Use ``ElementTree.iter()`` if available, instead of the deprecated
  ``getiterator()`` when parsing xlsx files.

- Fix #106 : Exception Value: unorderable types: Name() < Name()

- Create row generator expression with Sheet.get_rows()

- Fix for forward slash file separator and lowercase names within xlsx
  internals.

Thanks to the following for their contributions to this release:

- Corey Farwell
- Jonathan Kamens
- Deepak N
- Brandon R. Stoner
- John McNamara

0.9.3 (8 Apr 2014)
------------------

- Github issue #49

- Github issue #64 - skip meaningless chunk of 4 zero bytes between two
  otherwise-valid BIFF records

- Github issue #61 - fix updating of escapement attribute of Font objects read
  from workbooks.

- Implemented ``Sheet.visibility`` for xlsx files

- Ignore anchors (``$``) in cell references

- Dropped support for Python 2.5 and earlier, Python 2.6 is now the earliest
  Python release supported

- Read xlsx merged cell elements.

- Read cell comments in .xlsx files.

- Added xldate_as_datetime() function to convert from Excel
  serial date/time to datetime.datetime object.

Thanks to the following for their contributions to this release:

- John Machin
- Caleb Epstein
- Martin Panter
- John McNamara
- Gunnlaugur Þór Briem
- Stephen Lewis


0.9.2 (9 Apr 2013)
------------------

- Fix some packaging issues that meant docs and examples were missing from the tarball.

- Fixed a small but serious regression that caused problems opening .xlsx files.

0.9.1 (5 Apr 2013)
------------------

- Many fixes bugs in Python 3 support.
- Fix bug where ragged rows needed fixing when formatting info was being parsed.
- Improved handling of aberrant Excel 4.0 Worksheet files.
- Various bug fixes.
- Simplify a lot of the distribution packaging.
- Remove unused and duplicate imports.

Thanks to the following for their contributions to this release:

- Thomas Kluyver

0.9.0 (31 Jan 2013)
-------------------

- Support for Python 3.2+
- Many new unit test added.
- Continuous integration tests are now run.
- Various bug fixes.

Special thanks to Thomas Kluyver and Martin Panter for their work on
Python 3 compatibility.

Thanks to Manfred Moitzi for re-licensing his unit tests so we could include
them.

Thanks to the following for their contributions to this release:

- "holm"
- Victor Safronovich
- Ross Jones

0.8.0 (22 Aug 2012)
-------------------

- More work-arounds for broken source files.
- Support for reading .xlsx files.
- Drop support for Python 2.5 and older.

0.7.8 (7 June 2012)
-------------------

- Ignore superfluous zero bytes at end of xls OBJECT record.
- Fix assertion error when reading file with xlwt-written bitmap.

0.7.7 (13 Apr 2012)
-------------------

- More packaging changes, this time to support 2to3.

0.7.6 (3 Apr 2012)
------------------

- Fix more packaging issues.

0.7.5 (3 Apr 2012)
------------------
- Fix packaging issue that missed ``version.txt`` from the distributions.

0.7.4 (2 Apr 2012)
------------------

- More tolerance of out-of-spec files.
- Fix bugs reading long text formula results.

0.7.3 (28 Feb 2012)
-------------------

- Packaging and documentation updates.

0.7.2 (21 Feb 2012)
-------------------

- Tolerant handling of files with extra zero bytes at end of NUMBER record.
  Sample provided by Jan Kraus.
- Added access to cell notes/comments. Many cross-references added to Sheet
  class docs.
- Added code to extract hyperlink (HLINK) records. Based on a patch supplied by
  John Morrisey.
- Extraction of rich text formatting info based on code supplied by
  Nathan van Gheem.
- added handling of BIFF2 WINDOW2 record.
- Included modified version of page breaks patch from Sam Listopad.
- Added reading of the PANE record.
- Reading SCL record. New attribute ``Sheet.scl_mag_factor``.
- Lots of bug fixes.
- Added ``ragged_rows`` functionality.

0.7.1 (31 May 2009)
-------------------

- Backed out "slash'n'burn" of sheet resources in unload_sheet().
  Fixed problem with STYLE records on some Mac Excel files.
- quieten warnings
- Integrated on_demand patch by Armando Serrano Lombillo

0.7.0 (11 March 2009)
---------------------

+ colname utility function now supports more than 256 columns.
+ Fix bug where BIFF record type 0x806 was being regarded as a formula
  opcode.
+ Ignore PALETTE record when formatting_info is false.
+ Tolerate up to 4 bytes trailing junk on PALETTE record.
+ Fixed bug in unused utility function xldate_from_date_tuple which
  affected some years after 2099.
+ Added code for inspecting as-yet-unused record types: FILEPASS, TXO,
  NOTE.
+ Added inspection code for add_in function calls.
+ Added support for unnumbered biff_dump (better for doing diffs).
+ ignore distutils cruft
+ Avoid assertion error in compdoc when -1 used instead of -2 for
  first_SID of empty SCSS
+ Make version numbers match up.
+ Enhanced recovery from out-of-order/missing/wrong CODEPAGE record.
+ Added Name.area2d convenience method.
+ Avoided some checking of XF info when formatting_info is false.
+ Minor changes in preparation for XLSX support.
+ remove duplicate files that were out of date.
+ Basic support for Excel 2.0
+ Decouple Book init & load.
+ runxlrd: minor fix for xfc.
+ More Excel 2.x work.
+ is_date_format() tweak.
+ Better detection of IronPython.
+ Better error message (including first 8 bytes of file) when file is
  not in a supported format.
+ More BIFF2 formatting: ROW, COLWIDTH, and COLUMNDEFAULT records;
+ finished stage 1 of XF records.
+ More work on supporting BIFF2 (Excel 2.x) files.
+ Added support for Excel 2.x (BIFF2) files. Data only, no formatting
  info. Alpha.
+ Wasn't coping with EXTERNSHEET record followed by CONTINUE
  record(s).
+ Allow for BIFF2/3-style FORMAT record in BIFF4/8 file
+ Avoid crash when zero-length Unicode string missing options byte.
+ Warning message if sector sizes are extremely large.
+ Work around corrupt STYLE record
+ Added missing entry for blank cell type to ctype_text
+ Added "fonts" command to runxlrd script
+ Warning: style XF whose parent XF index != 0xFFF
+ Logfile arg wasn't being passed from open_workbook to
  compdoc.CompDoc.


0.6.1  (10 June 2007)
---------------------

+ Version number updated to 0.6.1
+ Documented runxlrd.py commands in its usage message. Changed
  commands: dump to biff_dump, count_records to biff_count.


0.6.1a5
-------

+ Bug fixed: Missing "<" in a struct.unpack call means can't open
  files on bigendian platforms. Discovered by "Mihalis".
+ Removed antique undocumented Book.get_name_dict method and
  experimental "trimming" facility.
+ Meaningful exception instead of IndexError if a SAT (sector
  allocation table) is corrupted.
+ If no CODEPAGE record in pre-8.0 file, assume ascii and keep going
  (instead of raising exception).


0.6.1a4
-------

+ At least one source of XLS files writes parent style XF records
  *after* the child cell XF records that refer to them, triggering
  IndexError in 0.5.2 and AssertionError in later versions. Reported
  with sample file by Todd O'Bryan. Fixed by changing to two-pass
  processing of XF records.
+ Formatting info in pre-BIFF8 files: Ensured appropriate defaults and
  lossless conversions to make the info BIFF8-compatible. Fixed bug in
  extracting the "used" flags.
+ Fixed problems discovered with opening test files from Planmaker
  2006 (http://www.softmaker.com/english/ofwcomp_en.htm): (1) Four files
  have reduced size of PALETTE record (51 and 32 colours; Excel writes
  56 always). xlrd now emits a NOTE to the logfile and continues. (2)
  FORMULA records use the Excel 2.x record code 0x0021 instead of
  0x0221. xlrd now continues silently. (3) In two files, at the OLE2
  compound document level, the internal directory says that the length
  of the Short-Stream Container Stream is 16384 bytes, but the actual
  contents are 11264 and 9728 bytes respectively. xlrd now emits a
  WARNING to the logfile and continues.
+ After discussion with Daniel Rentz, the concept of two lists of XF
  (eXtended Format) objects (raw_xf_list and computed_xf_list) has been
  abandoned. There is now a single list, called xf_list


0.6.1a3
-------

+ Added Book.sheets ... for sheetx, sheet in enumerate(book.sheets):
+ Formatting info: extraction of sheet-level flags from WINDOW2
  record, and sheet.visibility from BOUNDSHEET record. Added Macintosh-
  only Font attributes "outline" and "shadow'.


0.6.1a2
-------

+ Added extraction of merged cells info.
+ pyExcelerator uses "general" instead of "General" for the generic
  "number format". Worked around.
+ Crystal Reports writes "WORKBOOK" in the OLE2 Compound Document
  directory instead of "Workbook". Changed to case-insensitive directory
  search. Reported by Vic Simkus.


0.6.1a1 (18 Dec 2006)
---------------------

+ Added formatting information for cells (font, "number format",
  background, border, alignment and protection) and rows/columns
  (height/width etc). To save memory and time for those who don't need
  it, this information is extracted only if formatting_info=1 is
  supplied to the open_workbook() function. The cell records BLANK and
  MULBLANKS which contain no data, only formatting information, will
  continue to be ignored in the default (no formatting info) case.
+ Ralph Heimburger reported a problem with xlrd being intolerant about
  an Excel 4.0 file (created by "some web app") with a DIMENSIONS record
  that omitted Microsoft's usual padding with 2 unused bytes. Fixed.


0.6.0a4 (not released)
----------------------

+ Added extraction of human-readable formulas from NAME records.
+ Worked around OOo Calc writing 9-byte BOOLERR records instead of 8.
  Reported by Rory Campbell-Lange.
+ This history file converted to descending chronological order and
  HTML format.


0.6.0a3 (19 Sept 2006)
----------------------

+ Names: minor bugfixes; added script xlrdnameAPIdemo.py
+ ROW records were being used as additional hints for sizing memory
  requirements. In some files the ROW records overstate the number of
  used columns, and/or there are ROW records for rows that have no data
  in them. This would cause xlrd to report sheet.ncols and/or
  sheet.nrows as larger than reasonably expected. Change: ROW records
  are ignored. The number of columns/rows is based solely on the highest
  column/row index seen in non-empty data records. Empty data records
  (types BLANK and MULBLANKS) which contain no data, only formatting
  information, have always been ignored, and this will continue.
  Consequence: trailing rows and columns which contain only empty cells
  will vanish.


0.6.0a2 (13 Sept 2006)
----------------------


+ Fixed a bug reported by Rory Campbell-Lange.: "open failed";
  incorrect assumptions about the layout of array formulas which return
  strings.
+ Further work on defined names, especially the API.


0.6.0a1 (8 Sept 2006)
---------------------

+ Sheet objects have two new convenience methods: col_values(colx,
  start_rowx=0, end_rowx=None) and the corresponding col_types.
  Suggested by Dennis O'Brien.
+ BIFF 8 file missing its CODEPAGE record: xlrd will now assume
  utf_16_le encoding (the only possibility) and keep going.
+ Older files missing a CODEPAGE record: an exception will be raised.
  Thanks to Sergey Krushinsky for a sample file. The open_workbook()
  function has a new argument (encoding_override) which can be used if
  the CODEPAGE record is missing or incorrect (for example,
  codepage=1251 but the data is actually encoded in koi8_r). The
  runxlrd.py script takes a corresponding -e argument, for example -e
  cp1251
+ Further work done on parsing "number formats". Thanks to Chris
  Withers for the ``"General_)"`` example.
+ Excel 97 introduced the concept of row and column labels, defined by
  Insert > Name > Labels. The ranges containing the labels are now
  exposed as the Sheet attributes row_label_ranges and col_label_ranges.
+ The major effort in this 0.6.0 release has been the provision of
  access to named cell ranges and named constants (Excel:
  Insert/Name/Define). Juan C. Mendez provided very useful real-world
  sample files.


0.5.3a1 (24 May 2006)
---------------------

+ John Popplewell and Richard Sharp provided sample files which caused
  any reliance at all on DIMENSIONS records and ROW records to be
  abandoned.
+ If the file size is not a whole number of OLE sectors, a warning
  message is logged. Previously this caused an exception to be raised.


0.5.2 (14 March 2006)
---------------------

+ public release
+ Updated version numbers, README, HISTORY.


0.5.2a3 (13 March 2006)
-----------------------

+ Gnumeric writes user-defined formats with format codes starting at
  50 instead of 164; worked around.
+ Thanks to Didrik Pinte for reporting the need for xlrd to be more
  tolerant of the idiosyncracies of other software, for supplying sample
  files, and for performing alpha testing.
+ '_' character in a format should be treated like an escape
  character; fixed.
+ An "empty" formula result means a zero-length string, not an empty
  cell! Fixed.


0.5.2a2 (9 March 2006)
----------------------

+ Found that Gnumeric writes all DIMENSIONS records with nrows and
  ncols each 1 less than they should be (except when it clamps ncols at
  256!), and pyXLwriter doesn't write ROW records. Cell memory pre-
  allocation was generalised to use ROW records if available with fall-
  back to DIMENSIONS records.


0.5.2a1 (6 March 2006)
----------------------


+ pyXLwriter writes DIMENSIONS record with antique opcode 0x0000
  instead of 0x0200; worked around
+ A file written by Gnumeric had zeroes in DIMENSIONS record but data
  in cell A1; worked around


0.5.1 (18 Feb 2006)
--------------------

+ released to Journyx
+ Python 2.1 mmap requires file to be opened for update access. Added
  fall-back to read-only access without mmap if 2.1 open fails because
  "permission denied".


0.5 (7 Feb 2006)
----------------

+ released to Journyx
+ Now works with Python 2.1. Backporting to Python 2.1 was partially
  funded by Journyx - provider of timesheet and project accounting
  solutions (http://journyx.com/)
+ open_workbook() can be given the contents of a file instead of its
  name. Thanks to Remco Boerma for the suggestion.
+ New module attribute __VERSION__ (as a string; for example "0.5")
+ Minor enhancements to classification of formats as date or not-date.
+ Added warnings about files with inconsistent OLE compound document
  structures. Thanks to Roman V. Kiseliov (author of pyExcelerator) for
  the tip-off.


0.4a1, (7 Sept 2005)
--------------------

+ released to Laurent T.
+ Book and sheet objects can now be pickled and unpickled. Instead of
  reading a large spreadsheet multiple times, consider pickling it once
  and loading the saved pickle; can be much faster. Thanks to Laurent
  Thioudellet for the enhancement request.
+ Using the mmap module can be turned off. But you would only do that
  for benchmarking purposes.
+ Handling NUMBER records has been made faster


0.3a1 (15 May 2005)
-------------------

- first public release
