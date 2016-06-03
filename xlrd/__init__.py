# Copyright (c) 2005-2012 Stephen John Machin, Lingfo Pty Ltd
# This module is part of the xlrd package, which is released under a
# BSD-style licence.
from .info import __VERSION__


# <h2>General information</h2>

# <h3>Formatting</h3>
#
# <h4>Introduction</h4>
#
# <p>This collection of features, new in xlrd version 0.6.1, is intended
# to provide the information needed to (1) display/render spreadsheet contents
# (say) on a screen or in a PDF file, and (2) copy spreadsheet data to another
# file without losing the ability to display/render it.</p>
#
# <h4>The Palette; Colour Indexes</h4>
#
# <p>A colour is represented in Excel as a (red, green, blue) ("RGB") tuple
# with each component in range(256). However it is not possible to access an
# unlimited number of colours; each spreadsheet is limited to a palette of 64 different
# colours (24 in Excel 3.0 and 4.0, 8 in Excel 2.0). Colours are referenced by an index
# ("colour index") into this palette.
#
# Colour indexes 0 to 7 represent 8 fixed built-in colours: black, white, red, green, blue,
# yellow, magenta, and cyan.<p>
#
# The remaining colours in the palette (8 to 63 in Excel 5.0 and later)
# can be changed by the user. In the Excel 2003 UI, Tools/Options/Color presents a palette
# of 7 rows of 8 colours. The last two rows are reserved for use in charts.<br />
# The correspondence between this grid and the assigned
# colour indexes is NOT left-to-right top-to-bottom.<br />
# Indexes 8 to 15 correspond to changeable
# parallels of the 8 fixed colours -- for example, index 7 is forever cyan;
# index 15 starts off being cyan but can be changed by the user.<br />
#
# The default colour for each index depends on the file version; tables of the defaults
# are available in the source code. If the user changes one or more colours,
# a PALETTE record appears in the XLS file -- it gives the RGB values for *all* changeable
# indexes.<br />
# Note that colours can be used in "number formats": "[CYAN]...." and "[COLOR8]...." refer
# to colour index 7; "[COLOR16]...." will produce cyan
# unless the user changes colour index 15 to something else.<br />
#
# <p>In addition, there are several "magic" colour indexes used by Excel:<br />
# 0x18 (BIFF3-BIFF4), 0x40 (BIFF5-BIFF8): System window text colour for border lines
# (used in XF, CF, and WINDOW2 records)<br />
# 0x19 (BIFF3-BIFF4), 0x41 (BIFF5-BIFF8): System window background colour for pattern background
# (used in XF and CF records )<br />
# 0x43: System face colour (dialogue background colour)<br />
# 0x4D: System window text colour for chart border lines<br />
# 0x4E: System window background colour for chart areas<br />
# 0x4F: Automatic colour for chart border lines (seems to be always Black)<br />
# 0x50: System ToolTip background colour (used in note objects)<br />
# 0x51: System ToolTip text colour (used in note objects)<br />
# 0x7FFF: System window text colour for fonts (used in FONT and CF records)<br />
# Note 0x7FFF appears to be the *default* colour index. It appears quite often in FONT
# records.<br />
#
# <h4>Default Formatting</h4>
#
# Default formatting is applied to all empty cells (those not described by a cell record).
# Firstly row default information (ROW record, Rowinfo class) is used if available.
# Failing that, column default information (COLINFO record, Colinfo class) is used if available.
# As a last resort the worksheet/workbook default cell format will be used; this
# should always be present in an Excel file,
# described by the XF record with the fixed index 15 (0-based). By default, it uses the
# worksheet/workbook default cell style, described by the very first XF record (index 0).
#
# <h4> Formatting features not included in xlrd version 0.6.1</h4>
# <ul>
#   <li>Rich text i.e. strings containing partial <b>bold</b> <i>italic</i>
#       and <u>underlined</u> text, change of font inside a string, etc.
#       See OOo docs s3.4 and s3.2.
#       <i> Rich text is included in version 0.7.2</i></li>
#   <li>Asian phonetic text (known as "ruby"), used for Japanese furigana. See OOo docs
#       s3.4.2 (p15)</li>
#   <li>Conditional formatting. See OOo docs
#       s5.12, s6.21 (CONDFMT record), s6.16 (CF record)</li>
#   <li>Miscellaneous sheet-level and book-level items e.g. printing layout, screen panes. </li>
#   <li>Modern Excel file versions don't keep most of the built-in
#       "number formats" in the file; Excel loads formats according to the
#       user's locale. Currently xlrd's emulation of this is limited to
#       a hard-wired table that applies to the US English locale. This may mean
#       that currency symbols, date order, thousands separator, decimals separator, etc
#       are inappropriate. Note that this does not affect users who are copying XLS
#       files, only those who are visually rendering cells.</li>
# </ul>
#
# <h3>Loading worksheets on demand</h3>
#
# <p>This feature, new in version 0.7.1, is governed by the on_demand argument
# to the open_workbook() function and allows saving memory and time by loading
# only those sheets that the caller is interested in, and releasing sheets
# when no longer required.</p>
#
# <p>on_demand=False (default): No change. open_workbook() loads global data
# and all sheets, releases resources no longer required (principally the
# str or mmap object containing the Workbook stream), and returns.</p>
#
# <p>on_demand=True and BIFF version < 5.0: A warning message is emitted,
# on_demand is recorded as False, and the old process is followed.</p>
#
# <p>on_demand=True and BIFF version >= 5.0: open_workbook() loads global
# data and returns without releasing resources. At this stage, the only
# information available about sheets is Book.nsheets and Book.sheet_names().</p>
#
# <p>Book.sheet_by_name() and Book.sheet_by_index() will load the requested
# sheet if it is not already loaded.</p>
#
# <p>Book.sheets() will load all/any unloaded sheets.</p>
#
# <p>The caller may save memory by calling
# Book.unload_sheet(sheet_name_or_index) when finished with the sheet.
# This applies irrespective of the state of on_demand.</p>
#
# <p>The caller may re-load an unloaded sheet by calling Book.sheet_by_xxxx()
#  -- except if those required resources have been released (which will
# have happened automatically when on_demand is false). This is the only
# case where an exception will be raised.</p>
#
# <p>The caller may query the state of a sheet:
# Book.sheet_loaded(sheet_name_or_index) -> a bool</p>
#
# <p> Book.release_resources() may used to save memory and close
# any memory-mapped file before proceding to examine already-loaded
# sheets. Once resources are released, no further sheets can be loaded.</p>
#
# <p> When using on-demand, it is advisable to ensure that
# Book.release_resources() is always called even if an exception
# is raised in your own code; otherwise if the input file has been
# memory-mapped, the mmap.mmap object will not be closed and you will
# not be able to access the physical file until your Python process
# terminates. This can be done by calling Book.release_resources()
# explicitly in the finally suite of a try/finally block.
# New in xlrd 0.7.2: the Book object is a "context manager", so if
# using Python 2.5 or later, you can wrap your code in a "with"
# statement.</p>
##

import sys, zipfile, pprint
from . import timemachine
from .biffh import (
    XLRDError,
    biff_text_from_num,
    error_text_from_code,
    XL_CELL_BLANK,
    XL_CELL_TEXT,
    XL_CELL_BOOLEAN,
    XL_CELL_ERROR,
    XL_CELL_EMPTY,
    XL_CELL_DATE,
    XL_CELL_NUMBER
    )
from .formula import * # is constrained by __all__
from .book import Book, colname #### TODO #### formula also has `colname` (restricted to 256 cols)
from .sheet import empty_cell
from .xldate import XLDateError, xldate_as_tuple, xldate_as_datetime
from .xlsx import X12Book

if sys.version.startswith("IronPython"):
    # print >> sys.stderr, "...importing encodings"
    import encodings

try:
    import mmap
    MMAP_AVAILABLE = 1
except ImportError:
    MMAP_AVAILABLE = 0
USE_MMAP = MMAP_AVAILABLE

##
#
# Open a spreadsheet file for data extraction.
#
# @param filename The path to the spreadsheet file to be opened.
#
# @param logfile An open file to which messages and diagnostics are written.
#
# @param verbosity Increases the volume of trace material written to the logfile.
#
# @param use_mmap Whether to use the mmap module is determined heuristically.
# Use this arg to override the result. Current heuristic: mmap is used if it exists.
#
# @param file_contents ... as a string or an mmap.mmap object or some other behave-alike object.
# If file_contents is supplied, filename will not be used, except (possibly) in messages.
#
# @param encoding_override Used to overcome missing or bad codepage information
# in older-version files. Refer to discussion in the <b>Unicode</b> section above.
# <br /> -- New in version 0.6.0
#
# @param formatting_info Governs provision of a reference to an XF (eXtended Format) object
# for each cell in the worksheet.
# <br /> Default is <i>False</i>. This is backwards compatible and saves memory.
# "Blank" cells (those with their own formatting information but no data) are treated as empty
# (by ignoring the file's BLANK and MULBLANK records).
# It cuts off any bottom "margin" of rows of empty (and blank) cells and
# any right "margin" of columns of empty (and blank) cells.
# Only cell_value and cell_type are available.
# <br /> <i>True</i> provides all cells, including empty and blank cells.
# XF information is available for each cell.
# <br /> -- New in version 0.6.1
#
# @param on_demand Governs whether sheets are all loaded initially or when demanded
# by the caller. Please refer back to the section "Loading worksheets on demand" for details.
# <br /> -- New in version 0.7.1
#
# @param ragged_rows False (the default) means all rows are padded out with empty cells so that all
# rows have the same size (Sheet.ncols). True means that there are no empty cells at the ends of rows.
# This can result in substantial memory savings if rows are of widely varying sizes. See also the
# Sheet.row_len() method.
# <br /> -- New in version 0.7.2
#
# @return An instance of the Book class.

def open_workbook(filename=None,
    logfile=sys.stdout,
    verbosity=0,
    use_mmap=USE_MMAP,
    file_contents=None,
    encoding_override=None,
    formatting_info=False,
    on_demand=False,
    ragged_rows=False,
    ):
    peeksz = 4
    if file_contents:
        peek = file_contents[:peeksz]
    else:
        with open(filename, "rb") as f:
            peek = f.read(peeksz)
    if peek == b"PK\x03\x04": # a ZIP file
        if file_contents:
            zf = zipfile.ZipFile(timemachine.BYTES_IO(file_contents))
        else:
            zf = zipfile.ZipFile(filename)

        # Workaround for some third party files that use forward slashes and
        # lower case names. We map the expected name in lowercase to the
        # actual filename in the zip container.
        component_names = dict([(X12Book.convert_filename(name), name)
                                for name in zf.namelist()])

        if verbosity:
            logfile.write('ZIP component_names:\n')
            pprint.pprint(component_names, logfile)
        if 'xl/workbook.xml' in component_names:
            from . import xlsx
            bk = xlsx.open_workbook_2007_xml(
                zf,
                component_names,
                logfile=logfile,
                verbosity=verbosity,
                use_mmap=use_mmap,
                formatting_info=formatting_info,
                on_demand=on_demand,
                ragged_rows=ragged_rows,
                )
            return bk
        if 'xl/workbook.bin' in component_names:
            raise XLRDError('Excel 2007 xlsb file; not supported')
        if 'content.xml' in component_names:
            raise XLRDError('Openoffice.org ODS file; not supported')
        raise XLRDError('ZIP file contents not a known type of workbook')

    from . import book
    bk = book.open_workbook_xls(
        filename=filename,
        logfile=logfile,
        verbosity=verbosity,
        use_mmap=use_mmap,
        file_contents=file_contents,
        encoding_override=encoding_override,
        formatting_info=formatting_info,
        on_demand=on_demand,
        ragged_rows=ragged_rows,
        )
    return bk

##
# For debugging: dump an XLS file's BIFF records in char & hex.
# @param filename The path to the file to be dumped.
# @param outfile An open file, to which the dump is written.
# @param unnumbered If true, omit offsets (for meaningful diffs).

def dump(filename, outfile=sys.stdout, unnumbered=False):
    from .biffh import biff_dump
    bk = Book()
    bk.biff2_8_load(filename=filename, logfile=outfile, )
    biff_dump(bk.mem, bk.base, bk.stream_len, 0, outfile, unnumbered)

##
# For debugging and analysis: summarise the file's BIFF records.
# I.e. produce a sorted file of (record_name, count).
# @param filename The path to the file to be summarised.
# @param outfile An open file, to which the summary is written.

def count_records(filename, outfile=sys.stdout):
    from .biffh import biff_count_records
    bk = Book()
    bk.biff2_8_load(filename=filename, logfile=outfile, )
    biff_count_records(bk.mem, bk.base, bk.stream_len, outfile)
