Formatting information in Excel Spreadsheets
============================================

Introduction
------------

This collection of features, new in xlrd version 0.6.1, is intended
to provide the information needed to:

- display/render spreadsheet contents (say) on a screen or in a PDF file
- copy spreadsheet data to another file without losing the ability to
  display/render it.

.. _palette:

The Palette; Colour Indexes
---------------------------

A colour is represented in Excel as a ``(red, green, blue)`` ("RGB") tuple
with each component in ``range(256)``. However it is not possible to access an
unlimited number of colours; each spreadsheet is limited to a palette of 64
different colours (24 in Excel 3.0 and 4.0, 8 in Excel 2.0).
Colours are referenced by an index ("colour index") into this palette.

Colour indexes 0 to 7 represent 8 fixed built-in colours:
black, white, red, green, blue, yellow, magenta, and cyan.

The remaining colours in the palette (8 to 63 in Excel 5.0 and later)
can be changed by the user. In the Excel 2003 UI,
Tools -> Options -> Color presents a palette
of 7 rows of 8 colours. The last two rows are reserved for use in charts.

The correspondence between this grid and the assigned
colour indexes is NOT left-to-right top-to-bottom.

Indexes 8 to 15 correspond to changeable
parallels of the 8 fixed colours -- for example, index 7 is forever cyan;
index 15 starts off being cyan but can be changed by the user.

The default colour for each index depends on the file version; tables of the
defaults are available in the source code. If the user changes one or more
colours, a ``PALETTE`` record appears in the XLS file -- it gives the RGB values
for *all* changeable
indexes.

Note that colours can be used in "number formats": ``[CYAN]....`` and
``[COLOR8]....`` refer to colour index 7; ``[COLOR16]....`` will produce cyan
unless the user changes colour index 15 to something else.

In addition, there are several "magic" colour indexes used by Excel:

``0x18`` (BIFF3-BIFF4), ``0x40`` (BIFF5-BIFF8):
  System window text colour for border lines (used in ``XF``, ``CF``, and
  ``WINDOW2`` records)

``0x19`` (BIFF3-BIFF4), ``0x41`` (BIFF5-BIFF8):
  System window background colour for pattern background (used in ``XF`` and
  ``CF`` records )

``0x43``:
  System face colour (dialogue background colour)

``0x4D``:
  System window text colour for chart border lines

``0x4E``:
  System window background colour for chart areas

``0x4F``:
  Automatic colour for chart border lines (seems to be always Black)

``0x50``:
  System ToolTip background colour (used in note objects)

``0x51``:
  System ToolTip text colour (used in note objects)

``0x7FFF``:
  System window text colour for fonts (used in ``FONT`` and ``CF`` records).

  .. note::
    ``0x7FFF`` appears to be the *default* colour index.
    It appears quite often in ``FONT`` records.

Default Formatting
------------------

Default formatting is applied to all empty cells (those not described by a cell
record):

- Firstly, row default information (``ROW`` record, :class:`~xlrd.sheet.Rowinfo`
  class) is used if available.

- Failing that, column default information (``COLINFO`` record,
  :class:`~xlrd.sheet.Colinfo` class) is used if available.

- As a last resort the worksheet/workbook default cell format will be used; this
  should always be present in an Excel file,
  described by the ``XF`` record with the fixed index 15 (0-based).
  By default, it uses the worksheet/workbook default cell style,
  described by the very first ``XF`` record (index 0).

Formatting features not included in xlrd
----------------------------------------

- Asian phonetic text (known as "ruby"), used for Japanese furigana.
  See OOo docs s3.4.2 (p15)

- Conditional formatting. See OOo docs s5.12, s6.21 (CONDFMT record), s6.16
  (CF record)

- Miscellaneous sheet-level and book-level items, e.g. printing layout,
  screen panes.

- Modern Excel file versions don't keep most of the built-in
  "number formats" in the file; Excel loads formats according to the
  user's locale. Currently, xlrd's emulation of this is limited to
  a hard-wired table that applies to the US English locale. This may mean
  that currency symbols, date order, thousands separator, decimals separator,
  etc are inappropriate.

  .. note::
    This does not affect users who are copying XLS
    files, only those who are visually rendering cells.
