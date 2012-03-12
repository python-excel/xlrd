.. xlrd documentation master file, created by B. Scott Michel (bscottm@ieee.org)
.. include:: <isonum.txt>
.. |OOodocs| replace:: OOo documents

.. toctree::
   :maxdepth: 2

################################################################################
xlrd: A Python module for reading Microsoft\ |trade| Excel\ |trade| spreadsheets
################################################################################

Version: |version| -- 2009-05-31

================
Acknowledgements
================

Development of this module would not have been possible without the
document "OpenOffice.org's Documentation of the Microsoft Excel File
Format" ("|OOodocs|" for short).  The latest version is available from
OpenOffice.org in `PDF format`_ and `ODT format`_.  Small portions of
the |OOodocs| are reproduced in this document. A study of these
documents is recommended for those who wish a deeper understanding of
the Excel\ |trade| file layout than the :mod:`xlrd` docs can provide.

.. _PDF format: http://sc.openoffice.org/excelfileformat.pdf
.. _ODT format: http://sc.openoffice.org/excelfileformat.odt

Backporting to Python 2.1 was partially funded by `Journyx`_ -
provider of timesheet and project accounting solutions.

.. _Journyx: http://journyx.com/

Provision of formatting information in version 0.6.1 was funded by
`Simplistix Ltd`_.

.. _Simplistix Ltd: http://www.simplistix.co.uk

========
Overview
========

:mod:`xlrd` is a Python module for **reading** Microsoft\ |trade|
Excel\ |trade| spreadsheet data. This module is not useful for
*writing* Excel\ |trade| spreadsheets.

====================
Implementation Notes
====================

.. _palette_and_colours:

---------------------------
The Palette; Colour Indices
---------------------------

A colour is represented in Excel as a (red, green, blue) ("RGB") tuple
with each component in range(256). However it is not possible to
access an unlimited number of colours; each spreadsheet is limited to
a palette of 64 different colours (24 in Excel 3.0 and 4.0, 8 in Excel
2.0). Colours are referenced by an index ("colour index") into this
palette. Colour indexes 0 to 7 represent 8 fixed built-in colours:
black, white, red, green, blue, yellow, magenta, and cyan.

The remaining colours in the palette (8 to 63 in Excel 5.0 and later)
can be changed by the user. In the Excel 2003 UI, Tools/Options/Color
presents a palette of 7 rows of 8 colours. The last two rows are
reserved for use in charts.

The correspondence between this grid and the assigned colour indexes
is NOT left-to-right top-to-bottom.

Indexes 8 to 15 correspond to changeable parallels of the 8 fixed
colours -- for example, index 7 is forever cyan; index 15 starts off
being cyan but can be changed by the user. The default colour for each
index depends on the file version; tables of the defaults are
available in the source code. If the user changes one or more colours,
a PALETTE record appears in the XLS file -- it gives the RGB values
for *all* changeable indexes.

Note that colours can be used in "number formats": "[CYAN]...." and
"[COLOR8]...." refer to colour index 7; "[COLOR16]...." will produce
cyan unless the user changes colour index 15 to something else.

In addition, there are several "magic" colour indexes used by Excel:

+----------------------------------------+--------------------------------------------------------+
| 0x18 (BIFF3-BIFF4), 0x40 (BIFF5-BIFF8) | System window text colour for border lines             |
|                                        | (used in XF, CF, and WINDOW2 records)                  |
+----------------------------------------+--------------------------------------------------------+
| 0x19 (BIFF3-BIFF4), 0x41 (BIFF5-BIFF8) | System window background colour for pattern background |
|                                        | (used in XF and CF records )                           |
+----------------------------------------+--------------------------------------------------------+
| 0x43                                   | System face colour (dialogue background colour)        |
+----------------------------------------+--------------------------------------------------------+
| 0x4D                                   | System window text colour for chart border lines       |
+----------------------------------------+--------------------------------------------------------+
| 0x4E                                   | System window background colour for chart areas        |
+----------------------------------------+--------------------------------------------------------+
| 0x4F                                   | Automatic colour for chart border lines (seems to be   |
|                                        | always Black)                                          |
+----------------------------------------+--------------------------------------------------------+
| 0x50                                   | System ToolTip background colour (used in note objects)|
+----------------------------------------+--------------------------------------------------------+
| 0x51                                   | System ToolTip text colour (used in note objects)      |
+----------------------------------------+--------------------------------------------------------+
| 0x7FFF                                 | System window text colour for fonts (used in FONT and  |
|                                        | CF records)                                            |
+----------------------------------------+--------------------------------------------------------+

Note 0x7FFF appears to be the *default* colour index. It appears quite
often in FONT records.

######################
The :mod:`xlrd` Module
######################

.. automodule:: xlrd
   :members:

===========================================================
:mod:`biffh`: Binary Interchange File Format (BIFF) Handler
===========================================================

.. automodule:: xlrd.biffh
   :members:

====================================================
:mod:`formatting`: Extended Formatting (XF) handling
====================================================

.. automodule:: xlrd.formatting
   :members:

##################
Indices and tables
##################

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`
