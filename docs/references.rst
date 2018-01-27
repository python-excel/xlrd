Named references, constants, formulas, and macros
=================================================

.. currentmodule:: xlrd.book


A name is used to refer to a cell, a group of cells, a constant
value, a formula, or a macro. Usually the scope of a name is global
across the whole workbook. However it can be local to a worksheet.
For example, if the sales figures are in different cells in
different sheets, the user may define the name "Sales" in each
sheet. There are built-in names, like "Print_Area" and
"Print_Titles"; these two are naturally local to a sheet.

To inspect the names with a user interface like MS Excel, OOo Calc,
or Gnumeric, click on Insert -> Names -> Define. This will show the global
names, plus those local to the currently selected sheet.

A :class:`Book` object provides two dictionaries (:attr:`Book.name_map` and
:attr:`Book.name_and_scope_map`) and a list (:attr:`Book.name_obj_list`) which
allow various ways of accessing the :class:`Name` objects.
There is one :class:`Name` object for each `NAME` record found in the workbook.
:class:`Name` objects have many attributes, several of which are relevant only
when ``obj.macro`` is ``1``.

In the examples directory you will find ``namesdemo.xls`` which
showcases the many different ways that names can be used, and
``xlrdnamesAPIdemo.py`` which offers 3 different queries for inspecting
the names in your files, and shows how to extract whatever a name is
referring to. There is currently one "convenience method",
:meth:`Name.cell`, which extracts the value in the case where the name
refers to a single cell. The source code for :meth:`Name.cell` is an extra
source of information on how the :class:`Name` attributes hang together.

.. note::

  Name information is *not* extracted from files older than
  Excel 5.0 (``Book.biff_version < 50``).
