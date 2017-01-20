Dates in Excel spreadsheets
===========================

.. currentmodule:: xlrd.xldate

In reality, there are no such things. What you have are floating point
numbers and pious hope.
There are several problems with Excel dates:

1. Dates are not stored as a separate data type; they are stored as
   floating point numbers and you have to rely on:

   - the "number format" applied to them in Excel and/or
   - knowing which cells are supposed to have dates in them.

   This module helps with the former by inspecting the
   format that has been applied to each number cell;
   if it appears to be a date format, the cell
   is classified as a date rather than a number.

   Feedback on this feature, especially from non-English-speaking locales,
   would be appreciated.

2. Excel for Windows stores dates by default as the number of
   days (or fraction thereof) since ``1899-12-31T00:00:00``. Excel for
   Macintosh uses a default start date of ``1904-01-01T00:00:00``.

   The date system can be changed in Excel on a per-workbook basis (for example:
   Tools -> Options -> Calculation, tick the "1904 date system" box).
   This is of course a bad idea if there are already dates in the
   workbook. There is no good reason to change it even if there are no
   dates in the workbook.

   Which date system is in use is recorded in the
   workbook. A workbook transported from Windows to Macintosh (or vice
   versa) will work correctly with the host Excel.

   When using this package's :func:`xldate_as_tuple` function to convert numbers
   from a workbook, you must use the :attr:`~xlrd.Book.datemode` attribute of
   the :class:`~xlrd.Book` object. If you guess, or make a judgement depending
   on where you believe the workbook was created, you run the risk of being 1462
   days out of kilter.

   Reference:
   https://support.microsoft.com/en-us/help/180162/xl-the-1900-date-system-vs.-the-1904-date-system


3. The Excel implementation of the Windows-default 1900-based date system
   works on the incorrect premise that 1900 was a leap year. It interprets the
   number 60 as meaning ``1900-02-29``, which is not a valid date.

   Consequently, any number less than 61 is ambiguous. For example, is 59 the
   result of ``1900-02-28`` entered directly, or is it ``1900-03-01`` minus 2
   days?

   The OpenOffice.org Calc program "corrects" the Microsoft problem;
   entering ``1900-02-27`` causes the number 59 to be stored.
   Save as an XLS file, then open the file with Excel and you'll see
   ``1900-02-28`` displayed.

   Reference: https://support.microsoft.com/en-us/help/214326/excel-incorrectly-assumes-that-the-year-1900-is-a-leap-year

4. The Macintosh-default 1904-based date system counts ``1904-01-02`` as day 1
   and ``1904-01-01`` as day zero. Thus any number such that
   ``(0.0 <= number < 1.0)`` is ambiguous. Is 0.625 a time of day
   (``15:00:00``), independent of the calendar, or should it be interpreted as
   an instant on a particular day (``1904-01-01T15:00:00``)?

   The functions in :mod:`~xlrd.xldate` take the view that such a number is a
   calendar-independent time of day (like Python's :class:`datetime.time` type)
   for both date systems. This is consistent with more recent Microsoft
   documentation. For example, the help file for Excel 2002, which says that the
   first day in the 1904 date system is ``1904-01-02``.

5. Usage of the Excel ``DATE()`` function may leave strange dates in a
   spreadsheet. Quoting the help file in respect of the 1900 date system::

     If year is between 0 (zero) and 1899 (inclusive),
     Excel adds that value to 1900 to calculate the year.
     For example, DATE(108,1,2) returns January 2, 2008 (1900+108).

   This gimmick, semi-defensible only for arguments up to 99 and only in the
   pre-Y2K-awareness era, means that ``DATE(1899, 12, 31)`` is interpreted as
   ``3799-12-31``.

   For further information, please refer to the documentation for the
   functions in :mod:`~xlrd.xldate`.
