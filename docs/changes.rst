Changes
=======

.. currentmodule:: xlrd

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


