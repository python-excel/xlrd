XML vulnerabilities and Excel files
===================================

If your code ingests ``.xlsx`` files that come from sources in which you do not
have absolute trust, please be aware that ``.xlsx`` files are made up of XML
and, as such, are susceptible to the vulnerabilities of XML.

xlrd uses ElementTree to parse XML, but as you'll find if you look into it,
there are many different ElementTree implementations. A good summary
of vulnerabilities you should worry can be found here:
:ref:`xml-vulnerabilities`.

For clarity, xlrd will try and import ElementTree from the following sources.
The list is in priority order, with those earlier in the list being preferred
to those later in the list:

1. `xml.etree.cElementTree`__

   __ https://docs.python.org/2/library/xml.etree.elementtree.html

2. `cElementTree`__

   __ http://effbot.org/zone/celementtree.htm

3. `lxml.etree`__

   __ http://lxml.de/api/lxml.etree-module.html

4. `xml.etree.ElementTree`__

   __ https://docs.python.org/2/library/xml.etree.elementtree.html

5. `elementtree.ElementTree`__

   __ http://effbot.org/zone/element-index.htm

To guard against these problems, you should consider the `defusedxml`__
project which can be used as follows:

__ https://pypi.org/project/defusedxml/

.. code-block:: python

    import defusedxml
    from defusedxml.common import EntitiesForbidden
    from xlrd import open_workbook
    defusedxml.defuse_stdlib()


    def secure_open_workbook(**kwargs):
        try:
            return open_workbook(**kwargs)
        except EntitiesForbidden:
            raise ValueError('Please use a xlsx file without XEE')
