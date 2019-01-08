Development
===========

.. highlight:: bash

This package is developed using continuous integration which can be
found here:

https://travis-ci.org/python-excel/xlrd

If you wish to contribute to this project, then you should fork the
repository found here:

https://github.com/python-excel/xlrd

Once that has been done and you have a checkout, you can follow these
instructions to perform various development tasks:

Setting up a virtualenv
-----------------------

The recommended way to set up a development environment is to turn
your checkout into a virtualenv and then install the package in
editable form as follows::

  $ virtualenv .
  $ bin/pip install -e .

Running the tests
-----------------

Once you've set up a virtualenv, the tests can be run as follows::

  $ python -m unittest discover

To run tests on all the versions of Python that are supported, you can do::

  $ bin/tox

If you change the supported python versions in ``.travis.yml``, please remember
to do the following to update ``tox.ini``::

  $ bin/panci --to=tox .travis.yml > tox.ini

Building the documentation
--------------------------

The Sphinx documentation is built by doing the following, having activated
the virtualenv above, from the directory containing setup.py::

  $ cd docs
  $ make html

Making a release
----------------

To make a release, just update the version in ``xlrd.info.__VERSION__``,
update the change log, tag it, push to https://github.com/python-excel/xlrd
and Travis CI should take care of the rest.

Once the above is done, make sure to go to
https://readthedocs.org/projects/xlrd/versions/
and make sure the new release is marked as an Active Version.
