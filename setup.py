#!/usr/bin/env python
# -*- coding: ascii -*-

import sys
from distutils.core import setup

the_url = 'http://www.lexicon.net/sjmachin/xlrd.htm'

setup(name = 'xlrd',
      version = '0.3a1',
      author = 'John Machin',
      author_email = 'sjmachin@lexicon.net',
      url = the_url,
      download_url = the_url,
      packages = ['xlrd'],
      scripts = ['scripts/runxlrd.py'],
      package_data={'xlrd': ['doc/*.htm*', 'doc/*.txt']},
      description = 'Library for developers to extract data from Microsoft Excel (tm) spreadsheet files',
      long_description = \
            "Extract data from new and old Excel spreadsheets on any platform. " \
            "Pure Python code. Strong support for Excel dates. Unicode-aware.",
      platforms = ["Any platform -- don't need Windows"],
      license = 'BSD',
      keywords = ['xls', 'excel', 'spreadsheet', 'workbook'],
      classifiers = [
        'Development Status :: 3 - Alpha',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: BSD License',
        'Programming Language :: Python',
        'Operating System :: OS Independent',
        'Topic :: Database',
        'Topic :: Office/Business',
        'Topic :: Software Development :: Libraries :: Python Modules',      
        ]
      )

# 0.3a1.1 Corrected URL