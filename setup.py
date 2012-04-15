#!/usr/bin/env python

from os import path
import sys
python_version = sys.version_info[:2]

av = sys.argv
if len(av) > 1 and av[1].lower() == "--egg":
    if python_version < (2, 3):
        raise Exception("Can't lay eggs with Python version %d.%d " % python_version)
    del av[1]
    from setuptools import setup
else:
    from distutils.core import setup

the_url = 'http://www.lexicon.net/sjmachin/xlrd.htm'

# Get version number without importing xlrd/__init__
# (this horrificness is needed while using 2to3 for
#  python 3 compatibility, it should go away once
#  we stop using that.)

__file__ = path.abspath(sys.argv[0])
sys.path.insert(0, path.join(path.dirname(__file__), 'xlrd'))
from info import __VERSION__
sys.path.pop(0)

def mkargs(**kwargs):
    return kwargs

args = mkargs(
    name = 'xlrd',
    version = __VERSION__,
    author = 'John Machin',
    author_email = 'sjmachin@lexicon.net',
    url = the_url,
    packages = ['xlrd'],
    scripts = [
        'scripts/runxlrd.py',
        ],
    description = 'Library for developers to extract data from Microsoft Excel (tm) spreadsheet files',
    long_description = \
        "Extract data from Excel spreadsheets (XLS only, versions 2.0 to 2003) on any platform. " \
        "Pure Python (2.1 to 2.7). Strong support for Excel dates. Unicode-aware.",
    platforms = ["Any platform -- don't need Windows"],
    license = 'BSD',
    keywords = ['xls', 'excel', 'spreadsheet', 'workbook'],
    )

if python_version >= (2, 3):
    args23 = mkargs(
        download_url = the_url,
        classifiers = [
            'Development Status :: 5 - Production/Stable',
            'Intended Audience :: Developers',
            'License :: OSI Approved :: BSD License',
            'Programming Language :: Python',
            'Operating System :: OS Independent',
            'Topic :: Database',
            'Topic :: Office/Business',
            'Topic :: Software Development :: Libraries :: Python Modules',
            ],
        )
    args.update(args23)

if python_version >= (2, 4):
    args24 = mkargs(
        package_data={
            'xlrd': [
                'doc/*.htm*',
                # 'doc/*.txt',
                'examples/*.*',
                ],

            },
        )
    args.update(args24)

setup(**args)
