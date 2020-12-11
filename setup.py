from setuptools import setup

from xlrd.info import __VERSION__

setup(
    name='xlrd',
    version=__VERSION__,
    author='Chris Withers',
    author_email='chris@withers.org',
    url='http://www.python-excel.org/',
    packages=['xlrd'],
    scripts=[
        'scripts/runxlrd.py',
    ],
    description=(
        'Library for developers to extract data from '
        'Microsoft Excel (tm) .xls spreadsheet files'
    ),
    long_description=open('README.rst').read(),
    license='BSD',
    keywords=['xls', 'excel', 'spreadsheet', 'workbook'],
    classifiers=[
        'Development Status :: 5 - Production/Stable',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: BSD License',
        'Programming Language :: Python',
        'Programming Language :: Python :: 2',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        'Operating System :: OS Independent',
        'Topic :: Database',
        'Topic :: Office/Business',
        'Topic :: Software Development :: Libraries :: Python Modules',
    ],
    python_requires=">=2.7, !=3.0.*, !=3.1.*, !=3.2.*, !=3.3.*, !=3.4.*, !=3.5.*",
    extras_require=dict(
        test=['pytest', 'pytest-cov'],
        docs=['sphinx'],
        build=['wheel', 'twine']
    )
)
