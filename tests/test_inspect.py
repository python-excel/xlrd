from xlrd import inspect_format

from .helpers import from_sample


def test_xlsx():
    assert inspect_format(from_sample('sample.xlsx')) == 'xlsx'


def test_xlsb():
    assert inspect_format(from_sample('sample.xlsb')) == 'xlsb'


def test_ods():
    assert inspect_format(from_sample('sample.ods')) == 'ods'


def test_zip():
    assert inspect_format(from_sample('sample.zip')) == 'zip'


def test_xls():
    assert inspect_format(from_sample('namesdemo.xls')) == 'xls'


def test_content():
    with open(from_sample('sample.xlsx'), 'rb') as source:
        assert inspect_format(content=source.read()) == 'xlsx'


def test_unknown():
    assert inspect_format(from_sample('sample.txt')) is None
