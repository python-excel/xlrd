###############################################################################
#
# Test xls with high memory usage even though mmap is used.
#

import unittest
import time

import xlrd
from .base import from_this_dir

# ----------------------------- test memory without psutil --------------------------
# solution from : https://stackoverflow.com/a/30316760/3655984
import sys
from gc import get_referents
from types import ModuleType, FunctionType


# Custom objects know their class.
# Function objects seem to know way too much, including modules.
# Exclude modules as well.
BLACKLIST = type, ModuleType, FunctionType


def getsize(obj):
    """sum size of object & members."""
    if isinstance(obj, BLACKLIST):
        raise TypeError('getsize() does not take argument of type: '+ str(type(obj)))
    seen_ids = set()
    size = 0
    objects = [obj]
    while objects:
        need_referents = []
        for obj in objects:
            if not isinstance(obj, BLACKLIST) and id(obj) not in seen_ids:
                seen_ids.add(id(obj))
                size += sys.getsizeof(obj)
                need_referents.append(obj)
        objects = get_referents(*need_referents)
    return size

# ------------------------------------------------------------------------------------


class TestXlsMMAP(unittest.TestCase):
    # Test opening xlsx on_demand

    def test_xls_on_demand_mmap(self):
        test_file = from_this_dir('test_high_mem_mmap.xls')
        workbook_size_old = 103761454   # bytes

        # with on_demand, everything should be quick
        with xlrd.open_workbook(test_file, on_demand=True) as workbook:
            # access the last sheet
            worksheet = workbook.sheet_by_name('VERB')
            self.assertEqual("stub_data", worksheet.cell(rowx=56, colx=10).value,
                             "Result should be the same as without on_demand")

            workbook_size_new = getsize(workbook)

        # with the new code path, the size of workbook should be very small
        self.assertLessEqual(50, workbook_size_old / workbook_size_new)
