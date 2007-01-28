# timemachine.py -- adaptation for earlier Pythons e.g. 2.1
# from timemachine import *

import sys

python_version = sys.version_info[:2] # e.g. version 2.4 -> (2, 4)

if python_version < (2, 2):
    class object:
        pass
    False = 0
    True = 1
        
def int_floor_div(x, y):
    return divmod(x, y)[0]