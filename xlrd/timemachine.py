# -*- coding: ascii -*-

##
# <p>Copyright (c) 2006-2012 Stephen John Machin, Lingfo Pty Ltd</p>
# <p>This module is part of the xlrd package, which is released under a BSD-style licence.</p>
##

# timemachine.py -- adaptation for single codebase.
# Currently supported: 2.1 to 2.7
# usage: from timemachine import *

from __future__ import nested_scopes
import sys

python_version = sys.version_info[:2] # e.g. version 2.4 -> (2, 4)

CAN_PICKLE_ARRAY = python_version >= (2, 5)
CAN_SUBCLASS_BUILTIN = python_version >= (2, 2)

if python_version >= (3, 0): # Might work on 3.0 but absolutely no support!
    BYTES_LITERAL = lambda x: x.encode('latin1')
    BYTES_ORD = lambda byte: byte
    BYTES_NULL = bytes(0)   # b''
    BYTES_X00  = bytes(1)   # b'\x00'
    BYTES_X01  = bytes([1]) # b'\x01'
    from io import BytesIO as BYTES_IO
    def fprintf(f, fmt, *vargs):
        fmt = fmt.replace("%r", "%a")
        f.write(fmt % vargs)
    EXCEL_TEXT_TYPES = (str, bytes, bytearray) # xlwt: isinstance(obj, EXCEL_TEXT_TYPES)
    REPR = ascii
else:
    BYTES_LITERAL = lambda x: x
    BYTES_ORD = ord
    BYTES_NULL = ''
    BYTES_X00  = '\x00'
    BYTES_X01  = '\x01'
    from cStringIO import StringIO as BYTES_IO
    def fprintf(f, fmt, *vargs):
        f.write(fmt % vargs)
    try:
        EXCEL_TEXT_TYPES = basestring # xlwt: isinstance(obj, EXCEL_TEXT_TYPES)
    except NameError:
        EXCEL_TEXT_TYPES = (str, unicode)
    REPR = repr

if python_version >= (2, 6):
    def BUFFER(obj, offset=0, size=None):
        if size is None:
            return memoryview(obj)[offset:]
        return memoryview(obj)[offset:offset+size]
else:
    BUFFER = buffer

try:
    from array import array as array_array
except ImportError:
    # old version of IronPython?
    array_array = None

try:
    object
except NameError:
    class object:
        pass

try:
    True
except NameError:
    setattr(sys.modules['__builtin__'], 'True', 1)

try:
    False
except NameError:
    setattr(sys.modules['__builtin__'], 'False', 0)
    
if python_version < (2, 2):
    def has_key(d, key):
        return d.has_key(key)

else:
    def has_key(d, key):
        return key in d

def int_floor_div(x, y):
    return divmod(x, y)[0]

def intbool(x):
    if x:
        return 1
    return 0

if python_version < (2, 3):
    def sum(sequence, start=0):
        tot = start
        for item in aseq:
            tot += item
        return tot

if python_version >= (3,):
    # Python 3
    def b(s):
        return s.encode('latin1')
    
    def get_int_1byte(data, pos):
        return data[pos]
    
else:
    # Python 2
    def b(s): return s

    def get_int_1byte(data, pos):
        return ord(data[pos])

byte_0 = b('\x00')
byte_empty = b('')
