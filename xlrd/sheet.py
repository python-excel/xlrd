##
# Part of the xlrd package.
##

from biffh import *
from struct import unpack
import array

DEBUG = 0

def fprintf(f, fmt, *vargs): print >> f, fmt % vargs,

##
# <p>Contains the data for one worksheet.</p>
#
# <p>In the cell access functions, "rowx" is a row index, counting from zero, and "colx" is a
# column index, counting from zero.
# Negative values for row/column indexes and slice positions are supported in the expected fashion.</p>
#
# <p>For information about cell types and cell values, refer to the documentation of the Cell class.</p>
#
# <p>WARNING: You don't call this class yourself. You access Sheet objects via the Book object that
# was returned when you called xlrd.open_workbook("myfile.xls").</p>

class Sheet(object):
    ##
    # Name of sheet.
    name = ''
    ##
    # Number of rows in sheet. A row index is in range(thesheet.nrows).
    nrows = 0
    ##
    # Number of columns in sheet. A column index is in range(thesheet.ncols).
    ncols = 0

    def __init__(self, biff_version, position, logfile):
        self.biff_version = biff_version
        self._position = position
        self.logfile = logfile
        self.name = ''
        self.nrows = 0
        self.ncols = 0
        self._cell_values = []
        self._cell_types = []

    ##
    # Value of the cell in the given row and column.
    def cell_value(self, rowx, colx):
        return self._cell_values[rowx][colx]

    ##
    # Type of the cell in the given row and column. Refer to the documentation of the Cell class.
    def cell_type(self, rowx, colx):
        return self._cell_types[rowx][colx]

    ##
    # Returns a sequence of the Cell objects in the given row.
    def row(self, rowx):
        return [
            Cell(self._cell_types[rowx][colx], self._cell_values[rowx][colx])
            for colx in xrange(self.ncols)
            ]

    ##
    # Returns a sequence of the types
    # of the cells in the given row.
    def row_types(self, rowx):
        return self._cell_types[rowx]

    ##
    # Returns a sequence of the values
    # of the cells in the given row.
    def row_values(self, rowx):
        return self._cell_values[rowx]

    ##
    # Returns a slice of the Cell objects in the given row.
    def row_slice(self, rowx, start_colx=0, end_colx=None):
        nc = self.ncols
        if start_colx < 0:
            start_colx += nc
            if start_colx < 0:
                start_colx = 0
        if end_colx is None or end_colx > nc:
            end_colx = nc
        elif end_colx < 0:
            end_colx += nc
        return [
            Cell(self._cell_types[rowx][colx], self._cell_values[rowx][colx])
            for colx in xrange(start_colx, end_colx)
            ]

    ##
    # Returns a slice of the Cell objects in the given column.
    def col_slice(self, colx, start_rowx=0, end_rowx=None):
        nr = self.nrows
        if start_rowx < 0:
            start_rowx += nr
            if start_rowx < 0:
                start_rowx = 0
        if end_rowx is None or end_rowx > nr:
            end_rowx = nr
        elif end_rowx < 0:
            end_rowx += nr
        return [
            Cell(self._cell_types[rowx][colx], self._cell_values[rowx][colx])
            for rowx in xrange(start_rowx, end_rowx)
            ]

    ##
    # Returns a sequence of the Cell objects in the given column.
    def col(self, colx):
        return self.col_slice(colx)
    # Above two lines just for the docs. Here's the real McCoy:
    col = col_slice

    # === Following methods are used in building the worksheet.
    # === They are not part of the API.

    def initcells(self):
        nc = self.ncols
        aa = array.array
        scta = self._cell_types.append
        scva = self._cell_values.append
        xce = XL_CELL_EMPTY
        for _unused in xrange(self.nrows):
            scta(aa('B', [xce]) * nc)
            scva([''] * nc)

    def put_cell(self, rowx, colx, ctype, value):
        self._cell_types[rowx][colx] = ctype
        self._cell_values[rowx][colx] = value

    def put_number_cell(self, rowx, colx, value, fmt_ty=FNU):
        self._cell_types[rowx][colx] = cellty_from_fmtty[fmt_ty]
        self._cell_values[rowx][colx] = value

    # === Methods after this line neither know nor care about how cells are stored.

    def read(self, bk):
        global rc_stats
        DEBUG = 0
        oldpos = bk._position
        bk.position(self._position)
        DEBUG = 0
        XL_SHRFMLA_ETC_ETC = (XL_SHRFMLA, XL_ARRAY, XL_TABLEOP, XL_TABLEOP2, XL_TABLEOP_B2)
        while 1:
            # if DEBUG: print "SHEET.READ: about to read from position %d" % bk._position
            rc, length, data = bk.get_record_parts()
            # if rc in rc_stats:
            #     rc_stats[rc] += 1
            # else:
            #     rc_stats[rc] = 1
            # if DEBUG: print "SHEET.READ: op 0x%04x, %d bytes %r" % (rc, len(data), data)
            if rc == XL_NUMBER:
                rowx, colx, xfindex, d = unpack('<HHHd', data)
                # if DEBUG: printf("NUMBER Double 8 byte: %d %d %d %f\n", rowx, colx, xfindex, d)
                fty = check_xf(bk, rowx, colx, xfindex, d)
                self.put_number_cell(rowx, colx, d, fty)
            elif rc == XL_LABELSST:
                rowx, colx, index = unpack('<HHxxi', data)
                self.put_cell(rowx, colx, XL_CELL_TEXT, bk._sharedstrings[index])
            elif rc == XL_LABEL or rc == XL_RSTRING:
                # RSTRING has extra richtext info at the end, but we ignore it.
                rowx, colx, xfindex = unpack('<HHH', data[0:6])
                if self.biff_version < BIFF_FIRST_UNICODE:
                    strg = unpack_string(data, 6, bk.encoding, lenlen=2)
                else:
                    strg = unpack_unicode(data, 6, lenlen=2)
                self.put_cell(rowx, colx, XL_CELL_TEXT, strg)
            elif rc == XL_RK:
                rowx, colx, xfindex = unpack('<HHH', data[:6])
                d = unpack_RK(data[6:10])
                # if DEBUG: printf("RK Double 4 byte: %f\n",d);
                fty = check_xf(bk, rowx, colx, xfindex, d)
                self.put_number_cell(rowx, colx, d, fty);
            elif rc == XL_MULRK:
                mulrk_row, mulrk_first = unpack('<HH', data[0:4])
                mulrk_last  = unpack('<H', data[-2:])[0]
                # mulrk_numrks = mulrk_last - mulrk_first + 1
                # if DEBUG: printf("MulRK first: %d last: %d records: %d\n",mulrk_first,mulrk_last,mulrk_numrks);
                pos = 4
                for colx in xrange(mulrk_first, mulrk_last+1):
                    xfindex = unpack('<H', data[pos:pos+2])[0]
                    d = unpack_RK(data[pos+2:pos+6])
                    # printf("MULRK r%d c%d: %s -> %f\n",
                    #     mulrk_row, colx, ''.join(["%02x " % ord(c) for c in data[pos+2:pos+6]]), d);
                    fty = check_xf(bk, mulrk_row, colx, xfindex, d)
                    pos += 6
                    self.put_number_cell(mulrk_row, colx, d, fty)
            elif rc == XL_ROW:
                # We don't use the ROW record ... but there are enough of them to warrant
                # this being here to save exhaustive testing.
                # A dictionary of known-but-ignored worksheet records would be a better idea.
                # Would need separate dicts, one for each version.
                pass
            elif rc & 0xff == XL_FORMULA: # 06, 0206, 0406
                # if DEBUG: print "FORMULA: rc: 0x%04x data: %r" % (rc, data)
                rowx, colx, xfindex = unpack('<HHH', data[0:6])
                # if DEBUG: print "FORMULA: rowx=%d colx=%d" % (rowx, colx)
                if data[12] == '\xff' and data[13] == '\xff':
                    if data[6] == '\x00':
                        # need to read next record (STRING)
                        gotstring = 0
                        if ord(data[14]) & 8:
                            # actually there's an optional SHRFMLA or ARRAY etc record to skip over
                            rc2, _unused_len, data2 = bk.get_record_parts()
                            if rc2 == XL_STRING:
                                gotstring = 1
                            elif rc2 not in XL_SHRFMLA_ETC_ETC:
                                raise XLRDError(
                                    "Expected SHRFMLA, ARRAY, TABLEOP* or STRING record; found 0x%04x" % rc2)
                            # if DEBUG: print "gotstring:", gotstring
                        # now for the STRING record
                        if not gotstring:
                            rc2, _unused_len, data2 = bk.get_record_parts()
                            if rc2 != XL_STRING: raise XLRDError("Expected STRING record; found 0x%04x" % rc2)
                        # if DEBUG: print "STRING: data=%r BIFF=%d cp=%d" % (data2, self.biff_version, bk.encoding)
                        if self.biff_version < BIFF_FIRST_UNICODE:
                            strg = unpack_string(data2, 0, bk.encoding, lenlen=2)
                        else:
                            strg = unpack_unicode(data2, 0, lenlen=2)
                        self.put_cell(rowx, colx, XL_CELL_TEXT, strg)
                        # if DEBUG: print "FORMULA strg %r" % strg
                    elif data[6] == '\x01':
                        # boolean formula result
                        value = ord(data[8])
                        self.put_cell(rowx, colx, XL_CELL_BOOLEAN, value)
                    elif data[6] == '\x02':
                        # Error in cell
                        value = ord(data[8])
                        self.put_cell(rowx, colx, XL_CELL_ERROR, value)
                    elif data[6] == '\x03':
                        # empty cell
                        # Do nothing, its place in the grid is already bound to empty_cell
                        pass
                    else:
                        raise XLRDError("unexpected special case (0x%02x) in FORMULA" % ord(data[6]))
                else:
                    # it is a number
                    d = unpack('<d', data[6:14])[0]
                    fty = check_xf(bk, rowx, colx, xfindex, d)
                    self.put_number_cell(rowx, colx, d, fty)
            elif rc == XL_BOOLERR:
                rowx, colx, xfindex, value, is_err = unpack('<HHHBB', data)
                cellty = (XL_CELL_BOOLEAN, XL_CELL_ERROR)[is_err]
                # if DEBUG: print "XL_BOOLERR", rowx, colx, xfindex, value, is_err
                self.put_cell(rowx, colx, cellty, value)
            elif rc == XL_DIMENSION:
                if length == 10:
                    self.nrows, self.ncols = unpack('<HxxH', data[2:8])
                else:
                    self.nrows, self.ncols = unpack('<ixxH', data[4:12])
                self.initcells()
                if DEBUG: fprintf(self.logfile, "Dimension ncols: %d nrows: %d\n", self.ncols, self.nrows)
            elif rc == XL_EOF:
                DEBUG = 0
                if DEBUG: print >> self.logfile, "SHEET.READ: EOF"
                break
            elif rc == XL_OBJ:
                bk.handle_obj(data)
            elif rc in boflen: ##### EMBEDDED BOF #####
                version, boftype = unpack('<HH', data[0:4])
                if boftype != 0x20: # embedded chart
                    print >> self.logfile, \
                        "*** Unexpected embedded BOF (0x%04x) at offset %d: version=0x%04x type=0x%04x" \
                        % (rc, bk._position - length - 4, version, boftype)
                while 1:
                    code, length, data = bk.get_record_parts()
                    if code == XL_EOF:
                        break
                if DEBUG: print >> self.logfile, "---> found EOF"
            elif rc == XL_COUNTRY:
                bk.handle_country(data)
            #### all of the following are for BIFF <= 4.0
            elif rc == XL_FORMAT or rc == XL_FORMAT2:
                bk.handle_format(data)
            elif rc == XL_BUILTINFMTCOUNT:
                bk.handle_builtinfmtcount(data)
            elif rc == XL_XF4 or rc == XL_XF3: #### N.B. not XL_XF
                bk.handle_xf(data)
            elif rc == XL_DATEMODE:
                bk.handle_datemode(data)
            elif rc == XL_CODEPAGE:
                bk.handle_codepage(data)
            elif rc == XL_FILEPASS:
                bk.handle_filepass(data)
            elif rc == XL_WRITEACCESS:
                bk.handle_writeaccess(data)
            else:
                # if DEBUG: print "SHEET.READ: Unhandled record type %02x %d bytes %r" % (rc, len(data), data)
                pass
        bk.position(oldpos)
        return 1

# === helpers ===

def unpack_RK(rk_str):
    flags = ord(rk_str[0])
    if flags & 2:
        # There's a SIGNED 30-bit integer in there!
        d = float(unpack('<i', rk_str)[0] // 4) # div by 4 to drop the 2 flag bits
    else:
        # It's the most significant 30 bits of an IEEE 754 64-bit FP number
        d = unpack('<d', '\0\0\0\0' + chr(flags & 252) + rk_str[1:4])[0]
    if flags & 1:
        return d / 100.0
    return d

def check_xf(bk, rowx, colx, xfindex, value):
    xfrec = bk.xfrecords[xfindex]
    if xfrec:
        # if DEBUG: print "OK xfindex %d; rowx=%d colx=%d value=%r" % (xfindex, rowx, colx, value)
        return xfrec.type
    else:
        print >> bk.logfile, "*** No XF for xfindex %d; rowx=%d colx=%d value=%r" % (xfindex, rowx, colx, value)
        return None

##### =============== Cell ======================================== #####

cellty_from_fmtty = {
    FNU: XL_CELL_NUMBER,
    FUN: XL_CELL_NUMBER,
    FGE: XL_CELL_NUMBER,
    FDT: XL_CELL_DATE,
    FTX: XL_CELL_NUMBER, # Yes, a number can be formatted as text.
    }

ctype_text = {
    XL_CELL_EMPTY: 'empty',
    XL_CELL_TEXT: 'text',
    XL_CELL_NUMBER: 'number',
    XL_CELL_DATE: 'xldate',
    XL_CELL_BOOLEAN: 'bool',
    XL_CELL_ERROR: 'error',
    }

##
# <p>Contains the data for one cell.</p>
#
# <p>WARNING: You don't call this class yourself. You access Cell objects
# via methods of the Sheet object(s) that you found in the Book object that
# was returned when you called xlrd.open_workbook("myfile.xls").</p>
# <p> Cell objects have two attributes: <i>ctype</i> is an int, and <i>value</i>
# which depends on <i>ctype</i>.
# The following table describes the types of cells and how their values
# are represented in Python.</p>
#
# <table border="1" cellpadding="7">
# <tr>
# <th>Type symbol</th>
# <th>Type number</th>
# <th>Python value</th>
# </tr>
# <tr>
# <td>XL_CELL_EMPTY</td>
# <td align="center">0</td>
# <td>empty string u''</td>
# </tr>
# <tr>
# <td>XL_CELL_TEXT</td>
# <td align="center">1</td>
# <td>a Unicode string</td>
# </tr>
# <tr>
# <td>XL_CELL_NUMBER</td>
# <td align="center">2</td>
# <td>float</td>
# </tr>
# <tr>
# <td>XL_CELL_DATE</td>
# <td align="center">3</td>
# <td>float</td>
# </tr>
# <tr>
# <td>XL_CELL_BOOLEAN</td>
# <td align="center">4</td>
# <td>int; 1 means TRUE, 0 means FALSE</td>
# </tr>
# <tr>
# <td>XL_CELL_ERROR</td>
# <td align="center">5</td>
# <td>int representing internal Excel codes; for a text representation,
# refer to the supplied dictionary error_text_from_code</td>
# </tr>
# </table>
#<p></p>

class Cell(object):

    __slots__ = ['ctype', 'value',]

    def __init__(self, ctype, value):
        self.ctype = ctype
        self.value = value

    def __repr__(self):
        return "%s:%r" % (ctype_text[self.ctype], self.value)

##
# There is one and only one instance of an empty cell -- it's a singleton. This is it.
# You may use a test like "acell is empty_cell".
empty_cell = Cell(XL_CELL_EMPTY, '')

# === grimoire ===

try:
    from _xlrdutils import *
    # print "_xlrdutils imported"
except ImportError:
    # print "_xlrdutils *NOT* imported"
    pass
