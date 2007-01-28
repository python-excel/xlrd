# -*- coding: cp1252 -*-

import licences

##
# <p><b>A Python module for extracting data from MS Excel ™ spreadsheet files.</b></p>
#
# <h2>General information</h2>
# <h3>Unicode</h3>
# <p>This module presents all text strings as Python unicode objects.
# From Excel 97 onwards, text in Excel spreadsheets has been stored as Unicode.
# Earlier spreadsheets have a "codepage" number indicating the local representation; this
# is used to derive an "encoding" which is used to translate to Unicode.
#
# <h3>Dates in Excel spreadsheets</h3>
# <p>In reality, there are no such things. What you have are floating point numbers and pious hope.
# There are several problems with Excel dates:</p>
#
# <p>(1) Dates are not stored as a separate data type; they are stored as floating point numbers
# and you have to rely on (a) the "number format" applied to them in Excel and/or (b) knowing
# which cells are supposed to have dates in them. This module helps with (a) by inspecting the
# format that has been applied to each number cell; if it appears to be a date format, the cell
# is classified as a date rather than a number. Feedback on this feature,
# especially from non-English-speaking locales, would be appreciated.</p>
#
# <p>(2) Excel for Windows stores dates by default as the number of days (or fraction thereof) since 1899-12-31T00:00:00.
# Excel for Macintosh uses a default start date of 1904-01-01T00:00:00. The date system can be changed in Excel
# on a per-workbook basis (for example: Tools -> Options -> Calculation, tick the "1904 date system" box).
# This is of course a bad idea if there are already dates in the workbook. There is no good reason to change it 
# even if there are no dates in the workbook. Which date system is in use is recorded in the workbook. 
# A workbook transported from Windows to Macintosh (or vice versa) will work correctly with the host Excel.
# When using this module's xldate_as_tuple function to convert numbers from a workbook, you must use
# the datemode attribute of the Book object. If you guess, or make a judgement depending on where you
# believe the workbook was created, you run the risk of being 1462 days out of kilter.</p>
#
# <p>Reference: http://support.microsoft.com/default.aspx?scid=KB;EN-US;q180162</p>
#
# <p>(3) The Windows-default 1900-based date system works on the incorrect premise that 1900 was a leap year.
# It interprets
# the number 60 as meaning 1900-02-29, which is not a valid date. Consequently any number less than 61
# is ambiguous. Example: is 59 the result of 1900-02-28 entered directly, or is it 1900-03-01 minus 2 days?</p>
#
# <p>Reference: http://support.microsoft.com/default.aspx?scid=kb;en-us;214326</p>
#
# <p>(4) The Macintosh-default 1904-based date system counts 1904-01-02 as day 1 and 1904-01-01 as day zero.
# Thus any number such that (0.0 <= number < 1.0) is ambiguous. Is 0.625 a time of day (15:00:00),
# independent of the calendar,
# or should it be interpreted as an instant on a particular day (1904-01-01T15:00:00)?
# The xldate_* functions in this module
# take the view that such a number is a calendar-independent time of day (like Python's datetime.time type) for both
# date systems. This is consistent with more recent Microsoft documentation
# (for example, the help file for Excel 2002 which says that the first day
# in the 1904 date system is 1904-01-02).
#
# <p>(5) Usage of the Excel DATE() function may leave strange dates in a spreadsheet. Quoting the help file,
# in respect of the 1900 date system: "If year is between 0 (zero) and 1899 (inclusive),
# Excel adds that value to 1900 to calculate the year. For example, DATE(108,1,2) returns January 2, 2008 (1900+108)."
# This gimmick, semi-defensible only for arguments up to 99 and only in the pre-Y2K-awareness era,
# means that DATE(1899, 12, 31) is interpreted as 3799-12-31.</p> 
#
# <p>For further information, please refer to the documentation for the xldate_* functions.</p>
##

from biffh import *
from struct import unpack
import sys
import sheet
import compdoc
from xldate import xldate_as_tuple, XLDateError
empty_cell = sheet.empty_cell # for exposure to the world ...

#  MS article on Excel ODBC: http://support.microsoft.com/kb/q141284/

DEBUG = 0

USE_MMAP = 1
USE_FANCY_CD = 1

if USE_MMAP:
    try:
        import mmap
    except ImportError:
        USE_MMAP = 0

MY_EOF = 0xF00BAAA # not a 16-bit number

def fprintf(f, fmt, *vargs):
    print >> f, fmt % vargs,

SUPPORTED_VERSIONS = (80, 70, 50, 45, 40, 30)

##
# Open a spreadsheet file for data extraction.
# @param filename The path to the spreadsheet file to be opened.
# @param logfile An open file to which messages and diagnostics are written.
# @param verbosity Increases the volume of trace material written to the logfile.
# @return An instance of the Book class.

def open_workbook(filename, logfile=sys.stdout, verbosity=0):
    bk = Book(filename, logfile=logfile, verbosity=verbosity)
    biff_version = bk.getbof(XL_WORKBOOK_GLOBALS)
    if not biff_version:
        raise XLRDError("Can't determine file's BIFF version")
    if biff_version not in SUPPORTED_VERSIONS:
        raise XLRDError("BIFF version %s is not supported" % biff_text_from_num[biff_version])
    bk.biff_version = biff_version
    if biff_version <= 40:
        # no workbook globals, only 1 worksheet
        bk.fake_globals_get_sheet()
    elif biff_version == 45:
        # worksheet(s) embedded in global stream
        bk.parse_globals()
    else:
        bk.parse_globals()
        bk.get_sheets()
    bk.nsheets = len(bk._sheet)
    bk.release_resources()
    return bk

##
# For debugging: dump the file's BIFF records in char & hex.
# @param filename The path to the file to be dumped.
# @param outfile An open file, to which the dump is written.

def dump(filename, outfile=sys.stdout):
    bk = Book(filename)
    biff_dump(bk.mem, bk.base, bk.stream_len, 0, outfile)

##
# Contents of a "workbook".
# <p>WARNING: You don't call this class yourself. You use the Book object that
# was returned when you called xlrd.open_workbook("myfile.xls").</p>

class Book(object):

    ##
    # The number of worksheets in the workbook.
    nsheets = 0

    ##
    # Which date system was in force when this file was last saved.<br>
    #    0 => 1900 system (the Excel for Windows default).<br>
    #    1 => 1904 system (the Excel for Macintosh default).<br>
    datemode = None

    ##
    # Version of BIFF (Binary Interchange File Format) used to create the file.
    # Latest is 8.0 (represented here as 80), introduced with Excel 97.
    # Earliest supported by this module: 3.0 (rep'd as 30).
    biff_version = 0

    ##
    # An integer denoting the character set used for strings in this file.
    # For BIFF 8 and later, this will be 1200, meaning Unicode; more precisely, UTF_16_LE.
    # For earlier versions, this is used to derive the appropriate Python encoding
    # to be used to convert to Unicode.
    # Examples: 1252 -> 'cp1252', 10000 -> 'mac_roman'
    codepage = None

    ##
    # The encoding that was derived from the codepage.
    encoding = 'unknown'

    ##
    # A tuple containing the (telephone system) country code for:<br>
    #    [0]: the user-interface setting when the file was created.<br>
    #    [1]: the regional settings.<br>
    # Example: (1, 61) meaning (USA, Australia).
    # This information may give a clue to the correct encoding for an unknown codepage.
    # For a long list of observed values, refer to the OpenOffice.org documentation for
    # the COUNTRY record.
    countries = (0, 0)

    ##
    # What (if anything) is recorded as the name of the last user to save the file.
    user_name = ''

    ##
    # @param sheetx Sheet index in range(nsheets)
    # @return An object of the Sheet class
    def sheet_by_index(self, sheetx):
        return self._sheet[sheetx]

    ##
    # @param sheet_name Name of sheet required
    # @return An object of the Sheet class
    def sheet_by_name(self, sheet_name):
        try:
            sheetx = self._sheet_names.index(sheet_name)
        except ValueError:
            raise XLRDError('No sheet named <%r>' % sheet_name)
        return self._sheet[sheetx]

    ##
    # @return A list of the names of the sheets in the book.
    def sheet_names(self):
        return self._sheet_names[:]

    def __init__(self, filename, logfile=sys.stdout, verbosity=0):
        # DEBUG = 0
        self.logfile = logfile
        self.verbosity = verbosity
        self._sheet = []
        #### self.sheet = self._sheet ###### self.sheet is slated for removal RSN
        self._sheet_names = []
        #### self.sheet_names = self._sheet_names ##### self.sheet_names is slated for removal RSN
        self.nsheets = 0
        self._sh_abs_posn = [] # sheet's absolute position in the stream
        self._sharedstrings = []
        self.raw_user_name = False
        self._sheethdr_count = 0 # BIFF 4W only
        self.builtinfmtcount = -1 # unknown as yet. BIFF 3, 4S, 4W
        self.initialise_format_info()

        f = file(filename, "rb")
        if USE_MMAP:
            self.fileno = f.fileno()
            f.seek(0, 2) # EOF
            size = f.tell()
            f.seek(0, 0) # BOF
            filestr = mmap.mmap(self.fileno, size, access=mmap.ACCESS_READ)
            self.stream_len = size
        else:
            filestr = f.read()
            self.stream_len = len(filestr)
            f.close()
        self.base = 0
        if filestr[:8] != compdoc.SIGNATURE:
            # got this one at the antique store
            self.mem = filestr
        else:
            cd = compdoc.CompDoc(filestr)
            if USE_FANCY_CD:
                for qname in [u'Workbook', u'Book']:
                    self.mem, self.base, self.stream_len = cd.locate_named_stream(qname)
                    if self.mem: break
                else:
                    raise XLRDError("Can't find workbook in OLE2 compound document")
            else:
                for qname in [u'Workbook', u'Book']:
                    self.mem = cd.get_named_stream(qname)
                    if self.mem: break
                else:
                    raise XLRDError("Can't find workbook in OLE2 compound document")
                self.stream_len = len(self.mem)
            del cd
            if self.mem is not filestr:
                if USE_MMAP:
                    filestr.close()
                    f.close()
                else:
                    del filestr
        self._position = self.base
        if DEBUG:
            print >> self.logfile, "mem: %s, base: %d, len: %d" % (type(self.mem), self.base, self.stream_len)

    def initialise_format_info(self):
        # needs to be done once per sheet for BIFF 4W :-(
        self.format_dict = {}
        self.format_list = []
        self.xfcount = 0
        self.actualfmtcount = 0 # number of FORMAT records seen so far
        self.xfrecords = []
        self.xf_style_fmt_no = []

    def release_resources(self):
        del self.mem

    def get2bytes(self):
        pos = self._position
        buff = self.mem[pos:pos+2]
        lenbuff = len(buff)
        self._position += lenbuff
        if lenbuff < 2:
            return MY_EOF
        lo, hi = buff
        return (ord(hi) << 8) | ord(lo)

    def get_record_parts(self):
        pos = self._position
        mem = self.mem
        code, length = unpack('<HH', mem[pos:pos+4])
        pos += 4
        data = mem[pos:pos+length]
        self._position = pos + length
        return (code, length, data)

    def get_record_parts_conditional(self, reqd_record):
        pos = self._position
        mem = self.mem
        code, length = unpack('<HH', mem[pos:pos+4])
        if code != reqd_record:
            return (None, 0, '')
        pos += 4
        data = mem[pos:pos+length]
        self._position = pos + length
        return (code, length, data)

    def get_sheet(self):
        _unused_biff_version = self.getbof(XL_WORKSHEET)
        # assert biff_version == self.biff_version ### FAILS
        # Have an example where book is v7 but sheet reports v8!!!
        # It appears to work OK if the sheet version is ignored.
        # Confirmed by Daniel Rentz: happens when Excel does "save as"
        # creating an old version file; ignore version details on sheet BOF.
        sh = sheet.Sheet(self.biff_version, self._position, self.logfile)
        sh.read(self)
        return sh

    def get_sheets(self):
        # DEBUG = 0
        if DEBUG: print >> self.logfile, "GET_SHEETS:", self._sheet_names, self._sh_abs_posn
        for sheetno in xrange(len(self._sheet_names)):
            if DEBUG: print >> self.logfile, "GET_SHEETS: sheetno =", sheetno, self._sheet_names, self._sh_abs_posn
            newposn = self._sh_abs_posn[sheetno]
            self.position(newposn)
            sht = self.get_sheet()
            sht.name = self._sheet_names[sheetno]
            self._sheet.append(sht)

    def fake_globals_get_sheet(self): # for BIFF 4.0 and earlier
        fake_sheet_name = u'Sheet 1'
        self._sheet_names = [fake_sheet_name]
        self._sh_abs_posn = [0]
        self.get_sheets()

    def handle_boundsheet(self, data):
        # DEBUG = 0
        bv = self.biff_version
        if DEBUG: fprintf(self.logfile, "BOUNDSHEET: bv=%d data %r\n", bv, data);
        if bv == 45: # BIFF4W
            name = unpack_string(data, 0, self.encoding, lenlen=1)
            _unused_visibility = 0
            sheet_type = XL_BOUNDSHEET_WORKSHEET # guess
            if len(self._sh_abs_posn) == 0:
                abs_posn = self._sheetsoffset + self.base
                # Note (a) this won't be used
                # (b) it's the position of the SHEETHDR record
                # (c) add 11 to get to the worksheet BOF record
            else:
                abs_posn = -1 # unknown
        else:
            offset, _unused_visibility, sheet_type = unpack('<iBB', data[0:6])
            abs_posn = offset + self.base # because global BOF is always at posn 0 in the stream
            if bv < BIFF_FIRST_UNICODE:
                name = unpack_string(data, 6, self.encoding, lenlen=1)
            else:
                name = unpack_unicode(data, 6, lenlen=1)
        if DEBUG: fprintf(self.logfile, "BOUNDSHEET: name=%r abs_posn=%d sheet_type=0x%02x\n", name, abs_posn, sheet_type)
        if sheet_type != XL_BOUNDSHEET_WORKSHEET:
            descr = {
                1: 'Macro sheet',
                2: 'Chart',
                6: 'Visual Basic module',
                }.get(sheet_type, 'UNKNOWN')
            print >> self.logfile, \
                "*** BOUNDSHEET: Ignoring non-worksheet data (type 0x%02x = %s)" % (sheet_type, descr)
            return
        self._sheet_names.append(name)
        self._sh_abs_posn.append(abs_posn)

    def handle_builtinfmtcount(self, data):
        # DEBUG = 1
        builtinfmtcount = unpack('<H', data[0:2])[0]
        if DEBUG: fprintf(self.logfile, "BUILTINFMTCOUNT: %r\n", builtinfmtcount)
        self.builtinfmtcount = builtinfmtcount

    def handle_codepage(self, data):
        # DEBUG = 0
        codepage = unpack('<H', data[0:2])[0]
        self.codepage = codepage
        if codepage in encoding_from_codepage:
            encoding = encoding_from_codepage[codepage]
        elif 300 <= codepage <= 1999:
            encoding = 'cp' + str(codepage)
        else:
            encoding = 'unknown_codepage_' + str(codepage)
        if DEBUG or self.verbosity: fprintf(self.logfile, "CODEPAGE: codepage %r -> encoding %r\n", codepage, encoding)
        if codepage != 1200: # utf_16_le
            # If we don't have a codec that can decode ASCII into Unicode,
            # we're well & truly stuffed -- let the punter know ASAP.
            try:
                _unused = 'trial'.decode(encoding)
            except:
                ei = sys.exc_info()[:2]
                msg = "*** codepage %d -> encoding %r -> %s" \
                    % (codepage, encoding, ei[1])
                print >> self.logfile, msg
                print >> sys.stderr, msg
                raise
        self.encoding = encoding
        if self.raw_user_name:
            strg = unpack_string(self.user_name, 0, self.encoding, lenlen=1)
            strg = strg.rstrip()
            # if DEBUG:
            #     print "CODEPAGE: user name decoded from %r to %r" % (self.user_name, strg)
            self.user_name = strg
            self.raw_user_name = False

    def handle_country(self, data):
        countries = unpack('<HH', data[0:4])
        if self.verbosity: print >> self.logfile, "Countries:", countries
        # Note: in BIFF7 and earlier, country record was put (redundantly?) in each worksheet.
        assert self.countries == (0, 0) or self.countries == countries
        self.countries = countries

    def handle_datemode(self, data):
        datemode = unpack('<H', data[0:2])[0]
        if DEBUG or self.verbosity: fprintf(self.logfile, "DATEMODE: datemode %r\n", datemode)
        assert datemode in (0, 1)
        self.datemode = datemode

    def handle_filepass(self, data):
        raise XLRDError("Workbook is encrypted")

    def handle_format(self, data):
        DEBUG = 0
        bv = self.biff_version
        strpos = 2
        if bv >= 50:
            fmtcode = unpack('<H', data[0:2])[0]
        else:
            fmtcode = self.actualfmtcount
            if bv <= 30:
                strpos = 0
        self.actualfmtcount += 1
        if bv >= BIFF_FIRST_UNICODE:
            unistrg = unpack_unicode(data, 2)
        else:
            unistrg = unpack_string(data, strpos, self.encoding, lenlen=1)
        if DEBUG or self.verbosity >= 3:
            print "FORMAT: count=%d code=0x%04x (%d) s=%r" % (self.actualfmtcount, fmtcode, fmtcode, unistrg)
        is_date_s = self.is_date_format_string(unistrg)
        ty = std_format_code_types.get(fmtcode, FUN)
        is_date_c = ty == FDT
        if (fmtcode >= 163 # user_defined
        or bv < 50):
            is_date = is_date_s
        else:
            if fmtcode >= 0 and (is_date_c ^ is_date_s):
                DEBUG = 2
                print >> self.logfile, '\n****** Conflict between std format code and its fmt string ***'
            is_date = is_date_c | is_date_s
        if is_date:
            ty = FDT
        if DEBUG == 2: print >> self.logfile, "ty: %d; is_date_c: %r; is_date_s: %r; fmt_strg: %r" \
            % (ty, is_date_c, is_date_s, unistrg)
        xfrec = Format(fmtcode, ty, unistrg)
        self.format_dict[fmtcode] = xfrec
        self.format_list.append(xfrec)

    def handle_obj(self, data):
        # Not doing much handling at all.
        # Worrying about embedded (BOF ... EOF) substreams is done elsewhere.
        # DEBUG = 1
        obj_type, obj_id = unpack('<HI', data[4:10])
        # if DEBUG: print "---> handle_obj type=%d id=0x%08x" % (obj_type, obj_id)

    def handle_xf(self, data):
        # DEBUG = 0
        bv = self.biff_version
        # fill in the known standard formats
        if bv >= 50 and not self.xfcount:
            # i.e. do this once before we process the first XF record
            for x in std_format_code_types:
                if x not in self.format_dict:
                    ty = std_format_code_types[x]
                    xfrec = Format(x, ty, u'')
                    self.format_dict[x] = xfrec
        if bv >= 80:
            fmtcode, pkd_type_par, pkd_used = unpack('<2xHH3xB', data[0:10])
            is_style = (pkd_type_par & 4) == 4
            parent = (pkd_type_par & 0xfff0) >> 4
            used = ((pkd_used & 0xfc) >> 2) & 1
        elif bv >= 50:
            fmtcode, pkd_type_par, pkd_used = unpack('<2xHHxB', data[0:8])
            is_style = (pkd_type_par & 4) == 4
            parent = (pkd_type_par & 0xfff0) >> 4
            used = ((pkd_used & 0xfc) >> 2) & 1
        elif bv >= 40:
            fmtcode, pkd_type_par, pkd_used = unpack('<xBHxB', data[0:6])
            is_style = (pkd_type_par & 4) == 4
            parent = (pkd_type_par & 0xfff0) >> 4
            used = ((pkd_used & 0xfc) >> 2) & 1
        elif bv == 30:
            fmtcode, pkd_type, pkd_used, pkd_par = unpack('<xBBBH', data[0:6])
            is_style = (pkd_type & 4) == 4
            parent = (pkd_par & 0xfff0) >> 4
            used = ((pkd_used & 0xfc) >> 2) & 1
        else:
            raise XLRDError('programmer stuff-up: bv=%d' % bv)
        if DEBUG: fprintf(self.logfile, "XF record: %d code: 0x%04x (%d) sty=%d par=%d used=%d\n", \
            self.xfcount, fmtcode, fmtcode, is_style, parent, used)
        if is_style:
            if used: # misnomer; bit set means "ignore attribute"
                xsfn = -1
                myfn = -1
            else:
                xsfn = fmtcode
                myfn = fmtcode
        else:
            xsfn = -1
            if used:
                myfn = fmtcode
            else:
                if self.xf_style_fmt_no >= 0:
                    myfn = self.xf_style_fmt_no[parent]
                else:
                    myfn = fmtcode
        if DEBUG: fprintf(self.logfile, "XF record: %d; style code %d, own code %d\n", \
            self.xfcount, xsfn, myfn)
        self.xf_style_fmt_no.append(xsfn)
        # if bv < 50:
        #     xfrec = self.format_list[fmtcode]
        # elif fmtcode not in self.format_dict:
        if myfn not in self.format_dict:
            if myfn != -1:
                print >> self.logfile, "*** XF(%d): Unknown format code 0x%04x (%d)" % (self.xfcount, myfn, myfn)
            ty = std_format_code_types.get(myfn, FUN)
            xfrec = Format(myfn, ty, u'')
            self.format_dict[myfn] = xfrec
        else:
            xfrec = self.format_dict[myfn]
        self.xfrecords.append(xfrec)
        self.xfcount += 1

    def handle_sheethdr(self, data):
        # This a BIFF 4W special.
        # The SHEETHDR record is followed by a (BOF ... EOF) substream containing
        # a worksheet.
        # DEBUG = 0
        sheet_len = unpack('<i', data[:4])[0]
        sheet_name = unpack_string(data, 4, self.encoding, lenlen=1)
        sheetno = self._sheethdr_count
        assert sheet_name == self._sheet_names[sheetno]
        self._sheethdr_count += 1
        BOF_posn = self._position
        posn = BOF_posn - 4 - len(data)
        if DEBUG: print >> self.logfile, 'SHEETHDR %d at posn %d: len=%d name=%r' % (sheetno, posn, sheet_len, sheet_name)
        self.initialise_format_info()
        sht = self.get_sheet()
        if DEBUG: print >> self.logfile, 'SHEETHDR: posn after get_sheet() =', self._position
        self.position(BOF_posn + sheet_len)
        sht.name = self._sheet_names[sheetno]
        self._sheet.append(sht)

    def handle_sheetsoffset(self, data):
        # DEBUG = 0
        posn = unpack('<i', data)[0]
        if DEBUG: print >> self.logfile, 'SHEETSOFFSET:', posn
        self._sheetsoffset = posn

    def handle_sst(self, data):
        if DEBUG: print >> self.logfile, "SST Processing"
        nbt = len(data)
        strlist = [data]
        uniquestrings = unpack('<i', data[4:8])[0]
        if DEBUG or self.verbosity >= 2: fprintf(self.logfile, "SST: unique strings: %d\n", uniquestrings)
        while 1:
            code, nb, data = self.get_record_parts_conditional(XL_CONTINUE)
            if code is None:
                break
            nbt += nb
            if DEBUG: fprintf(self.logfile, "CONTINUE: adding %d bytes to SST -> %d\n", nb, nbt)
            # if DEBUG: print "first 30", repr(data[:30])
            # if DEBUG: print " last 30", repr(data[-30:])
            strlist.append(data)
        pos = 8
        dinx = 0
        strings = []
        for _unused_i in xrange(uniquestrings):
            strg, newdinx, newpos = unpack_unicode_table(strlist, dinx, pos)
            pos = newpos
            dinx = newdinx
            strings.append(strg)
        self._sharedstrings = strings

    def handle_writeaccess(self, data):
        # DEBUG = 0
        if self.biff_version < 80:
            if self.encoding == "unknown":
                self.raw_user_name = True
                self.user_name = data
                return
            strg = unpack_string(data, 0, self.encoding, lenlen=1)
        else:
            strg = unpack_unicode(data, 0, lenlen=2)
        if DEBUG: print >> self.logfile, "WRITEACCESS: %d bytes; raw=%d %r" % (len(data), self.raw_user_name, strg)
        strg = strg.rstrip()
        self.user_name = strg

    def is_date_format_string(self, fmt):
        # Heuristics:
        # Ignore "text" and [stuff in square brackets (aarrgghh -- see below)].
        # Handle backslashed-escaped chars properly.
        # E.g. hh\hmm\mss\s should produce a display like 23h59m59s
        # Date formats have one or more of ymdhs (caseless) in them.
        # Numeric formats have # and 0.
        # N.B. u'General"."' hence get rid of "text" first.
        # ### TODO ### Find where formats are interpreted in Gnumeric
        # ### TODO ### u'[h]\\ \\h\\o\\u\\r\\s' where [h] means don't care about hours > 23
        state = 0
        s = ''
        for c in fmt:
            if state == 0:
                if c == u'"':
                    state = 1
                else:
                    s += c
            else:
                if c == u'"':
                    state = 0
        if s in non_date_formats:
            return False
        state = 0
        date_count = num_count = 0
        for c in s:
            if state == 0:
                if c == u'[':
                    state = 2
                elif c == u'\\':
                    state = 3
                elif c in date_char_dict:
                    date_count += date_char_dict[c]
                elif c in num_char_dict:
                    num_count += num_char_dict[c]
            elif state == 2:
                if c == u']':
                    state = 0
            elif state == 3:
                # ignore the escaped character
                state = 0
        if state != 0:
            print >> self.logfile, '*** is_date_format: parse failure: state=%d; s=%r' % (state, s)
        if date_count and not num_count:
            return True
        if num_count and not date_count:
            return False
        if date_count:
            print >> self.logfile, '*** is_date_format: ambiguous d=%d n=%d s=%r' % (date_count, num_count, s)
        else:
            print >> self.logfile, '*** is_date_format: no signif. format codes? s=%r' % s
        return date_count > num_count

    def parse_globals(self):
        # DEBUG = 0
        # self.position(self._own_bof) # no need to position, just start reading (after the BOF)
        while 1:
            rc, length, data = self.get_record_parts()
            if DEBUG: print "parse_globals: record code is 0x%04x" % rc
            if rc == XL_SST:
                self.handle_sst(data)
            elif rc == XL_FORMAT: # XL_FORMAT2 is BIFF <= 3.0, can't appear in globals
                self.handle_format(data)
            elif rc == XL_XF:
                self.handle_xf(data)
            elif rc ==  XL_BOUNDSHEET:
                self.handle_boundsheet(data)
            elif rc == XL_DATEMODE:
                self.handle_datemode(data)
            elif rc == XL_CODEPAGE:
                self.handle_codepage(data)
            elif rc == XL_COUNTRY:
                self.handle_country(data)
            elif rc == XL_FILEPASS:
                self.handle_filepass(data)
            elif rc == XL_WRITEACCESS:
                self.handle_writeaccess(data)
            elif rc == XL_SHEETSOFFSET:
                self.handle_sheetsoffset(data)
            elif rc == XL_SHEETHDR:
                self.handle_sheethdr(data)
            elif rc & 0xff == 9:
                print >> self.logfile, "*** Unexpected BOF at posn %d: 0x%04x len=%d data=%r" \
                    % (self._position - length - 4, rc, length, data)
            elif rc ==  XL_EOF:
                if self.biff_version == 45:
                    # DEBUG = 0
                    if DEBUG: print "global EOF: position", self._position
                    # if DEBUG:
                    #     pos = self._position - 4
                    #     print repr(self.mem[pos:pos+40])
                return
            else:
                # if DEBUG:
                #     print "parse_globals: ignoring record code 0x%04x" % rc
                pass

    def position(self, pos):
        self._position = pos

    def read(self, pos, length):
        data = self.mem[pos:pos+length]
        self._position = pos + len(data)
        return data

    def getbof(self, rqd_stream):
        # DEBUG = 0
        if DEBUG: print >> self.logfile, "getbof(): position", self._position
        savpos = self._position
        opcode = self.get2bytes()
        if opcode == MY_EOF: raise XLRDError('Expected BOF record; met end of file')
        if opcode not in boflen: raise XLRDError('Expected BOF record; found 0x%04x' % opcode)
        length = self.get2bytes()
        if length == MY_EOF: raise XLRDError('Incomplete BOF record[1]; met end of file')
        if length < boflen[opcode] or length > 20:
            raise XLRDError('Invalid length (%d) for BOF record type 0x%04x' % (length, opcode))
        data = self.read(self._position, length);
        if DEBUG: print >> self.logfile, "\ngetbof(): data=%r" % data
        if len(data) < length: raise XLRDError('Incomplete BOF record[2]; met end of file')
        version1 = opcode >> 8
        version2, streamtype = unpack('<HH', data[0:4])
        if DEBUG: print >> self.logfile, "getbof(): op=0x%04x version2=0x%04x streamtype=0x%04x" \
            % (opcode, version2, streamtype)
        bof_offset = self._position - 4 - length
        if DEBUG: print >> self.logfile, "getbof(): BOF found at offset %d; savpos=%d" % (bof_offset, savpos)
        version = build = year = 0
        if version1 == 0x08:
            build, year = unpack('<HH', data[4:8])
            if version2 == 0x0600:
                version = 80
            elif version2 == 0x0500:
                if year < 1994 or build in (2412, 3218, 3321):
                    version = 50
                else:
                    version = 70
            else:
                # dodgy one, created by a 3rd-party tool
                version = {
                    0x0000: 2,
                    0x0007: 2,
                    0x0200: 2,
                    0x0300: 3,
                    0x0400: 4,
                    }.get(version2, 0) * 10
        elif version1 in (0x04, 0x02, 0x00):
            version = (version1 // 2 + 2) * 10 #  i.e. 2, 3, or 4
        if version == 40 and streamtype == XL_WORKBOOK_GLOBALS_4W:
            version += 5 # i.e. 4W

        if DEBUG or self.verbosity >= 2:
            print >> self.logfile, "BOF: op=0x%04x vers=0x%04x stream=0x%04x buildid=%d buildyr=%d -> BIFF%d" \
                % (opcode, version2, streamtype, build, year, version)
        got_globals = streamtype == XL_WORKBOOK_GLOBALS or (version == 45 and streamtype == XL_WORKBOOK_GLOBALS_4W)
        if (rqd_stream == XL_WORKBOOK_GLOBALS and got_globals) or streamtype == rqd_stream:
            return version
        if version < 50 and streamtype == XL_WORKSHEET:
            return version
        raise XLRDError(
            'BOF not workbook/worksheet: op=0x%04x vers=0x%04x strm=0x%04x build=%d year=%d -> BIFF%d' \
            % (opcode, version2, streamtype, build, year, version)
            )

# === helper functions

def unpack_unicode_table(datatab, datainx, pos, lenlen=2):
    "Return (unicode_strg, updated_inx, updated_pos)"
    # DEBUG = 0
    data = datatab[datainx]
    datalen = len(data)
    nchars = unpack('<' + 'BH'[lenlen-1], data[pos:pos+lenlen])[0]
    pos += lenlen
    options = ord(data[pos])
    pos += 1
    phonetic = options & 0x04
    richtext = options & 0x08
    if richtext:
        rt = unpack('<H', data[pos:pos+2])[0]
        pos += 2
    if phonetic:
        sz = unpack('<i', data[pos:pos+4])[0]
        pos += 4
    accstrg = u''
    charsgot = 0
    while 1:
        charsneed = nchars - charsgot
        if options & 0x01:
            # Uncompressed UTF-16
            charsavail = min((datalen - pos) // 2, charsneed)
            rawstrg = data[pos:pos+2*charsavail]
            # if DEBUG: print "SST U16: nchars=%d pos=%d rawstrg=%r" % (nchars, pos, rawstrg)
            accstrg += rawstrg.decode('utf-16le')
            pos += 2*charsavail
        else:
            # Note: this is COMPRESSED (not ASCII!) encoding!!!
            charsavail = min(datalen - pos, charsneed)
            rawstrg = data[pos:pos+charsavail]
            # if DEBUG: print "SST CMPRSD: nchars=%d pos=%d rawstrg=%r" % (nchars, pos, rawstrg)
            accstrg += ''.join([unichr(ord(x)) for x in rawstrg])
            pos += charsavail
        charsgot += charsavail
        if charsgot == nchars:
            break
        datainx += 1
        data = datatab[datainx]
        datalen = len(data)
        options = ord(data[0])
        pos = 1
    if richtext:
        pos += 4 * rt
    if phonetic:
        pos += sz
    # also allow for the rich text etc being split ...
    if pos >= datalen:
        # adjust to correct position in next record
        pos = pos - datalen
        datainx += 1
    return (accstrg, datainx, pos)

# === formatting stuff ===

# Currently the format is used only in trying to tell which cells are dates.

class Format(object):
    def __init__(self, xf, ty, format_str):
        self.code = xf
        self.type = ty
        self.format_str = format_str

fmt_code_ranges = [ # both-inclusive ranges of "standard" format codes
    # Source: the openoffice.org doc't
    ( 0,  0, FGE),
    ( 1, 13, FNU),
    (14, 22, FDT),
    (27, 36, FDT), # Japanese
    (37, 44, FNU),
    (45, 47, FDT),
    (48, 48, FNU),
    (49, 49, FTX),
    (50, 58, FDT), # Japanese
    ]

std_format_code_types = {}
for lo, hi, ty in fmt_code_ranges:
    for x in xrange(lo, hi+1):
        std_format_code_types[x] = ty
del lo, hi, ty, x

date_chars = u'ymdhs' # year, month/minute, day, hour, second
date_char_dict = {}
for _c in date_chars + date_chars.upper():
    date_char_dict[_c] = 5
del _c, date_chars

num_char_dict = {
    u'0': 5,
    u'#': 5,
    u'?': 5,
    u';': 1,
    }

non_date_formats = {
    u'0.00E+00':1,
    u'##0.0E+0':1,
    u'General' :1,
    u'@'       :1,
    }

# Boolean format strings (actual cases)
# u'"Yes";"Yes";"No"'
# u'"True";"True";"False"'
# u'"On";"On";"Off"'

# ===================================================================================
