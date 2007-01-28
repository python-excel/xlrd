if __name__ == "__main__":

    PSYCO = 1

    import xlrd
    import sys, time, glob
    import string

    null_cell = xlrd.empty_cell

    def colname(x):
        sau = string.ascii_uppercase
        if x <= 25:
            return sau[x]
        else:
            return sau[x // 26 - 1] + sau[x % 26]

    def show_row(bk, sh, rowx, colrange, printit):
        if printit: print
        for colx, ty, val in get_row_data(bk, sh, rowx, colrange):
            if printit:
                print "cell %s%d: type=%d, data: %r" % (colname(colx), rowx+1, ty, val)

    def get_row_data(bk, sh, rowx, colrange):
        result = []
        dmode = bk.datemode
        ctys = sh.row_types(rowx)
        cvals = sh.row_values(rowx)
        for colx in colrange:
            cty = ctys[colx]
            cval = cvals[colx]
            if cty == xlrd.XL_CELL_DATE:
                try:
                    showval = xlrd.xldate_as_tuple(cval, dmode)
                except xlrd.XLDateError:
                    ei = sys.exc_info()[:2]
                    showval = "%s:%s" % ei
                    cty = xlrd.XL_CELL_ERROR
            elif cty == xlrd.XL_CELL_ERROR:
                showval = xlrd.error_text_from_code.get(cval, '<Unknown error code 0x%02x>' % cval)
            else:
                showval = cval
            result.append((colx, cty, showval))
        return result

    def show(bk, nshow=65535, printit=True):
        print
        print "BIFF version: %s, datemode: %s, codepage: %d (encoding: %s), countries: %r" \
            % (xlrd.biff_text_from_num[bk.biff_version],
            bk.datemode, bk.codepage, bk.encoding, bk.countries)
        print "nsheets: %d; sheet names: %r" % (bk.nsheets, bk.sheet_names())
        print
        for shx in range(bk.nsheets):
            sh = bk.sheet_by_index(shx)
            nrows, ncols = sh.nrows, sh.ncols
            colrange = range(ncols)
            anshow = min(nshow, nrows)
            print "sheet %d: name = %r, nrows = %d, ncols = %d" % \
                (shx, sh.name, sh.nrows, sh.ncols)
            for rowx in xrange(anshow-1):
                if not printit and rowx % 10000 == 1 and rowx > 1:
                    print "done %d rows" % (rowx-1,)
                show_row(bk, sh, rowx, colrange, printit)
            if anshow and nrows:
                show_row(bk, sh, nrows-1, colrange, printit)
            print

    def main():
        import optparse
        usage = "%prog [options] command input-file-patterns"
        oparser = optparse.OptionParser(usage)
        oparser.add_option(
            "-l", "--logfilename",
            default="",
            help="contains error messages")
        oparser.add_option(
           "-v", "--verbosity",
           type="int", default=0,
           help="level of information and diagnostics provided")
        options, args = oparser.parse_args()
        if len(args) != 2:
            oparser.error("Expected 2 args, found %d" % len(args))

        if PSYCO:
            try:
                import psyco
                psyco.log()
                psyco.profile()
            except ImportError:
                pass

        cmd = args[0]
        if cmd == 'dump':
            xlrd.dump(args[1])
            sys.exit(0)
        if options.logfilename:
            logfile = file(options.logfilename, 'w')
        else:
            logfile = sys.stdout
        for pattern in args[1:]:
            for fname in glob.glob(pattern):
                print >> logfile, "\n=== File: %s ===" % fname
                try:
                    t0 = time.time()
                    bk = xlrd.open_workbook(fname, verbosity=options.verbosity, logfile=logfile)
                    t1 = time.time()
                    print >> logfile, "Open took %.2f seconds" % (t1-t0,)
                except xlrd.XLRDError:
                    print >> logfile, "*** Open failed: %s: %s" % sys.exc_info()[:2]
                    continue
                t0 = time.time()
                if cmd == 'ov': # OverView
                    show(bk, 0)
                elif cmd == 'show': # all rows
                    show(bk)
                elif cmd == '2rows': # first row and last row
                    show(bk, 2)
                elif cmd == '3rows': # first row, 2nd row and last row
                    show(bk, 3)
                elif cmd == 'bench':
                    show(bk, printit=False)
                else:
                    print >> logfile, "*** Unknown command <%s>" % cmd
                    sys.exit(1)
                t1 = time.time()
                print >> logfile, "\ncommand took %.2f seconds\n" % (t1-t0,)
                
    main()
