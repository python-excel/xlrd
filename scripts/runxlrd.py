if __name__ == "__main__":

    PSYCO = 0

    import xlrd
    import sys, time, glob

    null_cell = xlrd.empty_cell

    def colname(x):
        sau = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        if x <= 25:
            return sau[x]
        else:
            xdiv26, xmod26 = divmod(x, 26)
            return sau[xdiv26 - 1] + sau[xmod26]

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

    def show(bk, nshow=65535, printit=1):
        print
        print "BIFF version: %s; datemode: %s" \
            % (xlrd.biff_text_from_num[bk.biff_version], bk.datemode)
        print "codepage: %d (encoding: %s); countries: %r" \
            % (bk.codepage, bk.encoding, bk.countries)
        print "last saved by: %r" % bk.user_name
        print "nsheets: %d; sheet names: %r" % (bk.nsheets, bk.sheet_names())
        print "Pickleable: %d; Use mmap: %d" \
            % (bk.pickleable, bk.use_mmap)
        print "Load time: %.2f seconds (stage 1) %.2f seconds (stage 2)" \
            % (bk.load_time_stage_1, bk.load_time_stage_2)

        if 0:
            rclist = xlrd.sheet.rc_stats.items()
            rclist.sort()
            print "rc stats"
            for k, v in rclist:
                print "0x%04x %7d" % (k, v)
        print
        for shx in range(bk.nsheets):
            sh = bk.sheet_by_index(shx)
            nrows, ncols = sh.nrows, sh.ncols
            colrange = range(ncols)
            anshow = min(nshow, nrows)
            print "sheet %d: name = %r; nrows = %d; ncols = %d" % \
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
        oparser.add_option(
           "-p", "--pickleable",
           type="int", default=1,
           help="1: ensure Book object is pickleable (default); 0: don't bother")
        oparser.add_option(
           "-m", "--mmap",
           type="int", default=-1,
           help="1: use mmap; 0: don't use mmap; -1: accept heuristic")


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
        xlrd_version = getattr(xlrd, "__VERSION__", "unknown; before 0.5")
        if cmd == 'dump':
            xlrd.dump(args[1])
            sys.exit(0)
        if cmd == 'version':
            print "xlrd version:", xlrd_version
            sys.exit(0)
        if options.logfilename:
            logfile = open(options.logfilename, 'w')
        else:
            logfile = sys.stdout
        mmap_opt = options.mmap
        mmap_arg = xlrd.USE_MMAP
        if mmap_opt in (1, 0):
            mmap_arg = mmap_opt
        elif mmap_opt != -1:
            print 'Unexpected value (%r) for mmap option -- assuming default' % mmap_opt
        for pattern in args[1:]:
            for fname in glob.glob(pattern):
                print >> logfile, "\n=== File: %s ===" % fname
                try:
                    t0 = time.time()
                    bk = xlrd.open_workbook(fname,
                        verbosity=options.verbosity, logfile=logfile,
                        pickleable=options.pickleable, use_mmap=mmap_arg)
                    t1 = time.time()
                    print >> logfile, "Open took %.2f seconds" % (t1-t0,)
                except xlrd.XLRDError:
                    print >> logfile, "*** Open failed: %s: %s" % sys.exc_info()[:2]
                    continue
                except:
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
