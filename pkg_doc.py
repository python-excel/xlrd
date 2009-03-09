from pythondoc import ET, parse, CompactHTML

MODULE_NAME = "xlrd"
PATH_TO_FILES = "C:/xlrd/svnco/trunk/xlrd"


module = ET.Element("module", name=MODULE_NAME)

parts = [
    '__init__',
    'sheet',
    'xldate',
    # 'compdoc',
    'biffh',
    'formatting',
    'formula',
    ]
flist = ["%s/%s.py" % (PATH_TO_FILES, p) for p in parts]
for fname in flist:
    print "about to parse", fname
    elem = parse(fname)
    for elem in elem:
        if module and elem.tag == "info":
            # skip all module info sections except the first
            continue
        module.append(elem)

formatter = CompactHTML()
print formatter.save(module, MODULE_NAME), "ok"

