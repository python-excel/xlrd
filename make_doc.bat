cd \xlrd\svnco-trunk-May2009
python pkg_doc.py c:/xlrd/svnco-trunk-May2009/xlrd
python pythondoc.py xlrd\compdoc.py
del compdoc.html
rename pythondoc-compdoc.html compdoc.html
copy xlrd.html xlrd\doc
copy compdoc.html xlrd\doc
