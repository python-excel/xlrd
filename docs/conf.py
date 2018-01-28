import datetime
import os

import pkginfo

on_rtd = os.environ.get('READTHEDOCS', None) == 'True'
pkg_info = pkginfo.Develop(os.path.join(os.path.dirname(__file__), os.pardir))

intersphinx_mapping = {'http://docs.python.org': None}
extensions = ['sphinx.ext.autodoc', 'sphinx.ext.intersphinx']
source_suffix = '.rst'
master_doc = 'index'
project = u'xlrd'
copyright = '2005-%s Stephen John Machin, Lingfo Pty Ltd' % datetime.datetime.now().year
version = release = pkg_info.version
exclude_patterns = ['_build']
pygments_style = 'sphinx'

if on_rtd:
    html_theme = 'default'
else:
    html_theme = 'classic'

htmlhelp_basename = project+'doc'
intersphinx_mapping = {'python': ('http://docs.python.org', None)}

autodoc_member_order = 'bysource'
