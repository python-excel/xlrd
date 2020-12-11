import datetime
import os

from xlrd.info import __VERSION__

on_rtd = os.environ.get('READTHEDOCS', None) == 'True'

intersphinx_mapping = {'http://docs.python.org': None}
extensions = ['sphinx.ext.autodoc', 'sphinx.ext.intersphinx']
source_suffix = '.rst'
master_doc = 'index'
project = u'xlrd'
copyright = '2005-%s Stephen John Machin, Lingfo Pty Ltd' % datetime.datetime.now().year
version = release = __VERSION__
exclude_patterns = ['_build']
pygments_style = 'sphinx'

if on_rtd:
    html_theme = 'default'
else:
    html_theme = 'classic'

htmlhelp_basename = project+'doc'
intersphinx_mapping = {'python': ('http://docs.python.org', None)}

autodoc_member_order = 'bysource'
