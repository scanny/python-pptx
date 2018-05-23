# encoding: utf-8

"""
Provides Python 2/3 compatibility objects
"""

from __future__ import (
    absolute_import, division, print_function, unicode_literals
)

import sys

if sys.version_info >= (3, 0):
    from .python3 import (  # noqa
        BytesIO, is_integer, is_string, is_unicode, to_unicode, Unicode
    )
else:
    from .python2 import (  # noqa
        BytesIO, is_integer, is_string, is_unicode, to_unicode, Unicode
    )
