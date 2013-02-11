# __init__.py
# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

import inspect
import sys

import pptx.exc as exceptions
sys.modules['pptx.exceptions'] = exceptions

from pptx.api import(Presentation)

# __all__ = sorted(name for name, obj in locals().items() if not (name.startswith('_') or inspect.ismodule(obj)))
                 
__version__ = '0.2.0'

del inspect, sys

