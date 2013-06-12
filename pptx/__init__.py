# __init__.py
# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

import sys

import pptx.exc as exceptions
sys.modules['pptx.exceptions'] = exceptions

from pptx.api import(Presentation)

__version__ = '0.2.5'

import logging
log = logging.getLogger('pptx')
log.setLevel(logging.DEBUG)
# log.setLevel(logging.INFO)
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s'
                              ' - %(message)s')
ch.setFormatter(formatter)
log.addHandler(ch)

del sys
