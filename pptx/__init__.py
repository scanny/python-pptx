# encoding: utf-8

"""
Initialization module for python-pptx
"""

import sys

import pptx.exc as exceptions
sys.modules['pptx.exceptions'] = exceptions

from pptx.api import Presentation  # noqa

__version__ = '0.2.6'

import logging
log = logging.getLogger('pptx')
log.setLevel(logging.DEBUG)
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)
formatter = logging.Formatter(
    '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
ch.setFormatter(formatter)
log.addHandler(ch)

del sys
