# encoding: utf-8

"""
Initialization module for python-pptx
"""


import pptx.exc as exceptions
import sys
sys.modules['pptx.exceptions'] = exceptions
del sys

from pptx.api import Presentation  # noqa

__version__ = '0.2.6'
