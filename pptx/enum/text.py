# encoding: utf-8

"""
Enumerations used by text and related objects
"""

from __future__ import absolute_import

from . import Enumeration_OLD


class MSO_AUTO_SIZE(Enumeration_OLD):
    """
    Corresponds to MsoAutoSize enumeration
    http://msdn.microsoft.com/en-us/library/office/ff865367(v=office.15).aspx
    """
    NONE = 0
    SHAPE_TO_FIT_TEXT = 1
    TEXT_TO_FIT_SHAPE = 2
