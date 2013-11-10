# encoding: utf-8

"""
Enumerations used in python-pptx.
"""

from __future__ import absolute_import


class Enumeration(object):
    pass


class MSO_COLOR_TYPE(Enumeration):
    """
    Corresponds to MsoColorType
    http://msdn.microsoft.com/en-us/library/office/aa432491(v=office.12).aspx
    """
    RGB = 1
    SCHEME = 2
