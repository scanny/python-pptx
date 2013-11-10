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


class MSO_THEME_COLOR(Enumeration):
    """
    Corresponds to MsoColorType
    http://msdn.microsoft.com/en-us/library/office/aa432491(v=office.12).aspx
    """
    NOT_THEME_COLOR = 0
    ACCENT_1 = 5
    ACCENT_2 = 6
    ACCENT_3 = 7
    ACCENT_4 = 8
    ACCENT_5 = 9
    ACCENT_6 = 10
    BACKGROUND_1 = 14
    BACKGROUND_2 = 16
    DARK_1 = 1
    DARK_2 = 3
    FOLLOWED_HYPERLINK = 12
    HYPERLINK = 11
    LIGHT_1 = 2
    LIGHT_2 = 4
    MIXED = -2
    TEXT_1 = 13
    TEXT_2 = 15

    _xml_to_val = {
        'bg1':      BACKGROUND_1,
        'tx1':      TEXT_1,
        'bg2':      BACKGROUND_2,
        'tx2':      TEXT_2,
        'accent1':  ACCENT_1,
        'accent2':  ACCENT_2,
        'accent3':  ACCENT_3,
        'accent4':  ACCENT_4,
        'accent5':  ACCENT_5,
        'accent6':  ACCENT_6,
        'hlink':    HYPERLINK,
        'folHlink': FOLLOWED_HYPERLINK,
        # 'phClr':    None?,
        'dk1':      DARK_1,
        'lt1':      LIGHT_1,
        'dk2':      DARK_2,
        'lt2':      LIGHT_2,
    }

    @classmethod
    def from_xml(cls, st_schemecolorval):
        return cls._xml_to_val[st_schemecolorval]
