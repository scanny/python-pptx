# encoding: utf-8

"""
Enumerations used in python-pptx.
"""

from __future__ import absolute_import


class Enumeration_OLD(object):

    @classmethod
    def from_xml(cls, xml_val):
        return cls._xml_to_idx[xml_val]

    @classmethod
    def to_xml(cls, enum_val):
        return cls._idx_to_xml[enum_val]


class MSO_COLOR_TYPE(Enumeration_OLD):
    """
    Corresponds to MsoColorType
    http://msdn.microsoft.com/en-us/library/office/aa432491(v=office.12).aspx
    """
    RGB = 1
    SCHEME = 2
    HSL = 101
    PRESET = 102
    SCRGB = 103
    SYSTEM = 104


class MSO_FILL_TYPE(Enumeration_OLD):
    """
    Corresponds to MsoFillType enumeration
    http://msdn.microsoft.com/EN-US/library/office/ff861408.aspx
    """
    BACKGROUND = 5
    GRADIENT = 3
    GROUP = 101
    PATTERNED = 2
    PICTURE = 6
    SOLID = 1
    TEXTURED = 4


class MSO_THEME_COLOR(Enumeration_OLD):
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

    _idx_to_xml = {
        ACCENT_1:           'accent1',
        ACCENT_2:           'accent2',
        ACCENT_3:           'accent3',
        ACCENT_4:           'accent4',
        ACCENT_5:           'accent5',
        ACCENT_6:           'accent6',
        BACKGROUND_1:       'bg1',
        BACKGROUND_2:       'bg2',
        DARK_1:             'dk1',
        DARK_2:             'dk2',
        FOLLOWED_HYPERLINK: 'folHlink',
        HYPERLINK:          'hlink',
        LIGHT_1:            'lt1',
        LIGHT_2:            'lt2',
        TEXT_1:             'tx1',
        TEXT_2:             'tx2',
    }

    _xml_to_idx = {
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
        return cls._xml_to_idx[st_schemecolorval]

    @classmethod
    def to_xml(cls, mso_theme_color_idx):
        return cls._idx_to_xml[mso_theme_color_idx]


class MSO_VERTICAL_ANCHOR(Enumeration_OLD):
    """
    Corresponds to MsoVerticalAnchor enumeration
    http://msdn.microsoft.com/en-us/library/office/ff865255.aspx
    """
    TOP = 1
    MIDDLE = 3
    BOTTOM = 4
    MIXED = -2

    _idx_to_xml = {
        None:   None,
        BOTTOM: 'b',
        MIDDLE: 'ctr',
        TOP:    't',
    }

    _xml_to_idx = {
        None:  None,
        'b':   BOTTOM,
        'ctr': MIDDLE,
        't':   TOP,
    }

MSO_ANCHOR = MSO_VERTICAL_ANCHOR
