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
