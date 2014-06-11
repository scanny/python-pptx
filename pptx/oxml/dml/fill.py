# encoding: utf-8

"""
lxml custom element classes for DrawingML-related XML elements.
"""

from __future__ import absolute_import

from ..xmlchemy import BaseOxmlElement, Choice, ZeroOrOneChoice


class CT_BlipFillProperties(BaseOxmlElement):
    """
    Custom element class for <a:blipFill> element.
    """


class CT_GradientFillProperties(BaseOxmlElement):
    """
    Custom element class for <a:gradFill> element.
    """


class CT_GroupFillProperties(BaseOxmlElement):
    """
    Custom element class for <a:grpFill> element.
    """


class CT_NoFillProperties(BaseOxmlElement):
    """
    Custom element class for <a:NoFill> element.
    """


class CT_PatternFillProperties(BaseOxmlElement):
    """
    Custom element class for <a:pattFill> element.
    """


class CT_SolidColorFillProperties(BaseOxmlElement):
    """
    Custom element class for <a:solidFill> element.
    """
    eg_colorChoice = ZeroOrOneChoice((
        Choice('a:scrgbClr'), Choice('a:srgbClr'), Choice('a:hslClr'),
        Choice('a:sysClr'), Choice('a:schemeClr'), Choice('a:prstClr')),
        successors=()
    )
