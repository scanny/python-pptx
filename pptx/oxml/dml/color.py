# encoding: utf-8

"""
lxml custom element classes for DrawingML-related XML elements.
"""

from __future__ import absolute_import

from ...enum.dml import MSO_THEME_COLOR
from ..xmlchemy import ZeroOrOneChoice, Choice
from ..simpletypes import ST_HexColorRGB, ST_Percentage, ST_AlphaValue
from ..xmlchemy import BaseOxmlElement, RequiredAttribute, ZeroOrOne


class _BaseColorElement(BaseOxmlElement):
    """
    Base class for <a:srgbClr> and <a:schemeClr> elements.
    """
    lumMod = ZeroOrOne('a:lumMod')
    lumOff = ZeroOrOne('a:lumOff')

    eg_alphaChoice = ZeroOrOneChoice(
        (Choice('a:alpha'),),
        successors=('a:alpha',)
    )

    def add_lumMod(self, value):
        """
        Return a newly added <a:lumMod> child element.
        """
        lumMod = self._add_lumMod()
        lumMod.val = value
        return lumMod

    def add_lumOff(self, value):
        """
        Return a newly added <a:lumOff> child element.
        """
        lumOff = self._add_lumOff()
        lumOff.val = value
        return lumOff

    def clear_lum(self):
        """
        Return self after removing any <a:lumMod> and <a:lumOff> child
        elements.
        """
        self._remove_lumMod()
        self._remove_lumOff()
        return self


class CT_HslColor(_BaseColorElement):
    """
    Custom element class for <a:hslClr> element.
    """


class CT_Percentage(BaseOxmlElement):
    """
    Custom element class for <a:lumMod> and <a:lumOff> elements.
    """
    val = RequiredAttribute('val', ST_Percentage)


class CT_PresetColor(_BaseColorElement):
    """
    Custom element class for <a:prstClr> element.
    """


class CT_SchemeColor(_BaseColorElement):
    """
    Custom element class for <a:schemeClr> element.
    """
    val = RequiredAttribute('val', MSO_THEME_COLOR)


class CT_ScRgbColor(_BaseColorElement):
    """
    Custom element class for <a:scrgbClr> element.
    """


class CT_SRgbColor(_BaseColorElement):
    """
    Custom element class for <a:srgbClr> element.
    """
    val = RequiredAttribute('val', ST_HexColorRGB)


class CT_SystemColor(_BaseColorElement):
    """
    Custom element class for <a:sysClr> element.
    """


class CT_Alpha(BaseOxmlElement):
    """
    Custom element class for <a:alpha> element.
    """
    val = RequiredAttribute('val', ST_AlphaValue)
