"""lxml custom element classes for DrawingML-related XML elements."""

from __future__ import annotations

from pptx.enum.dml import MSO_THEME_COLOR
from pptx.oxml.simpletypes import ST_HexColorRGB, ST_Percentage, ST_PositiveFixedPercentage
from pptx.oxml.xmlchemy import (
    BaseOxmlElement,
    Choice,
    RequiredAttribute,
    ZeroOrOne,
    ZeroOrOneChoice,
)


class _BaseColorElement(BaseOxmlElement):
    """
    Base class for <a:srgbClr> and <a:schemeClr> elements.
    """

    lumMod = ZeroOrOne("a:lumMod")
    lumOff = ZeroOrOne("a:lumOff")
    alpha = ZeroOrOne("a:alpha")
    alphaOff = ZeroOrOne("a:alphaOff")
    alphaMod = ZeroOrOne("a:alphaMod")

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

    def add_alpha(self, value):
        """
        Return a newly added <a:alpha> child element.
        """
        alpha = self._add_alpha()
        alpha.val = value
        return alpha

    def add_alphaOff(self, value):
        """
        Return a newly added <a:alphaOff> child element.
        """
        alphaOff = self._add_alphaOff()
        alphaOff.val = value
        return alphaOff

    def add_alphaMod(self, value):
        """
        Return a newly added <a:alphaMod> child element.
        """
        alphaMod = self._add_alphaMod()
        alphaMod.val = value
        return alphaMod

    def clear_lum(self):
        """
        Return self after removing any <a:lumMod> and <a:lumOff> child
        elements.
        """
        self._remove_lumMod()
        self._remove_lumOff()
        return self

    def clear_alpha(self):
        """
        Return self after removing any <a:alpha>, <a:alphaOff> and <a:alphaMod> child
        elements.
        """
        self._remove_alpha()
        self._remove_alphaOff()
        self._remove_alphaMod()
        return self


class CT_Color(BaseOxmlElement):
    """Custom element class for `a:fgClr`, `a:bgClr` and perhaps others."""

    eg_colorChoice = ZeroOrOneChoice(
        (
            Choice("a:scrgbClr"),
            Choice("a:srgbClr"),
            Choice("a:hslClr"),
            Choice("a:sysClr"),
            Choice("a:schemeClr"),
            Choice("a:prstClr"),
        ),
        successors=(),
    )


class CT_HslColor(_BaseColorElement):
    """
    Custom element class for <a:hslClr> element.
    """


class CT_Percentage(BaseOxmlElement):
    """
    Custom element class for <a:lumMod> and <a:lumOff> elements.
    """

    val = RequiredAttribute("val", ST_Percentage)


class CT_PositiveFixedPercentage(BaseOxmlElement):
    """
    Custom element class for <a:alpha>, <a:alphaOff> and <a:alphaMod> elements.
    """

    val = RequiredAttribute("val", ST_PositiveFixedPercentage)


class CT_PresetColor(_BaseColorElement):
    """
    Custom element class for <a:prstClr> element.
    """


class CT_SchemeColor(_BaseColorElement):
    """
    Custom element class for <a:schemeClr> element.
    """

    val = RequiredAttribute("val", MSO_THEME_COLOR)


class CT_ScRgbColor(_BaseColorElement):
    """
    Custom element class for <a:scrgbClr> element.
    """


class CT_SRgbColor(_BaseColorElement):
    """
    Custom element class for <a:srgbClr> element.
    """

    val = RequiredAttribute("val", ST_HexColorRGB)


class CT_SystemColor(_BaseColorElement):
    """
    Custom element class for <a:sysClr> element.
    """
