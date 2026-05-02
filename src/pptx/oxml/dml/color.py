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

    def set_alpha(self, value):
        """Set the `<a:alpha>` child to `value`, a float in `[0.0, 1.0]`.

        Returns the newly set `<a:alpha>` child element. Any existing
        `<a:alpha>` is replaced.
        """
        self._remove_alpha()
        alpha = self._add_alpha()
        alpha.val = value
        return alpha

    def clear_alpha(self):
        """Remove any `<a:alpha>` child, returning self."""
        self._remove_alpha()
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
    """Custom element class for `<a:alpha>` and similar elements.

    The `val` attribute is an `ST_PositiveFixedPercentage` (0%-100%),
    represented in Python as a float in `[0.0, 1.0]`.
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
