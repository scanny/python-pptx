"""lxml custom element classes for theme-related XML elements."""

from __future__ import annotations

from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

from . import parse_from_template
from .xmlchemy import BaseOxmlElement


class CT_OfficeStyleSheet(BaseOxmlElement):
    """
    ``<a:theme>`` element, root of a theme part.
    """

    _tag_seq = (
        "a:themeElements",
        "a:objectDefaults",
        "a:extraClrSchemeLst",
        "a:custClrLst",
        "a:extLst",
    )
    del _tag_seq

    @classmethod
    def new_default(cls):
        """
        Return a new ``<a:theme>`` element containing default settings
        suitable for use with a notes master.
        """
        return parse_from_template("theme")

    # -- hyperlink color-scheme accessors (issue #940) -----------------

    @property
    def hlink_color(self) -> RGBColor | None:
        """RGB color in `a:clrScheme/a:hlink` (theme's unvisited-hyperlink color).

        Returns |None| when no `a:srgbClr` is present under `a:hlink` (e.g. the
        theme stores the color as an `a:sysClr` or the element is missing
        altogether).
        """
        return self._get_scheme_srgb("a:hlink")

    @hlink_color.setter
    def hlink_color(self, rgb: RGBColor):
        self._set_scheme_srgb("a:hlink", rgb)

    @property
    def folHlink_color(self) -> RGBColor | None:
        """RGB color in `a:clrScheme/a:folHlink` (theme's followed-hyperlink color).

        Returns |None| when no `a:srgbClr` is present under `a:folHlink`.
        """
        return self._get_scheme_srgb("a:folHlink")

    @folHlink_color.setter
    def folHlink_color(self, rgb: RGBColor):
        self._set_scheme_srgb("a:folHlink", rgb)

    # -- helpers -------------------------------------------------------

    def _get_scheme_srgb(self, nsptag: str) -> RGBColor | None:
        """Return an `RGBColor` for the `a:srgbClr` descendant of `a:clrScheme/<nsptag>`."""
        srgbClr = self.find(
            f".//{qn('a:themeElements')}/{qn('a:clrScheme')}/{qn(nsptag)}/{qn('a:srgbClr')}"
        )
        if srgbClr is None:
            return None
        val = srgbClr.get("val")
        if val is None:
            return None
        return RGBColor.from_string(val)

    def _set_scheme_srgb(self, nsptag: str, rgb: RGBColor) -> None:
        """Set `val` on `a:clrScheme/<nsptag>/a:srgbClr` to `rgb`, creating nodes as needed."""
        themeElements = self.find(qn("a:themeElements"))
        if themeElements is None:
            raise ValueError("theme has no `a:themeElements` child")
        clrScheme = themeElements.find(qn("a:clrScheme"))
        if clrScheme is None:
            raise ValueError("theme has no `a:themeElements/a:clrScheme` descendant")
        link = clrScheme.find(qn(nsptag))
        if link is None:
            raise ValueError(f"theme's color-scheme has no `{nsptag}` child")
        # -- remove any existing color child so the new `a:srgbClr` takes its place --
        for child in list(link):
            link.remove(child)
        from pptx.oxml.xmlchemy import OxmlElement

        srgbClr = OxmlElement("a:srgbClr")
        srgbClr.set("val", str(rgb))
        link.append(srgbClr)
