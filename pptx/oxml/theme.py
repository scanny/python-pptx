# encoding: utf-8

"""
lxml custom element classes for theme-related XML elements.
"""

from __future__ import absolute_import, division, print_function, unicode_literals

from . import parse_from_template
from .xmlchemy import (
    BaseOxmlElement,
    RequiredAttribute,
    OptionalAttribute,
    OneAndOnlyOne,
    ZeroOrMoreChoice,
    Choice,
    OneOrMore,
)

from pptx.oxml.simpletypes import (
    XsdString
)
class CT_OfficeStyleSheet(BaseOxmlElement):
    """
    ``<a:theme>`` element, root of a theme part
    """

    _tag_seq = (
        "a:themeElements",
        "a:objectDefaults",
        "a:extraClrSchemeLst",
        "a:custClrLst",
        "a:extLst",
    )
    del _tag_seq

    theme_elements = OneAndOnlyOne("a:themeElements")

    name = OptionalAttribute('name', XsdString, default="")

    @classmethod
    def new_default(cls):
        """
        Return a new ``<a:theme>`` element containing default settings
        suitable for use with a notes master.
        """
        return parse_from_template("theme")



class CT_BaseStyles(BaseOxmlElement):

    _tag_seq = (
        "a:clrScheme",
        "a:fontScheme",
        "a:fmtScheme",
        "a:extList",
    )
    clrScheme = OneAndOnlyOne("a:clrScheme")
    fontScheme = OneAndOnlyOne("a:fontScheme")
    fmtScheme = OneAndOnlyOne("a:fmtScheme")
    

class CT_ColorScheme(BaseOxmlElement):
    
    _tag_seq = (
        "a:dk1",
        "a:lt1",
        "a:dk2",
        "a:lt2",
        "a:accent1",
        "a:accent2",
        "a:accent3",
        "a:accent4",
        "a:accent5",
        "a:accent6",
        "a:hlink",
        "a:folHlink",
        "a:extLst",
    )

    dk1 = OneAndOnlyOne("a:dk1")
    lt1 = OneAndOnlyOne("a:lt1")
    dk2 = OneAndOnlyOne("a:dk2")
    lt2 = OneAndOnlyOne("a:lt2")
    accent1 = OneAndOnlyOne("a:accent1")
    accent2 = OneAndOnlyOne("a:accent2")
    accent3 = OneAndOnlyOne("a:accent3")
    accent4 = OneAndOnlyOne("a:accent4")
    accent5 = OneAndOnlyOne("a:accent5")
    accent6 = OneAndOnlyOne("a:accent6")
    hlink = OneAndOnlyOne("a:hlink")
    folHlink = OneAndOnlyOne("a:folHlink")

    name = RequiredAttribute('name', XsdString)


class CT_FontScheme(BaseOxmlElement):
    _tag_seq = (
        "a:majorFont",
        "a:minorFont",
        "a_extList"
    )
    majorFont = OneAndOnlyOne("a:majorFont")
    minorFont = OneAndOnlyOne("a:minorFont")

    name = RequiredAttribute('name', XsdString)

    
class CT_StyleMatrix(BaseOxmlElement):
    _tag_seq = (
        "a:fillStyleLst",
        "a:lnStyleLst",
        "a:effectStyleLst",
        "a:bgFillStyleLst"
    )

    fillStyleLst = OneAndOnlyOne("a:fillStyleLst")
    lnStyleLst = OneAndOnlyOne("a:lnStyleLst")
    effectStyleLst = OneAndOnlyOne("a:effectStyleLst")
    bgFillStyleLst = OneAndOnlyOne("a:bgFillStyleLst")

    name = OptionalAttribute("name", XsdString, "")


class CT_FillStyleList(BaseOxmlElement):
    eg_fillProperties = ZeroOrMoreChoice(
        (
            Choice("a:noFill"),
            Choice("a:solidFill"),
            Choice("a:gradFill"),
            Choice("a:blipFill"),
            Choice("a:pattFill"),
            Choice("a:grpFill"),
        )
    )


class CT_LineStyleList(BaseOxmlElement):
    ln = ZeroOrMoreChoice(
        (Choice("a:ln"),)
    )

class CT_EffectStyleList(BaseOxmlElement):
    # Not Implementing for now
    pass

class CT_BackgroundFillStyleList(BaseOxmlElement):
    eg_fillProperties = ZeroOrMoreChoice(
        (
            Choice("a:noFill"),
            Choice("a:solidFill"),
            Choice("a:gradFill"),
            Choice("a:blipFill"),
            Choice("a:pattFill"),
            Choice("a:grpFill"),
        )
    )
