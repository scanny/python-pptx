# encoding: utf-8

"""
XML test data builders for pptx.oxml.text unit tests
"""

from __future__ import absolute_import, print_function

from ...unitdata import BaseBuilder


class CT_Hyperlink(BaseBuilder):
    __tag__ = "a:hlinkClick"
    __nspfxs__ = ("a", "r", )
    __attrs__ = (
        "r:id",
        "invalidUrl",
        "action",
        "tgtFrame",
        "tooltip",
        "history",
        "highlightClick",
        "endSnd",
    )

    def with_rId(self, rId):
        self._set_xmlattr("r:id", rId)
        return self


class CT_OfficeArtExtensionList(BaseBuilder):
    __tag__ = "a:extLst"
    __nspfxs__ = ("a",)
    __attrs__ = ()


class CT_PositiveSize2D(BaseBuilder):
    __tag__ = "a:ext"
    __nspfxs__ = ("a",)
    __attrs__ = ("uri")

    def with_uri(self, uri):
        self._set_xmlattr("uri", uri)
        return self

class HyperlinkColorExtension(BaseBuilder):
    __tag__ = "ahyp:hlinkClr"
    __nspfxs__ = ("ahyp",)
    __attrs__ = ("val")

    def with_val(self, val):
        self._set_xmlattr("val", val)
        return self

class CT_RegularTextRunBuilder(BaseBuilder):
    __tag__ = "a:r"
    __nspfxs__ = ("a",)
    __attrs__ = ()


class CT_TextCharacterPropertiesBuilder(BaseBuilder):
    """
    Test data builder for CT_TextCharacterProperties XML element that appears
    as <a:endParaRPr> child of <a:p> and <a:rPr> child of <a:r>.
    """

    __nspfxs__ = ("a",)
    __attrs__ = ("b", "i", "sz", "u")

    def __init__(self, tag):
        self.__tag__ = tag
        super(CT_TextCharacterPropertiesBuilder, self).__init__()


class CT_TextParagraphBuilder(BaseBuilder):
    """
    Test data builder for CT_TextParagraph (<a:p>) XML element that appears
    as a child of <p:txBody>.
    """

    __tag__ = "a:p"
    __nspfxs__ = ("a",)
    __attrs__ = ()

class CT_TextParagraphPropertiesBuilder(BaseBuilder):
    """
    Test data builder for CT_TextParagraphProperties (<a:pPr>) XML element
    that appears as a child of <a:p>.
    """

    __tag__ = "a:pPr"
    __nspfxs__ = ("a",)
    __attrs__ = (
        "marL",
        "marR",
        "lvl",
        "indent",
        "algn",
        "defTabSz",
        "rtl",
        "eaLnBrk",
        "fontAlgn",
        "latinLnBrk",
        "hangingPunct",
    )
class XsdString(BaseBuilder):
    __attrs__ = ()

    def __init__(self, tag, nspfxs):
        self.__tag__ = tag
        self.__nspfxs__ = nspfxs
        super(XsdString, self).__init__()


def a_p():
    """Return a CT_TextParagraphBuilder instance"""
    return CT_TextParagraphBuilder()


def a_t():
    return XsdString("a:t", ("a",))


def an_endParaRPr():
    return CT_TextCharacterPropertiesBuilder("a:endParaRPr")


def an_extLst():
    return CT_OfficeArtExtensionList()

def an_ext():
    return CT_PositiveSize2D()

def an_hlinkClr():
    return HyperlinkColorExtension()

def an_hlinkClick():
    return CT_Hyperlink()

def an_r():
    return CT_RegularTextRunBuilder()


def an_rPr():
    return CT_TextCharacterPropertiesBuilder("a:rPr")

def an_pPr():
    """Return a CT_TextParagraphPropertiesBuilder instance"""
    return CT_TextParagraphPropertiesBuilder()
