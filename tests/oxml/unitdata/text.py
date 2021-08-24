# encoding: utf-8

"""
XML test data builders for pptx.oxml.text unit tests
"""

from __future__ import absolute_import, print_function

from ...unitdata import BaseBuilder


class CT_Hyperlink(BaseBuilder):
    __tag__ = "a:hlinkClick"
    __nspfxs__ = ("a", "r")
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


def an_hlinkClick():
    return CT_Hyperlink()


def an_r():
    return CT_RegularTextRunBuilder()


def an_rPr():
    return CT_TextCharacterPropertiesBuilder("a:rPr")
