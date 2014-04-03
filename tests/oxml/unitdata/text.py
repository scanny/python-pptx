# encoding: utf-8

"""
XML test data builders for pptx.oxml.text unit tests
"""

from __future__ import absolute_import, print_function

from ...unitdata import BaseBuilder


class CT_Hyperlink(BaseBuilder):
    __tag__ = 'a:hlinkClick'
    __nspfxs__ = ('a', 'r')
    __attrs__ = (
        'r:id', 'invalidUrl', 'action', 'tgtFrame', 'tooltip', 'history',
        'highlightClick', 'endSnd'
    )

    def with_rId(self, rId):
        self._set_xmlattr('r:id', rId)
        return self


class CT_OfficeArtExtensionList(BaseBuilder):
    __tag__ = 'a:extLst'
    __nspfxs__ = ('a',)
    __attrs__ = ()


class CT_RegularTextRunBuilder(BaseBuilder):
    __tag__ = 'a:r'
    __nspfxs__ = ('a',)
    __attrs__ = ()


class CT_TextBodyBuilder(BaseBuilder):
    __tag__ = 'p:txBody'
    __nspfxs__ = ('p', 'a')
    __attrs__ = ()


class CT_TextBodyPropertiesBuilder(BaseBuilder):
    __tag__ = 'a:bodyPr'
    __nspfxs__ = ('a',)
    __attrs__ = (
        'rot', 'spcFirstLastPara', 'vertOverflow', 'horzOverflow', 'vert',
        'wrap', 'lIns', 'tIns', 'rIns', 'bIns', 'numCol', 'spcCol', 'rtlCol',
        'fromWordArt', 'anchor', 'anchorCtr', 'forceAA', 'upright',
        'compatLnSpc',
    )


class CT_TextCharacterPropertiesBuilder(BaseBuilder):
    """
    Test data builder for CT_TextCharacterProperties XML element that appears
    as <a:endParaRPr> child of <a:p> and <a:rPr> child of <a:r>.
    """
    __nspfxs__ = ('a',)
    __attrs__ = ('b', 'i', 'sz', 'u')

    def __init__(self, tag):
        self.__tag__ = tag
        super(CT_TextCharacterPropertiesBuilder, self).__init__()


class CT_TextFontBuilder(BaseBuilder):
    __tag__ = 'a:latin'
    __nspfxs__ = ('a',)
    __attrs__ = ('typeface', 'panose', 'pitchFamily', 'charset')


class CT_TextNoAutofitBuilder(BaseBuilder):
    __tag__ = 'a:noAutofit'
    __nspfxs__ = ('a',)
    __attrs__ = ()


class CT_TextNormalAutofitBuilder(BaseBuilder):
    __tag__ = 'a:normAutofit'
    __nspfxs__ = ('a',)
    __attrs__ = ('fontScale', 'lnSpcReduction')


class CT_TextParagraphBuilder(BaseBuilder):
    """
    Test data builder for CT_TextParagraph (<a:p>) XML element that appears
    as a child of <p:txBody>.
    """
    __tag__ = 'a:p'
    __nspfxs__ = ('a',)
    __attrs__ = ()


class CT_TextParagraphPropertiesBuilder(BaseBuilder):
    """
    Test data builder for CT_TextParagraphProperties (<a:pPr>) XML element
    that appears as a child of <a:p>.
    """
    __tag__ = 'a:pPr'
    __nspfxs__ = ('a',)
    __attrs__ = (
        'marL', 'marR', 'lvl', 'indent', 'algn', 'defTabSz', 'rtl',
        'eaLnBrk', 'fontAlgn', 'latinLnBrk', 'hangingPunct'
    )


class CT_TextShapeAutofitBuilder(BaseBuilder):
    __tag__ = 'a:spAutoFit'
    __nspfxs__ = ('a',)
    __attrs__ = ()


class XsdString(BaseBuilder):
    __attrs__ = ()

    def __init__(self, tag, nspfxs):
        self.__tag__ = tag
        self.__nspfxs__ = nspfxs
        super(XsdString, self).__init__()


def a_bodyPr():
    return CT_TextBodyPropertiesBuilder()


def a_defRPr():
    return CT_TextCharacterPropertiesBuilder('a:defRPr')


def a_latin():
    return CT_TextFontBuilder()


def a_noAutofit():
    return CT_TextNoAutofitBuilder()


def a_normAutofit():
    return CT_TextNormalAutofitBuilder()


def a_p():
    """Return a CT_TextParagraphBuilder instance"""
    return CT_TextParagraphBuilder()


def a_pPr():
    """Return a CT_TextParagraphPropertiesBuilder instance"""
    return CT_TextParagraphPropertiesBuilder()


def a_t():
    return XsdString('a:t', ('a',))


def a_txBody():
    return CT_TextBodyBuilder()


def an_endParaRPr():
    return CT_TextCharacterPropertiesBuilder('a:endParaRPr')


def an_extLst():
    return CT_OfficeArtExtensionList()


def an_hlinkClick():
    return CT_Hyperlink()


def an_r():
    return CT_RegularTextRunBuilder()


def an_rPr():
    return CT_TextCharacterPropertiesBuilder('a:rPr')


def an_spAutoFit():
    return CT_TextShapeAutofitBuilder()
