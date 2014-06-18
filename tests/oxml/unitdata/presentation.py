# encoding: utf-8

"""
Test data builders for presentation-related oxml elements
"""

from __future__ import absolute_import, print_function, unicode_literals

from ...unitdata import BaseBuilder


class CT_PositiveSize2DBuilder(BaseBuilder):
    __tag__ = 'p:notesSz'
    __nspfxs__ = ('p',)
    __attrs__ = ('cx', 'cy')


class CT_PresentationBuilder(BaseBuilder):
    __tag__ = 'p:presentation'
    __nspfxs__ = ('p',)
    __attrs__ = (
        'serverZoom', 'firstSlideNum', 'showSpecialPlsOnTitleSld', 'rtl',
        'removePersonalInfoOnSave', 'compatMode', 'strictFirstAndLastChars',
        'embedTrueTypeFonts', 'saveSubsetFonts', 'autoCompressPictures',
        'bookmarkIdSeed', 'conformance'
    )


class CT_SlideIdBuilder(BaseBuilder):
    __tag__ = 'p:sldId'
    __nspfxs__ = ('p', 'r')
    __attrs__ = ('r:id', 'id')

    def with_rId(self, rId):
        self._set_xmlattr('r:id', rId)
        return self


class CT_SlideIdListBuilder(BaseBuilder):
    __tag__ = 'p:sldIdLst'
    __nspfxs__ = ('p', 'r')
    __attrs__ = ()


class CT_SlideMasterIdBuilder(BaseBuilder):
    __tag__ = 'p:sldMasterId'
    __nspfxs__ = ('p',)
    __attrs__ = ('id', 'r:id')

    def with_rId(self, rId):
        self._set_xmlattr('r:id', rId)
        return self


class CT_SlideMasterIdListBuilder(BaseBuilder):
    __tag__ = 'p:sldMasterIdLst'
    __nspfxs__ = ('p',)
    __attrs__ = ()


class CT_SlideSizeBuilder(BaseBuilder):
    __tag__ = 'p:sldSz'
    __nspfxs__ = ('p',)
    __attrs__ = ('cx', 'cy', 'type')


def a_notesSz():
    return CT_PositiveSize2DBuilder()


def a_presentation():
    """Return a CT_PresentationBuilder instance"""
    return CT_PresentationBuilder()


def a_sldId():
    return CT_SlideIdBuilder()


def a_sldIdLst():
    return CT_SlideIdListBuilder()


def a_sldMasterId():
    return CT_SlideMasterIdBuilder()


def a_sldMasterIdLst():
    return CT_SlideMasterIdListBuilder()


def a_sldSz():
    return CT_SlideSizeBuilder()
