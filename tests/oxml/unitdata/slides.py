# encoding: utf-8

"""
Test data builders for presentation-related oxml elements
"""

from __future__ import absolute_import, print_function, unicode_literals

from ...unitdata import BaseBuilder


class CT_CommonSlideDataBuilder(BaseBuilder):
    __tag__ = 'p:cSld'
    __nspfxs__ = ('p',)
    __attrs__ = ('name',)


class CT_SlideBuilder(BaseBuilder):
    __tag__ = 'p:sld'
    __nspfxs__ = ('p',)
    __attrs__ = ('showMasterSp', 'showMasterPhAnim', 'show')


class CT_SlideLayoutIdBuilder(BaseBuilder):
    __tag__ = 'p:sldLayoutId'
    __nspfxs__ = ('p',)
    __attrs__ = ('id', 'r:id')

    def with_rId(self, rId):
        self._set_xmlattr('r:id', rId)
        return self


class CT_SlideLayoutIdListBuilder(BaseBuilder):
    __tag__ = 'p:sldLayoutIdLst'
    __nspfxs__ = ('p',)
    __attrs__ = ()


class CT_SlideMasterBuilder(BaseBuilder):
    __tag__ = 'p:sldMaster'
    __nspfxs__ = ('p',)
    __attrs__ = ('preserve',)


def a_cSld():
    return CT_CommonSlideDataBuilder()


def a_sld():
    return CT_SlideBuilder()


def a_sldLayoutId():
    return CT_SlideLayoutIdBuilder()


def a_sldLayoutIdLst():
    return CT_SlideLayoutIdListBuilder()


def a_sldMaster():
    return CT_SlideMasterBuilder()
