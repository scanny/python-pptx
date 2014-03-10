# encoding: utf-8

"""
Test data builders for presentation-related oxml elements
"""

from __future__ import absolute_import, print_function, unicode_literals

from ...unitdata import BaseBuilder


class CT_SlideLayoutIdListBuilder(BaseBuilder):
    __tag__ = 'p:sldLayoutIdLst'
    __nspfxs__ = ('p',)
    __attrs__ = ()


class CT_SlideMasterBuilder(BaseBuilder):
    __tag__ = 'p:sldMaster'
    __nspfxs__ = ('p',)
    __attrs__ = ('preserve',)


def a_sldLayoutIdLst():
    return CT_SlideLayoutIdListBuilder()


def a_sldMaster():
    return CT_SlideMasterBuilder()
