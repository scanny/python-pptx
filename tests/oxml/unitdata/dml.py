# encoding: utf-8

"""
XML test data builders for pptx.oxml.dml unit tests
"""

from __future__ import absolute_import, print_function

from ...unitdata import BaseBuilder


class CT_BlipFillPropertiesBuilder(BaseBuilder):
    __tag__ = 'a:blipFill'
    __nspfxs__ = ('a',)
    __attrs__ = ('dpi', 'rotWithShape')


class CT_GradientFillPropertiesBuilder(BaseBuilder):
    __tag__ = 'a:gradFill'
    __nspfxs__ = ('a',)
    __attrs__ = ('flip', 'rotWithShape')


class CT_GroupFillPropertiesBuilder(BaseBuilder):
    __tag__ = 'a:grpFill'
    __nspfxs__ = ('a',)
    __attrs__ = ()


class CT_HslColorBuilder(BaseBuilder):
    __tag__ = 'a:hslClr'
    __nspfxs__ = ('a',)
    __attrs__ = ('hue', 'sat', 'lum')


class CT_LinePropertiesBuilder(BaseBuilder):
    __tag__ = 'a:ln'
    __nspfxs__ = ('a',)
    __attrs__ = ('w', 'cap', 'cmpd', 'algn')


class CT_NoFillPropertiesBuilder(BaseBuilder):
    __tag__ = 'a:noFill'
    __nspfxs__ = ('a',)
    __attrs__ = ()


class CT_PatternFillPropertiesBuilder(BaseBuilder):
    __tag__ = 'a:pattFill'
    __nspfxs__ = ('a',)
    __attrs__ = ('prst',)


class CT_PercentageBuilder(BaseBuilder):
    __nspfxs__ = ('a',)
    __attrs__ = ('val',)

    def __init__(self, tag):
        self.__tag__ = tag
        super(CT_PercentageBuilder, self).__init__()


class CT_PresetColorBuilder(BaseBuilder):
    __tag__ = 'a:prstClr'
    __nspfxs__ = ('a',)
    __attrs__ = ('val',)


class CT_SolidColorFillPropertiesBuilder(BaseBuilder):
    __tag__ = 'a:solidFill'
    __nspfxs__ = ('a',)
    __attrs__ = ()


class CT_SchemeColorBuilder(BaseBuilder):
    __tag__ = 'a:schemeClr'
    __nspfxs__ = ('a',)
    __attrs__ = ('val',)


class CT_ScRgbColorBuilder(BaseBuilder):
    __tag__ = 'a:scrgbClr'
    __nspfxs__ = ('a',)
    __attrs__ = ('r', 'g', 'b')


class CT_ShapePropertiesBuilder(BaseBuilder):
    __tag__ = 'p:spPr'
    __nspfxs__ = ('p', 'a')
    __attrs__ = ('bwMode',)


class CT_SRgbColorBuilder(BaseBuilder):
    __tag__ = 'a:srgbClr'
    __nspfxs__ = ('a',)
    __attrs__ = ('val',)


class CT_SystemColorBuilder(BaseBuilder):
    __tag__ = 'a:sysClr'
    __nspfxs__ = ('a',)
    __attrs__ = ('val', 'lastClr')


def a_blipFill():
    return CT_BlipFillPropertiesBuilder()


def a_gradFill():
    return CT_GradientFillPropertiesBuilder()


def a_grpFill():
    return CT_GroupFillPropertiesBuilder()


def a_lumMod():
    return CT_PercentageBuilder('a:lumMod')


def a_lumOff():
    return CT_PercentageBuilder('a:lumOff')


def a_noFill():
    return CT_NoFillPropertiesBuilder()


def a_pattFill():
    return CT_PatternFillPropertiesBuilder()


def a_prstClr():
    return CT_PresetColorBuilder()


def a_schemeClr():
    return CT_SchemeColorBuilder()


def a_solidFill():
    return CT_SolidColorFillPropertiesBuilder()


def an_hslClr():
    return CT_HslColorBuilder()


def an_ln():
    return CT_LinePropertiesBuilder()


def an_spPr():
    return CT_ShapePropertiesBuilder()


def an_scrgbClr():
    return CT_ScRgbColorBuilder()


def an_srgbClr():
    return CT_SRgbColorBuilder()


def a_sysClr():
    return CT_SystemColorBuilder()
