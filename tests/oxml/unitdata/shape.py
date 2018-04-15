# encoding: utf-8

"""
Test data for autoshape-related unit tests.
"""

from __future__ import absolute_import

from ...unitdata import BaseBuilder


class CT_ApplicationNonVisualDrawingPropsBuilder(BaseBuilder):
    __tag__ = 'p:nvPr'
    __nspfxs__ = ('p',)
    __attrs__ = ('isPhoto', 'userDrawn')


class CT_ConnectorBuilder(BaseBuilder):
    __tag__ = 'p:cxnSp'
    __nspfxs__ = ('p', 'a')
    __attrs__ = ()


class CT_GeomGuideBuilder(BaseBuilder):
    __tag__ = 'a:gd'
    __nspfxs__ = ('a',)
    __attrs__ = ('name', 'fmla')


class CT_GeomGuideListBuilder(BaseBuilder):
    __tag__ = 'a:avLst'
    __nspfxs__ = ('a',)
    __attrs__ = ()


class CT_GraphicalObjectBuilder(BaseBuilder):
    __tag__ = 'a:graphic'
    __nspfxs__ = ('a',)
    __attrs__ = ()


class CT_GraphicalObjectDataBuilder(BaseBuilder):
    __tag__ = 'a:graphicData'
    __nspfxs__ = ('a',)
    __attrs__ = ('uri',)


class CT_GraphicalObjectFrameBuilder(BaseBuilder):
    __tag__ = 'p:graphicFrame'
    __nspfxs__ = ('p', 'a')
    __attrs__ = ('bwMode',)


class CT_GraphicalObjectFrameNonVisualBuilder(BaseBuilder):
    __tag__ = 'p:nvGraphicFramePr'
    __nspfxs__ = ('p',)
    __attrs__ = ()


class CT_GroupShapeBuilder(BaseBuilder):
    __nspfxs__ = ('p',)
    __attrs__ = ()

    def __init__(self, tag):
        self.__tag__ = tag
        super(CT_GroupShapeBuilder, self).__init__()


class CT_GroupShapePropertiesBuilder(BaseBuilder):
    __tag__ = 'p:grpSpPr'
    __nspfxs__ = ('p', 'a')
    __attrs__ = ('bwMode',)


class CT_NonVisualDrawingPropsBuilder(BaseBuilder):
    __tag__ = 'p:cNvPr'
    __nspfxs__ = ('p',)
    __attrs__ = ('id', 'name', 'descr', 'hidden', 'title')


class CT_NonVisualDrawingShapePropsBuilder(BaseBuilder):
    __tag__ = 'p:cNvSpPr'
    __nspfxs__ = ('p',)
    __attrs__ = ('txBox',)


class CT_PictureBuilder(BaseBuilder):
    __tag__ = 'p:pic'
    __nspfxs__ = ('p', 'a')
    __attrs__ = ()


class CT_PictureNonVisualBuilder(BaseBuilder):
    __tag__ = 'p:nvPicPr'
    __nspfxs__ = ('p',)
    __attrs__ = ()


class CT_PlaceholderBuilder(BaseBuilder):
    __tag__ = 'p:ph'
    __nspfxs__ = ('p',)
    __attrs__ = ('type', 'orient', 'sz', 'idx', 'hasCustomPropt')


class CT_Point2DBuilder(BaseBuilder):
    __tag__ = 'a:off'
    __nspfxs__ = ('a',)
    __attrs__ = ('x', 'y')


class CT_PositiveSize2DBuilder(BaseBuilder):
    __tag__ = 'a:ext'
    __nspfxs__ = ('a',)
    __attrs__ = ('cx', 'cy')


class CT_PresetGeometry2DBuilder(BaseBuilder):
    __tag__ = 'a:prstGeom'
    __nspfxs__ = ('a',)
    __attrs__ = ('prst',)


class CT_ShapeBuilder(BaseBuilder):
    __tag__ = 'p:sp'
    __nspfxs__ = ('p', 'a')
    __attrs__ = ('useBgFill',)


class CT_ShapeNonVisualBuilder(BaseBuilder):
    __tag__ = 'p:nvSpPr'
    __nspfxs__ = ('p',)
    __attrs__ = ()


class CT_ShapePropertiesBuilder(BaseBuilder):
    __tag__ = 'p:spPr'
    __nspfxs__ = ('p', 'a')
    __attrs__ = ('bwMode',)


class CT_Transform2DBuilder(BaseBuilder):
    __nspfxs__ = ('a',)
    __attrs__ = ('rot', 'flipH', 'flipV')

    def __init__(self, tag):
        self.__tag__ = tag
        super(CT_Transform2DBuilder, self).__init__()


def a_cNvPr():
    return CT_NonVisualDrawingPropsBuilder()


def a_cNvSpPr():
    return CT_NonVisualDrawingShapePropsBuilder()


def a_cxnSp():
    return CT_ConnectorBuilder()


def a_gd():
    return CT_GeomGuideBuilder()


def a_graphic():
    return CT_GraphicalObjectBuilder()


def a_graphicData():
    return CT_GraphicalObjectDataBuilder()


def a_graphicFrame():
    return CT_GraphicalObjectFrameBuilder()


def a_grpSp():
    return CT_GroupShapeBuilder('p:grpSp')


def a_grpSpPr():
    return CT_GroupShapePropertiesBuilder()


def a_p_xfrm():
    return CT_Transform2DBuilder('p:xfrm')


def a_ph():
    return CT_PlaceholderBuilder()


def a_pic():
    return CT_PictureBuilder()


def a_prstGeom():
    return CT_PresetGeometry2DBuilder()


def an_avLst():
    return CT_GeomGuideListBuilder()


def an_ext():
    return CT_PositiveSize2DBuilder()


def an_nvGraphicFramePr():
    return CT_GraphicalObjectFrameNonVisualBuilder()


def an_nvPicPr():
    return CT_PictureNonVisualBuilder()


def an_nvPr():
    return CT_ApplicationNonVisualDrawingPropsBuilder()


def an_nvSpPr():
    return CT_ShapeNonVisualBuilder()


def an_off():
    return CT_Point2DBuilder()


def an_sp():
    return CT_ShapeBuilder()


def an_spPr():
    return CT_ShapePropertiesBuilder()


def an_xfrm():
    return CT_Transform2DBuilder('a:xfrm')
