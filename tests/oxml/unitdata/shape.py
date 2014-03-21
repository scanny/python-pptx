# encoding: utf-8

"""
Test data for autoshape-related unit tests.
"""

from __future__ import absolute_import

from ...unitdata import BaseBuilder

from pptx.oxml import parse_xml_bytes
from pptx.oxml.ns import nsdecls
from pptx.shapes.shapetree import ShapeCollection


class CT_ApplicationNonVisualDrawingPropsBuilder(BaseBuilder):
    __tag__ = 'p:nvPr'
    __nspfxs__ = ('p',)
    __attrs__ = ('isPhoto', 'userDrawn')


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
    __tag__ = 'a:xfrm'
    __nspfxs__ = ('a',)
    __attrs__ = ('rot', 'flipH', 'flipV')


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


def an_spTree():
    return CT_GroupShapeBuilder('p:spTree')


def an_xfrm():
    return CT_Transform2DBuilder()


class _TestShapeXml(object):
    """XML snippets of various shapes for use in unit tests"""
    @property
    def autoshape(self):
        """
        XML for an autoshape for unit testing purposes, a rounded rectangle in
        this case.
        """
        return (
            '<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/'
            '2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/'
            '2006/main"><p:nvSpPr><p:cNvPr id="3" name="Rounded Rectangle 2"/'
            '><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="76009'
            '6" y="562720"/><a:ext cx="2520824" cy="914400"/></a:xfrm><a:prst'
            'Geom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 30346"'
            '/></a:avLst></a:prstGeom></p:spPr><p:style><a:lnRef idx="1"><a:s'
            'chemeClr val="accent1"/></a:lnRef><a:fillRef idx="3"><a:schemeCl'
            'r val="accent1"/></a:fillRef><a:effectRef idx="2"><a:schemeClr v'
            'al="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr '
            'val="lt1"/></a:fontRef></p:style><p:txBody><a:bodyPr rtlCol="0" '
            'anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr l'
            'ang="en-US" dirty="0" smtClean="0"/><a:t>This is text inside a r'
            'ounded rectangle</a:t></a:r><a:endParaRPr lang="en-US" dirty="0"'
            '/></a:p></p:txBody></p:sp>'
        )

    @property
    def empty_spTree(self):
        return (
            '<p:spTree %s>\n'
            '  <p:nvGrpSpPr>\n'
            '    <p:cNvPr id="1" name=""/>\n'
            '    <p:cNvGrpSpPr/>\n'
            '    <p:nvPr/>\n'
            '  </p:nvGrpSpPr>\n'
            '  <p:grpSpPr/>\n'
            '</p:spTree>\n' % nsdecls('p', 'a')
        )

    @property
    def placeholder(self):
        """Generic placeholder XML, a date placeholder in this case"""
        return (
            '<p:sp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/'
            'main" xmlns:p="http://schemas.openxmlformats.org/presentationml/'
            '2006/main">\n'
            '  <p:nvSpPr>\n'
            '    <p:cNvPr id="9" name="Date Placeholder 8"/>\n'
            '    <p:cNvSpPr>\n'
            '      <a:spLocks noGrp="1"/>\n'
            '    </p:cNvSpPr>\n'
            '    <p:nvPr>\n'
            '      <p:ph type="dt" sz="half" idx="10"/>\n'
            '    </p:nvPr>\n'
            '  </p:nvSpPr>\n'
            '  <p:spPr/>\n'
            '</p:sp>\n'
        )

    @property
    def rounded_rectangle(self):
        """XML for a rounded rectangle auto shape"""
        return self.autoshape

    @property
    def textbox(self):
        """Generic text box XML"""
        return (
            '<p:sp %s>\n'
            '  <p:nvSpPr>\n'
            '    <p:cNvPr id="9" name="TextBox 8"/>\n'
            '    <p:cNvSpPr txBox="1"/>\n'
            '    <p:nvPr/>\n'
            '  </p:nvSpPr>\n'
            '  <p:spPr>\n'
            '    <a:xfrm>\n'
            '      <a:off x="111" y="222"/>\n'
            '      <a:ext cx="333" cy="444"/>\n'
            '    </a:xfrm>\n'
            '    <a:prstGeom prst="rect">\n'
            '      <a:avLst/>\n'
            '    </a:prstGeom>\n'
            '    <a:noFill/>\n'
            '  </p:spPr>\n'
            '  <p:txBody>\n'
            '    <a:bodyPr wrap="none">\n'
            '      <a:spAutoFit/>\n'
            '    </a:bodyPr>\n'
            '    <a:lstStyle/>\n'
            '    <a:p/>\n'
            '  </p:txBody>\n'
            '</p:sp>' % nsdecls('a', 'p')
        )


class _TestShapeElements(object):
    """Shape elements for use in unit tests"""
    @property
    def autoshape(self):
        return parse_xml_bytes(test_shape_xml.autoshape)

    @property
    def empty_spTree(self):
        return parse_xml_bytes(test_shape_xml.empty_spTree)

    @property
    def placeholder(self):
        return parse_xml_bytes(test_shape_xml.placeholder)

    @property
    def rounded_rectangle(self):
        return parse_xml_bytes(test_shape_xml.rounded_rectangle)

    @property
    def textbox(self):
        return parse_xml_bytes(test_shape_xml.textbox)


class _TestShapes(object):
    """Shape instances for use in unit tests"""
    @property
    def empty_shape_collection(self):
        return ShapeCollection(test_shape_elements.empty_spTree)


test_shape_xml = _TestShapeXml()
test_shape_elements = _TestShapeElements()
test_shapes = _TestShapes()
