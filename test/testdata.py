# -*- coding: utf-8 -*-
#
# testdata.py
#
# Copyright (C) 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

"""Test data for unit tests"""

from pptx.oxml import nsdecls, oxml_fromstring
from pptx.shapes import _Paragraph, _Picture, _Shape, _ShapeCollection


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
    def picture(self):
        """ XML for an pic shape, for unit testing purposes """
        return (
            '<p:pic %s>\n'
            '  <p:nvPicPr>\n'
            '    <p:cNvPr id="9" name="Picture 8" descr="image.png"/>\n'
            '    <p:cNvPicPr>\n'
            '      <a:picLocks noChangeAspect="1"/>\n'
            '    </p:cNvPicPr>\n'
            '    <p:nvPr/>\n'
            '  </p:nvPicPr>\n'
            '  <p:blipFill>\n'
            '    <a:blip r:embed="rId7"/>\n'
            '    <a:stretch>\n'
            '      <a:fillRect/>\n'
            '    </a:stretch>\n'
            '  </p:blipFill>\n'
            '  <p:spPr>\n'
            '    <a:xfrm>\n'
            '      <a:off x="111" y="222"/>\n'
            '      <a:ext cx="333" cy="444"/>\n'
            '    </a:xfrm>\n'
            '    <a:prstGeom prst="rect">\n'
            '      <a:avLst/>\n'
            '    </a:prstGeom>\n'
            '  </p:spPr>\n'
            '</p:pic>\n' % nsdecls('a', 'p', 'r')
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


test_shape_xml = _TestShapeXml()


class _TestTextXml(object):
    """XML snippets of text-related elements for use in unit tests"""
    @property
    def paragraph(self):
        """
        XML for an autoshape for unit testing purposes, a rounded rectangle in
        this case.
        """
        return '<a:p %s/>' % nsdecls('a')


test_text_xml = _TestTextXml()


class _TestShapeElements(object):
    """Shape elements for use in unit tests"""
    @property
    def autoshape(self):
        return oxml_fromstring(test_shape_xml.autoshape)

    @property
    def empty_spTree(self):
        return oxml_fromstring(test_shape_xml.empty_spTree)

    @property
    def picture(self):
        return oxml_fromstring(test_shape_xml.picture)

    @property
    def placeholder(self):
        return oxml_fromstring(test_shape_xml.placeholder)

    @property
    def rounded_rectangle(self):
        return oxml_fromstring(test_shape_xml.rounded_rectangle)

    @property
    def textbox(self):
        return oxml_fromstring(test_shape_xml.textbox)


test_shape_elements = _TestShapeElements()


class _TestTextElements(object):
    """Text elements for use in unit tests"""
    @property
    def paragraph(self):
        return oxml_fromstring(test_text_xml.paragraph)


test_text_elements = _TestTextElements()


class _TestShapes(object):
    """Shape instances for use in unit tests"""
    @property
    def autoshape(self):
        return _Shape(test_shape_elements.autoshape)

    @property
    def empty_shape_collection(self):
        return _ShapeCollection(test_shape_elements.empty_spTree)

    @property
    def picture(self):
        return _Picture(test_shape_elements.picture)

    @property
    def placeholder(self):
        return _Shape(test_shape_elements.placeholder)

    @property
    def rounded_rectangle(self):
        return _Shape(test_shape_elements.rounded_rectangle)

    @property
    def textbox(self):
        return _Shape(test_shape_elements.textbox)


test_shapes = _TestShapes()


class _TestTextObjects(object):
    """Text object instances for use in unit tests"""
    @property
    def paragraph(self):
        return _Paragraph(test_text_elements.paragraph)


test_text_objects = _TestTextObjects()
