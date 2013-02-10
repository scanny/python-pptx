# -*- coding: utf-8 -*-
#
# test_oxml.py
#
# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

"""Test suite for pptx.oxml module."""

from hamcrest  import (assert_that, has_item, has_property, instance_of,
                       is_, is_not, equal_to, greater_than)
from mock      import Mock, patch
from unittest2 import TestCase

from lxml import etree
from lxml.etree import Element

from pptx.oxml import CT_Shape, CT_ShapeNonVisual
from pptx.packaging import prettify_nsdecls
from pptx.spec import namespaces, qtag

# default namespace map
nsmap = namespaces('a', 'r', 'p')

def txbox_xml():
    xml = (\
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n"
        '<p:sp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"\n'
        '      xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"\n'
        '      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n'
        '  <p:nvSpPr>\n'
        '    <p:cNvPr id="2" name="TextBox 1"/>\n'
        '    <p:cNvSpPr txBox="1"/>\n'
        '    <p:nvPr/>\n'
        '  </p:nvSpPr>\n'
        '  <p:spPr>\n'
        '    <a:xfrm>\n'
        '      <a:off x="1997289" y="2529664"/>\n'
        '      <a:ext cx="2390398" cy="369332"/>\n'
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
        '</p:sp>' )
    return xml


class TestCT_Shape(TestCase):
    """Test CT_Shape class"""
    def test_construction_xml(self):
        """CT_Shape() produces correct XML on construction"""
        # setup -----------------------
        sp = CT_Shape()
        sp.nvSpPr.cNvPr.id      = 2
        sp.nvSpPr.cNvPr.name    = 'TextBox 1'
        sp.nvSpPr.cNvSpPr.txBox = 1
        sp.spPr.xfrm.off.x      = 1997289
        sp.spPr.xfrm.off.y      = 2529664
        sp.spPr.xfrm.ext.cx     = 2390398
        sp.spPr.xfrm.ext.cy     =  369332
        sp.spPr.prstGeom.prst   = 'rect'
        sp.txBody.bodyPr.wrap   = 'none'
        # exercise --------------------
        xml = etree.tostring(sp.element, encoding='UTF-8',
                             pretty_print=True, standalone=True)
        xml = prettify_nsdecls(xml)
        xml_lines = xml.split('\n')
        txbox_xml_lines = txbox_xml().split('\n')
        # verify ----------------------
        # self.assertTrue(False, xml)
        for idx, line in enumerate(xml_lines):
            assert_that(line, is_(equal_to(txbox_xml_lines[idx])))
    

