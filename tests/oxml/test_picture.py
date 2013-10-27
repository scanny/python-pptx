# encoding: utf-8

"""Test suite for pptx.oxml.picture module."""

from __future__ import absolute_import

from pptx.oxml.ns import nsdecls
from pptx.oxml.picture import CT_Picture

from ..unitutil import TestCase


class TestCT_Picture(TestCase):
    """Test CT_Picture"""
    def test_new_pic_generates_correct_xml(self):
        """CT_Picture.new_pic() returns correct XML"""
        # setup ------------------------
        id_, name, desc, rId = 9, 'Picture 8', 'test-image.png', 'rId7'
        left, top, width, height = 111, 222, 333, 444
        xml = (
            '<p:pic %s>\n  <p:nvPicPr>\n    <p:cNvPr id="%d" name="%s" descr='
            '"%s"/>\n    <p:cNvPicPr>\n      <a:picLocks noChangeAspect="1"/>'
            '\n    </p:cNvPicPr>\n    <p:nvPr/>\n  </p:nvPicPr>\n  <p:blipFil'
            'l>\n    <a:blip r:embed="%s"/>\n    <a:stretch>\n      <a:fillRe'
            'ct/>\n    </a:stretch>\n  </p:blipFill>\n  <p:spPr>\n    <a:xfrm'
            '>\n      <a:off x="%d" y="%d"/>\n      <a:ext cx="%d" cy="%d"/>'
            '\n    </a:xfrm>\n    <a:prstGeom prst="rect">\n      <a:avLst/>'
            '\n    </a:prstGeom>\n  </p:spPr>\n</p:pic>\n' %
            (nsdecls('a', 'p', 'r'), id_, name, desc, rId, left, top, width,
             height)
        )
        # exercise ---------------------
        pic = CT_Picture.new_pic(id_, name, desc, rId, left, top,
                                 width, height)
        # verify -----------------------
        self.assertEqualLineByLine(xml, pic)
