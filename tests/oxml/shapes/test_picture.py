# encoding: utf-8

"""
Test suite for pptx.oxml.shapes.picture module.
"""

from __future__ import absolute_import, print_function, unicode_literals

import pytest

from pptx.oxml.ns import nsdecls
from pptx.oxml.shapes.picture import CT_Picture


class DescribeCT_Picture(object):

    def it_can_generate_a_new_pic_element(self, new_fixture):
        id_, name, desc, rId, x, y, cx, cy, expected_xml = new_fixture
        pic = CT_Picture.new_pic(id_, name, desc, rId, x, y, cx, cy)
        print(expected_xml)
        assert pic.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def new_fixture(self):
        id_, name, desc, rId = 9, 'Diam. > 1 mm', 'desc', 'rId1'
        x, y, cx, cy = 1, 2, 3, 4
        expected_xml = (
            '<p:pic %s>\n  <p:nvPicPr>\n    <p:cNvPr id="%d" name="%s" descr'
            '="%s"/>\n    <p:cNvPicPr>\n      <a:picLocks noChangeAspect="1"'
            '/>\n    </p:cNvPicPr>\n    <p:nvPr/>\n  </p:nvPicPr>\n  <p:blip'
            'Fill>\n    <a:blip r:embed="%s"/>\n    <a:stretch>\n      <a:fi'
            'llRect/>\n    </a:stretch>\n  </p:blipFill>\n  <p:spPr>\n    <a'
            ':xfrm>\n      <a:off x="%d" y="%d"/>\n      <a:ext cx="%d" cy="'
            '%d"/>\n    </a:xfrm>\n    <a:prstGeom prst="rect">\n      <a:av'
            'Lst/>\n    </a:prstGeom>\n  </p:spPr>\n</p:pic>\n' % (
                nsdecls('a', 'p', 'r'), id_, 'Diam. &gt; 1 mm', desc, rId,
                x, y, cx, cy
            )
        )
        return id_, name, desc, rId, x, y, cx, cy, expected_xml
