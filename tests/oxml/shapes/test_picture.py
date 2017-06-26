# encoding: utf-8

"""
Test suite for pptx.oxml.shapes.picture module.
"""

from __future__ import absolute_import, print_function, unicode_literals

import pytest

from pptx.oxml.ns import nsdecls
from pptx.oxml.shapes.picture import CT_Picture


class DescribeCT_Picture(object):

    def it_can_create_a_new_pic_element(self, pic_fixture):
        shape_id, name, desc, rId, x, y, cx, cy, expected_xml = pic_fixture
        pic = CT_Picture.new_pic(shape_id, name, desc, rId, x, y, cx, cy)
        assert pic.xml == expected_xml

    def it_can_create_a_new_video_pic_element(self, video_pic_fixture):
        shape_id, shape_name, video_rId, media_rId = video_pic_fixture[:4]
        poster_frame_rId, x, y, cx, cy, expected_xml = video_pic_fixture[4:]
        pic = CT_Picture.new_video_pic(
            shape_id, shape_name, video_rId, media_rId, poster_frame_rId,
            x, y, cx, cy
        )
        print(pic.xml)
        assert pic.xml == expected_xml

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def pic_fixture(self):
        shape_id, name, desc, rId = 9, 'Diam. > 1 mm', 'desc', 'rId1'
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
                nsdecls('a', 'p', 'r'), shape_id, 'Diam. &gt; 1 mm', desc,
                rId, x, y, cx, cy
            )
        )
        return shape_id, name, desc, rId, x, y, cx, cy, expected_xml

    @pytest.fixture
    def video_pic_fixture(self):
        shape_id, shape_name = 42, 'media.mp4'
        video_rId, media_rId, poster_frame_rId = 'rId1', 'rId2', 'rId3'
        x, y, cx, cy = 1, 2, 3, 4
        expected_xml = (
            '<p:pic xmlns:a="http://schemas.openxmlformats.org/drawingml/200'
            '6/main" xmlns:p="http://schemas.openxmlformats.org/presentation'
            'ml/2006/main" xmlns:r="http://schemas.openxmlformats.org/office'
            'Document/2006/relationships">\n  <p:nvPicPr>\n    <p:cNvPr id="'
            '42" name="media.mp4">\n      <a:hlinkClick r:id="" action="ppac'
            'tion://media"/>\n    </p:cNvPr>\n    <p:cNvPicPr>\n      <a:pic'
            'Locks noChangeAspect="1"/>\n    </p:cNvPicPr>\n    <p:nvPr>\n  '
            '    <a:videoFile r:link="rId1"/>\n      <p:extLst>\n        <p:'
            'ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">\n          <p'
            '14:media xmlns:p14="http://schemas.microsoft.com/office/powerpo'
            'int/2010/main" r:embed="rId2"/>\n        </p:ext>\n      </p:ex'
            'tLst>\n    </p:nvPr>\n  </p:nvPicPr>\n  <p:blipFill>\n    <a:bl'
            'ip r:embed="rId3"/>\n    <a:stretch>\n      <a:fillRect/>\n    '
            '</a:stretch>\n  </p:blipFill>\n  <p:spPr>\n    <a:xfrm>\n      '
            '<a:off x="1" y="2"/>\n      <a:ext cx="3" cy="4"/>\n    </a:xfr'
            'm>\n    <a:prstGeom prst="rect">\n      <a:avLst/>\n    </a:prs'
            'tGeom>\n  </p:spPr>\n</p:pic>\n'
        )
        return (
            shape_id, shape_name, video_rId, media_rId, poster_frame_rId, x,
            y, cx, cy, expected_xml
        )
