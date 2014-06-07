# encoding: utf-8

"""Test suite for pptx.oxml.graphfrm module."""

from __future__ import absolute_import

from hamcrest import assert_that, equal_to, is_

from pptx.oxml.ns import nsdecls
from pptx.oxml.shapes.graphfrm import CT_GraphicalObjectFrame

from ...unitutil import TestCase


class TestCT_GraphicalObjectFrame(TestCase):

    def test_has_table_return_value(self):
        """
        CT_GraphicalObjectFrame.has_table property has correct value
        """
        # setup ------------------------
        id_, name = 9, 'Table 8'
        left, top, width, height = 111, 222, 333, 444
        tbl_uri = 'http://schemas.openxmlformats.org/drawingml/2006/table'
        chart_uri = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
        graphicFrame = CT_GraphicalObjectFrame.new_graphicFrame(
            id_, name, left, top, width, height)
        graphicData = graphicFrame.graphic.graphicData
        # verify -----------------------
        graphicData.set('uri', tbl_uri)
        assert_that(graphicFrame.has_table, is_(equal_to(True)))
        graphicData.set('uri', chart_uri)
        assert_that(graphicFrame.has_table, is_(equal_to(False)))

    def test_new_graphicFrame_generates_correct_xml(self):
        """CT_GraphicalObjectFrame.new_graphicFrame() returns correct XML"""
        # setup ------------------------
        id_, name = 9, 'Table 8'
        left, top, width, height = 111, 222, 333, 444
        xml = (
            '<p:graphicFrame %s>\n  <p:nvGraphicFramePr>\n    <p:cNvPr id="%d'
            '" name="%s"/>\n    <p:cNvGraphicFramePr>\n      <a:graphicFrameL'
            'ocks noGrp="1"/>\n    </p:cNvGraphicFramePr>\n    <p:nvPr/>\n  <'
            '/p:nvGraphicFramePr>\n  <p:xfrm>\n    <a:off x="%d" y="%d"/>\n  '
            '  <a:ext cx="%d" cy="%d"/>\n  </p:xfrm>\n  <a:graphic>\n    <a:g'
            'raphicData/>\n  </a:graphic>\n</p:graphicFrame>\n' %
            (nsdecls('a', 'p'), id_, name, left, top, width, height)
        )
        # exercise ---------------------
        graphicFrame = CT_GraphicalObjectFrame.new_graphicFrame(
            id_, name, left, top, width, height)
        # verify -----------------------
        self.assertEqualLineByLine(xml, graphicFrame)

    def test_new_table_generates_correct_xml(self):
        """CT_GraphicalObjectFrame.new_table() returns correct XML"""
        # setup ------------------------
        id_, name = 9, 'Table 8'
        rows, cols = 2, 3
        left, top, width, height = 111, 222, 334, 445
        xml = (
            '<p:graphicFrame %s>\n  <p:nvGraphicFramePr>\n    <p:cNvP''r id="'
            '%d" name="%s"/>\n    <p:cNvGraphicFramePr>\n      <a:graphicFram'
            'eLocks noGrp="1"/>\n    </p:cNvGraphicFramePr>\n    <p:nvPr/>\n '
            ' </p:nvGraphicFramePr>\n  <p:xfrm>\n    <a:off x="%d" y="%d"/>\n'
            '    <a:ext cx="%d" cy="%d"/>\n  </p:xfrm>\n  <a:graphic>\n    <a'
            ':graphicData uri="http://schemas.openxmlformats.org/drawingml/20'
            '06/table">\n      <a:tbl>\n        <a:tblPr firstRow="1" bandRow'
            '="1">\n          <a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9'
            'FD1C3A}</a:tableStyleId>\n        </a:tblPr>\n        <a:tblGrid'
            '>\n          <a:gridCol w="111"/>\n          <a:gridCol w="111"/'
            '>\n          <a:gridCol w="112"/>\n        </a:tblGrid>\n       '
            ' <a:tr h="222">\n          <a:tc>\n            <a:txBody>\n     '
            '         <a:bodyPr/>\n              <a:lstStyle/>\n             '
            ' <a:p/>\n            </a:txBody>\n            <a:tcPr/>\n       '
            '   </a:tc>\n          <a:tc>\n            <a:txBody>\n          '
            '    <a:bodyPr/>\n              <a:lstStyle/>\n              <a:p'
            '/>\n            </a:txBody>\n            <a:tcPr/>\n          </'
            'a:tc>\n          <a:tc>\n            <a:txBody>\n              <'
            'a:bodyPr/>\n              <a:lstStyle/>\n              <a:p/>\n '
            '           </a:txBody>\n            <a:tcPr/>\n          </a:tc>'
            '\n        </a:tr>\n        <a:tr h="223">\n          <a:tc>\n   '
            '         <a:txBody>\n              <a:bodyPr/>\n              <a'
            ':lstStyle/>\n              <a:p/>\n            </a:txBody>\n    '
            '        <a:tcPr/>\n          </a:tc>\n          <a:tc>\n        '
            '    <a:txBody>\n              <a:bodyPr/>\n              <a:lstS'
            'tyle/>\n              <a:p/>\n            </a:txBody>\n         '
            '   <a:tcPr/>\n          </a:tc>\n          <a:tc>\n            <'
            'a:txBody>\n              <a:bodyPr/>\n              <a:lstStyle/'
            '>\n              <a:p/>\n            </a:txBody>\n            <a'
            ':tcPr/>\n          </a:tc>\n        </a:tr>\n      </a:tbl>\n   '
            ' </a:graphicData>\n  </a:graphic>\n</p:graphicFrame>\n' %
            (nsdecls('a', 'p'), id_, name, left, top, width, height)
        )
        # exercise ---------------------
        graphicFrame = CT_GraphicalObjectFrame.new_table(
            id_, name, rows, cols, left, top, width, height)
        # verify -----------------------
        self.assertEqualLineByLine(xml, graphicFrame)
