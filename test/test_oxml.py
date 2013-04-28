# -*- coding: utf-8 -*-
#
# test_oxml.py
#
# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

"""Test suite for pptx.oxml module."""

from hamcrest import assert_that, equal_to, is_
from unittest2 import TestCase

from pptx.oxml import (
    CT_GraphicalObjectFrame, CT_Picture, CT_Shape, CT_Table, oxml_tostring, qn
)
from pptx.spec import (
    PH_ORIENT_HORZ, PH_ORIENT_VERT, PH_SZ_FULL, PH_SZ_HALF, PH_SZ_QUARTER,
    PH_TYPE_CTRTITLE, PH_TYPE_DT, PH_TYPE_FTR, PH_TYPE_OBJ, PH_TYPE_SLDNUM,
    PH_TYPE_SUBTITLE, PH_TYPE_TBL
)


def _assert_equal_line_by_line(testcase, expected_xml, element):
    """
    Apply testcase.assertEqual() to each line of *expected_xml* and
    corresponding line of XML derived from *element*.
    """
    actual_xml = oxml_tostring(element, encoding='UTF-8',
                               pretty_print=True, standalone=True)
    actual_xml_lines = actual_xml.split('\n')
    expected_xml_lines = expected_xml.split('\n')
    for idx, line in enumerate(actual_xml_lines):
        msg = ("\n\nexpected:\n\n%s'\nbut got\n\n%s'" %
               (expected_xml, actual_xml))
        testcase.assertEqual(line, expected_xml_lines[idx], msg)


class TestCT_GraphicalObjectFrame(TestCase):
    """Test CT_GraphicalObjectFrame"""
    def test_has_table_return_value(self):
        """CT_GraphicalObjectFrame.has_table property has correct value"""
        # setup -----------------------
        id_, name = 9, 'Table 8'
        left, top, width, height = 111, 222, 333, 444
        tbl_uri = 'http://schemas.openxmlformats.org/drawingml/2006/table'
        chart_uri = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
        graphicFrame = CT_GraphicalObjectFrame.new_graphicFrame(
            id_, name, left, top, width, height)
        graphicData = graphicFrame[qn('a:graphic')].graphicData
        # verify ----------------------
        graphicData.set('uri', tbl_uri)
        assert_that(graphicFrame.has_table, is_(equal_to(True)))
        graphicData.set('uri', chart_uri)
        assert_that(graphicFrame.has_table, is_(equal_to(False)))

    def test_new_graphicFrame_generates_correct_xml(self):
        """CT_GraphicalObjectFrame.new_graphicFrame() returns correct XML"""
        # setup -----------------------
        id_, name = 9, 'Table 8'
        left, top, width, height = 111, 222, 333, 444
        xml = (
            '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>\n'
            '<p:graphicFrame xmlns:a="http://schemas.openxmlformats.org/drawi'
            'ngml/2006/main" xmlns:p="http://schemas.openxmlformats.org/prese'
            'ntationml/2006/main">\n  <p:nvGraphicFramePr>\n    <p:cNvPr id="'
            '%d" name="%s"/>\n    <p:cNvGraphicFramePr>\n      <a:graphicFram'
            'eLocks noGrp="1"/>\n    </p:cNvGraphicFramePr>\n    <p:nvPr/>\n '
            ' </p:nvGraphicFramePr>\n  <p:xfrm>\n    <a:off x="%d" y="%d"/>\n'
            '    <a:ext cx="%d" cy="%d"/>\n  </p:xfrm>\n  <a:graphic>\n    <a'
            ':graphicData/>\n  </a:graphic>\n</p:graphicFrame>\n' %
            (id_, name, left, top, width, height)
        )
        # exercise --------------------
        graphicFrame = CT_GraphicalObjectFrame.new_graphicFrame(
            id_, name, left, top, width, height)
        # verify ----------------------
        _assert_equal_line_by_line(self, xml, graphicFrame)

    def test_new_table_generates_correct_xml(self):
        """CT_GraphicalObjectFrame.new_table() returns correct XML"""
        # setup -----------------------
        id_, name = 9, 'Table 8'
        rows, cols = 2, 3
        left, top, width, height = 111, 222, 334, 445
        xml = (
            '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>\n'
            '<p:graphicFrame xmlns:a="http://schemas.openxmlformats.org/drawi'
            'ngml/2006/main" xmlns:p="http://schemas.openxmlformats.org/prese'
            'ntationml/2006/main">\n  <p:nvGraphicFramePr>\n    <p:cNvP''r id'
            '="%d" name="%s"/>\n    <p:cNvGraphicFramePr>\n      <a:graphicFr'
            'ameLocks noGrp="1"/>\n    </p:cNvGraphicFramePr>\n    <p:nvPr/>'
            '\n  </p:nvGraphicFramePr>\n  <p:xfrm>\n    <a:off x="%d" y="%d"/'
            '>\n    <a:ext cx="%d" cy="%d"/>\n  </p:xfrm>\n  <a:graphic>\n   '
            ' <a:graphicData uri="http://schemas.openxmlformats.org/drawingml'
            '/2006/table">\n      <a:tbl>\n        <a:tblPr firstRow="1" band'
            'Row="1">\n          <a:tableStyleId>{5C22544A-7EE6-4342-B048-85B'
            'DC9FD1C3A}</a:tableStyleId>\n        </a:tblPr>\n        <a:tblG'
            'rid>\n          <a:gridCol w="111"/>\n          <a:gridCol w="11'
            '1"/>\n          <a:gridCol w="112"/>\n        </a:tblGrid>\n    '
            '    <a:tr h="222">\n          <a:tc>\n            <a:txBody>\n  '
            '            <a:bodyPr/>\n              <a:lstStyle/>\n          '
            '    <a:p/>\n            </a:txBody>\n            <a:tcPr/>\n    '
            '      </a:tc>\n          <a:tc>\n            <a:txBody>\n       '
            '       <a:bodyPr/>\n              <a:lstStyle/>\n              <'
            'a:p/>\n            </a:txBody>\n            <a:tcPr/>\n         '
            ' </a:tc>\n          <a:tc>\n            <a:txBody>\n            '
            '  <a:bodyPr/>\n              <a:lstStyle/>\n              <a:p/>'
            '\n            </a:txBody>\n            <a:tcPr/>\n          </a:'
            'tc>\n        </a:tr>\n        <a:tr h="223">\n          <a:tc>\n'
            '            <a:txBody>\n              <a:bodyPr/>\n             '
            ' <a:lstStyle/>\n              <a:p/>\n            </a:txBody>\n '
            '           <a:tcPr/>\n          </a:tc>\n          <a:tc>\n     '
            '       <a:txBody>\n              <a:bodyPr/>\n              <a:l'
            'stStyle/>\n              <a:p/>\n            </a:txBody>\n      '
            '      <a:tcPr/>\n          </a:tc>\n          <a:tc>\n          '
            '  <a:txBody>\n              <a:bodyPr/>\n              <a:lstSty'
            'le/>\n              <a:p/>\n            </a:txBody>\n           '
            ' <a:tcPr/>\n          </a:tc>\n        </a:tr>\n      </a:tbl>\n'
            '    </a:graphicData>\n  </a:graphic>\n</p:graphicFrame>\n' %
            (id_, name, left, top, width, height)
        )
        # exercise --------------------
        graphicFrame = CT_GraphicalObjectFrame.new_table(
            id_, name, rows, cols, left, top, width, height)
        # verify ----------------------
        _assert_equal_line_by_line(self, xml, graphicFrame)


class TestCT_Picture(TestCase):
    """Test CT_Picture"""
    def test_new_pic_generates_correct_xml(self):
        """CT_Picture.new_pic() returns correct XML"""
        # setup -----------------------
        id_, name, desc, rId = 9, 'Picture 8', 'test-image.png', 'rId7'
        left, top, width, height = 111, 222, 333, 444
        xml = (
            '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>\n<'
            'p:pic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/'
            'main" xmlns:p="http://schemas.openxmlformats.org/presentationml/'
            '2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocu'
            'ment/2006/relationships">\n  <p:nvPicPr>\n    <p:cNvPr id="%d" n'
            'ame="%s" descr="%s"/>\n    <p:cNvPicPr>\n      <a:picLocks noCha'
            'ngeAspect="1"/>\n    </p:cNvPicPr>\n    <p:nvPr/>\n  </p:nvPicPr'
            '>\n  <p:blipFill>\n    <a:blip r:embed="%s"/>\n    <a:stretch>\n'
            '      <a:fillRect/>\n    </a:stretch>\n  </p:blipFill>\n  <p:spP'
            'r>\n    <a:xfrm>\n      <a:off x="%d" y="%d"/>\n      <a:ext cx='
            '"%d" cy="%d"/>\n    </a:xfrm>\n    <a:prstGeom prst="rect">\n   '
            '   <a:avLst/>\n    </a:prstGeom>\n  </p:spPr>\n</p:pic>\n' %
            (id_, name, desc, rId, left, top, width, height))
        # exercise --------------------
        pic = CT_Picture.new_pic(id_, name, desc, rId, left, top,
                                 width, height)
        # verify ----------------------
        _assert_equal_line_by_line(self, xml, pic)


class TestCT_Shape(TestCase):
    """Test CT_Shape"""
    def test_new_placeholder_sp_generates_correct_xml(self):
        """CT_Shape._new_placeholder_sp() returns correct XML"""
        # setup -----------------------
        expected_xml_tmpl = (
            '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>\n<'
            'p:sp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/m'
            'ain" xmlns:p="http://schemas.openxmlformats.org/presentationml/2'
            '006/main">\n  <p:nvSpPr>\n    <p:cNvPr id="%d" name="%s"/>\n    '
            '<p:cNvSpPr>\n      <a:spLocks noGrp="1"/>\n    </p:cNvSpPr>\n   '
            ' <p:nvPr>\n      <p:ph%s/>\n    </p:nvPr>\n  </p:nvSpPr>\n  <p:s'
            'pPr/>\n%s</p:sp>\n')
        txBody_snippet = (
            '  <p:txBody>\n    <a:bodyPr/>\n    <a:lstStyle/>\n    <a:p/>\n  '
            '</p:txBody>\n')
        test_cases = (
            (2, 'Title 1', PH_TYPE_CTRTITLE, PH_ORIENT_HORZ, PH_SZ_FULL,
             0),
            (3, 'Date Placeholder 2', PH_TYPE_DT, PH_ORIENT_HORZ, PH_SZ_HALF,
             10),
            (4, 'Vertical Subtitle 3', PH_TYPE_SUBTITLE, PH_ORIENT_VERT,
             PH_SZ_FULL, 1),
            (5, 'Table Placeholder 4', PH_TYPE_TBL, PH_ORIENT_HORZ,
             PH_SZ_QUARTER, 14),
            (6, 'Slide Number Placeholder 5', PH_TYPE_SLDNUM, PH_ORIENT_HORZ,
             PH_SZ_QUARTER, 12),
            (7, 'Footer Placeholder 6', PH_TYPE_FTR, PH_ORIENT_HORZ,
             PH_SZ_QUARTER, 11),
            (8, 'Content Placeholder 7', PH_TYPE_OBJ, PH_ORIENT_HORZ,
             PH_SZ_FULL, 15)
        )
        expected_values = (
            (2, 'Title 1', ' type="%s"' % PH_TYPE_CTRTITLE, txBody_snippet),
            (3, 'Date Placeholder 2', ' type="%s" sz="half" idx="10"' %
             PH_TYPE_DT, ''),
            (4, 'Vertical Subtitle 3', ' type="%s" orient="vert" idx="1"' %
             PH_TYPE_SUBTITLE, txBody_snippet),
            (5, 'Table Placeholder 4', ' type="%s" sz="quarter" idx="14"' %
             PH_TYPE_TBL, ''),
            (6, 'Slide Number Placeholder 5', ' type="%s" sz="quarter" '
             'idx="12"' % PH_TYPE_SLDNUM, ''),
            (7, 'Footer Placeholder 6', ' type="%s" sz="quarter" idx="11"' %
             PH_TYPE_FTR, ''),
            (8, 'Content Placeholder 7', ' idx="15"', txBody_snippet)
        )
        # exercise --------------------
        for case_idx, argv in enumerate(test_cases):
            id_, name, ph_type, orient, sz, idx = argv
            sp = CT_Shape.new_placeholder_sp(id_, name, ph_type, orient, sz,
                                             idx)
            # verify ------------------
            expected_xml = expected_xml_tmpl % expected_values[case_idx]
            _assert_equal_line_by_line(self, expected_xml, sp)

    def test_new_textbox_sp_generates_correct_xml(self):
        """CT_Shape._new_textbox_sp() returns correct XML"""
        # setup -----------------------
        id = 9
        name = 'TextBox 8'
        left, top, width, height = 111, 222, 333, 444
        xml = (
            '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>\n<'
            'p:sp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/m'
            'ain" xmlns:p="http://schemas.openxmlformats.org/presentationml/2'
            '006/main">\n  <p:nvSpPr>\n    <p:cNvPr id="%d" name="%s"/>\n    '
            '<p:cNvSpPr txBox="1"/>\n    <p:nvPr/>\n  </p:nvSpPr>\n  <p:spPr>'
            '\n    <a:xfrm>\n      <a:off x="%d" y="%d"/>\n      <a:ext cx="%'
            'd" cy="%d"/>\n    </a:xfrm>\n    <a:prstGeom prst="rect">\n     '
            ' <a:avLst/>\n    </a:prstGeom>\n    <a:noFill/>\n  </p:spPr>\n  '
            '<p:txBody>\n    <a:bodyPr wrap="none">\n      <a:spAutoFit/>\n  '
            '  </a:bodyPr>\n    <a:lstStyle/>\n    <a:p/>\n  </p:txBody>\n</p'
            ':sp>\n' % (id, name, left, top, width, height)
        )
        # exercise --------------------
        sp = CT_Shape.new_textbox_sp(id, name, left, top, width, height)
        # verify ----------------------
        _assert_equal_line_by_line(self, xml, sp)


class TestCT_Table(TestCase):
    """Test CT_Table"""
    def test_new_tbl_generates_correct_xml(self):
        """CT_Table._new_tbl() returns correct XML"""
        # setup -----------------------
        rows, cols = 2, 3
        width, height = 334, 445
        xml = (
            '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>\n<'
            'a:tbl xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/'
            'main">\n  <a:tblPr firstRow="1" bandRow="1">\n    <a:tableStyleI'
            'd>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>\n  </a'
            ':tblPr>\n  <a:tblGrid>\n    <a:gridCol w="111"/>\n    <a:gridCol'
            ' w="111"/>\n    <a:gridCol w="112"/>\n  </a:tblGrid>\n  <a:tr h='
            '"222">\n    <a:tc>\n      <a:txBody>\n        <a:bodyPr/>\n     '
            '   <a:lstStyle/>\n        <a:p/>\n      </a:txBody>\n      <a:tc'
            'Pr/>\n    </a:tc>\n    <a:tc>\n      <a:txBody>\n        <a:body'
            'Pr/>\n        <a:lstStyle/>\n        <a:p/>\n      </a:txBody>\n'
            '      <a:tcPr/>\n    </a:tc>\n    <a:tc>\n      <a:txBody>\n    '
            '    <a:bodyPr/>\n        <a:lstStyle/>\n        <a:p/>\n      </'
            'a:txBody>\n      <a:tcPr/>\n    </a:tc>\n  </a:tr>\n  <a:tr h="2'
            '23">\n    <a:tc>\n      <a:txBody>\n        <a:bodyPr/>\n       '
            ' <a:lstStyle/>\n        <a:p/>\n      </a:txBody>\n      <a:tcPr'
            '/>\n    </a:tc>\n    <a:tc>\n      <a:txBody>\n        <a:bodyPr'
            '/>\n        <a:lstStyle/>\n        <a:p/>\n      </a:txBody>\n  '
            '    <a:tcPr/>\n    </a:tc>\n    <a:tc>\n      <a:txBody>\n      '
            '  <a:bodyPr/>\n        <a:lstStyle/>\n        <a:p/>\n      </a:'
            'txBody>\n      <a:tcPr/>\n    </a:tc>\n  </a:tr>\n</a:tbl>\n'
        )
        # exercise --------------------
        tbl = CT_Table.new_tbl(rows, cols, width, height)
        # verify ----------------------
        _assert_equal_line_by_line(self, xml, tbl)
