# -*- coding: utf-8 -*-
#
# test_oxml.py
#
# Copyright (C) 2012, 2013 Steve Canny scanny@cisco.com
#
# This module is part of python-pptx and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

"""Test suite for pptx.oxml module."""

from hamcrest import assert_that, equal_to, instance_of, is_, none

from pptx.constants import (
    TEXT_ALIGN_TYPE as TAT, TEXT_ANCHORING_TYPE as TANC
)
from pptx.oxml import (
    CT_GraphicalObjectFrame, CT_Picture, CT_PresetGeometry2D, CT_Shape,
    CT_Table, nsdecls, qn
)
from pptx.spec import (
    PH_ORIENT_HORZ, PH_ORIENT_VERT, PH_SZ_FULL, PH_SZ_HALF, PH_SZ_QUARTER,
    PH_TYPE_CTRTITLE, PH_TYPE_DT, PH_TYPE_FTR, PH_TYPE_OBJ, PH_TYPE_SLDNUM,
    PH_TYPE_SUBTITLE, PH_TYPE_TBL
)

from testdata import (
    a_coreProperties, a_prstGeom, a_tbl, test_shape_elements,
    test_table_elements, test_table_xml, test_text_elements, test_text_xml
)
from testing import TestCase


class TestCT_CoreProperties(TestCase):
    """Test CT_CoreProperties"""
    def test_getter_values_match_xml(self):
        """CT_CoreProperties property values match parsed XML"""
        # setup ------------------------
        cases = (
            ('author',            'Creator'),
            ('category',          'Category'),
            ('comments',          'Description'),
            ('content_status',    'Content Status'),
            ('identifier',        'Identifier'),
            ('language',          'Language'),
            ('last_modified_by',  'Last Modified By'),
            ('subject',           'Subject'),
            ('title',             'Title'),
            ('version',           'Version'),
        )
        childless_core_prop_builder = a_coreProperties()
        # verify -----------------------
        for attr_name, value in cases:
            # string values should return empty string if element is missing
            childless_coreProperties = childless_core_prop_builder.element
            attr_value = getattr(childless_coreProperties, attr_name)
            reason = ("attr '%s' with missing element did not return ''" %
                      attr_name)
            assert_that(attr_value, is_(equal_to('')), reason)
            # and exact XML text otherwise
            builder = a_coreProperties().with_child(attr_name, value)
            coreProperties = builder.element
            attr_value = getattr(coreProperties, attr_name)
            xml = builder.xml
            reason = ("failed for property '%s' with this XML:\n\n%s" %
                      (attr_name, xml))
            assert_that(attr_value, is_(equal_to(value)), reason)

    def test_setters_produce_correct_xml(self):
        """Assignment to CT_CoreProperties properties produces correct XML"""
        # setup ------------------------
        cases = (
            ('author',            'Creator'),
            ('category',          'Category'),
            ('comments',          'Description'),
            ('content_status',    'Content Status'),
            ('identifier',        'Identifier'),
            ('language',          'Language'),
            ('last_modified_by',  'Last Modified By'),
            ('subject',           'Subject'),
            ('title',             'Title'),
            ('version',           'Version'),
        )
        # verify -----------------------
        for attr_name, value in cases:
            coreProperties = a_coreProperties().element  # no child elements
            # exercise ---------------------
            setattr(coreProperties, attr_name, value)
            # verify -----------------------
            expected_xml = a_coreProperties().with_child(attr_name, value).xml
            self.assertEqualLineByLine(expected_xml, coreProperties)


class TestCT_GraphicalObjectFrame(TestCase):
    """Test CT_GraphicalObjectFrame"""
    def test_has_table_return_value(self):
        """CT_GraphicalObjectFrame.has_table property has correct value"""
        # setup ------------------------
        id_, name = 9, 'Table 8'
        left, top, width, height = 111, 222, 333, 444
        tbl_uri = 'http://schemas.openxmlformats.org/drawingml/2006/table'
        chart_uri = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
        graphicFrame = CT_GraphicalObjectFrame.new_graphicFrame(
            id_, name, left, top, width, height)
        graphicData = graphicFrame[qn('a:graphic')].graphicData
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


class TestCT_PresetGeometry2D(TestCase):
    """Test CT_PresetGeometry2D"""
    def test_gd_return_value(self):
        """CT_PresetGeometry2D.gd value is correct"""
        # setup ------------------------
        long_prstGeom = a_prstGeom().with_gd(111, 'adj1')\
                                    .with_gd(222, 'adj2')\
                                    .with_gd(333, 'adj3')\
                                    .with_gd(444, 'adj4')
        cases = (
            (a_prstGeom(), ()),
            (a_prstGeom().with_gd(999, 'adj2'), ((999, 'adj2'),)),
            (long_prstGeom, ((111, 'adj1'), (222, 'adj2'), (333, 'adj3'),
                             (444, 'adj4'))),
        )
        for prstGeom_builder, expected_vals in cases:
            prstGeom = prstGeom_builder.element
            # exercise -----------------
            gd_elms = prstGeom.gd
            # verify -------------------
            assert_that(isinstance(gd_elms, tuple))
            assert_that(len(gd_elms), is_(equal_to(len(expected_vals))))
            for idx, gd_elm in enumerate(gd_elms):
                val, name = expected_vals[idx]
                fmla = 'val %d' % val
                assert_that(gd_elm.get('name'), is_(equal_to(name)))
                assert_that(gd_elm.get('fmla'), is_(equal_to(fmla)))

    def test_it_should_rewrite_guides_correctly(self):
        """CT_PresetGeometry2D.rewrite_guides() produces correct XML"""
        # setup ------------------------
        cases = (
            (a_prstGeom('chevron'), ()),
            (a_prstGeom('chevron').with_gd(111, 'foo'), (('bar', 222),)),
            (a_prstGeom('circularArrow').with_gd(333, 'adj1'),
             (('adj1', 12500), ('adj2', 1142319), ('adj3', 12500),
              ('adj4', 10800000), ('adj5', 12500))),
        )
        for prstGeom_builder, guides in cases:
            prstGeom = prstGeom_builder.element
            # exercise -----------------
            prstGeom.rewrite_guides(guides)
            # verify -------------------
            prstGeom_builder.reset()
            for name, val in guides:
                prstGeom_builder.with_gd(val, name)
            expected_xml = prstGeom_builder.with_avLst.xml
            self.assertEqualLineByLine(expected_xml, prstGeom)


class TestCT_Shape(TestCase):
    """Test CT_Shape"""
    def test_is_autoshape_distinguishes_auto_shape(self):
        """CT_Shape.is_autoshape distinguishes auto shape"""
        # setup ------------------------
        autoshape = test_shape_elements.autoshape
        placeholder = test_shape_elements.placeholder
        textbox = test_shape_elements.textbox
        # verify -----------------------
        assert_that(autoshape.is_autoshape, is_(True))
        assert_that(placeholder.is_autoshape, is_(False))
        assert_that(textbox.is_autoshape, is_(False))

    def test_is_placeholder_distinguishes_placeholder(self):
        """CT_Shape.is_autoshape distinguishes placeholder"""
        # setup ------------------------
        autoshape = test_shape_elements.autoshape
        placeholder = test_shape_elements.placeholder
        textbox = test_shape_elements.textbox
        # verify -----------------------
        assert_that(autoshape.is_autoshape, is_(True))
        assert_that(placeholder.is_autoshape, is_(False))
        assert_that(textbox.is_autoshape, is_(False))

    def test_is_textbox_distinguishes_text_box(self):
        """CT_Shape.is_textbox distinguishes text box"""
        # setup ------------------------
        autoshape = test_shape_elements.autoshape
        placeholder = test_shape_elements.placeholder
        textbox = test_shape_elements.textbox
        # verify -----------------------
        assert_that(autoshape.is_textbox, is_(False))
        assert_that(placeholder.is_textbox, is_(False))
        assert_that(textbox.is_textbox, is_(True))

    def test_new_autoshape_sp_generates_correct_xml(self):
        """CT_Shape._new_autoshape_sp() returns correct XML"""
        # setup ------------------------
        id_ = 9
        name = 'Rounded Rectangle 8'
        prst = 'roundRect'
        left, top, width, height = 111, 222, 333, 444
        xml = (
            '<p:sp %s>\n  <p:nvSpPr>\n    <p:cNvPr id="%d" name="%s"/>\n    <'
            'p:cNvSpPr/>\n    <p:nvPr/>\n  </p:nvSpPr>\n  <p:spPr>\n    <a:xf'
            'rm>\n      <a:off x="%d" y="%d"/>\n      <a:ext cx="%d" cy="%d"/'
            '>\n    </a:xfrm>\n    <a:prstGeom prst="%s">\n      <a:avLst/>\n'
            '    </a:prstGeom>\n  </p:spPr>\n  <p:style>\n    <a:lnRef idx="1'
            '">\n      <a:schemeClr val="accent1"/>\n    </a:lnRef>\n    <a:f'
            'illRef idx="3">\n      <a:schemeClr val="accent1"/>\n    </a:fil'
            'lRef>\n    <a:effectRef idx="2">\n      <a:schemeClr val="accent'
            '1"/>\n    </a:effectRef>\n    <a:fontRef idx="minor">\n      <a:'
            'schemeClr val="lt1"/>\n    </a:fontRef>\n  </p:style>\n  <p:txBo'
            'dy>\n    <a:bodyPr rtlCol="0" anchor="ctr"/>\n    <a:lstStyle/>'
            '\n    <a:p>\n      <a:pPr algn="ctr"/>\n    </a:p>\n  </p:txBody'
            '>\n</p:sp>\n' %
            (nsdecls('a', 'p'), id_, name, left, top, width, height, prst)
        )
        # exercise ---------------------
        sp = CT_Shape.new_autoshape_sp(id_, name, prst, left, top,
                                       width, height)
        # verify -----------------------
        self.assertEqualLineByLine(xml, sp)

    def test_new_placeholder_sp_generates_correct_xml(self):
        """CT_Shape._new_placeholder_sp() returns correct XML"""
        # setup ------------------------
        expected_xml_tmpl = (
            '<p:sp %s>\n  <p:nvSpPr>\n    <p:cNvPr id="%s" name="%s"/>\n    <'
            'p:cNvSpPr>\n      <a:spLocks noGrp="1"/>\n    </p:cNvSpPr>\n    '
            '<p:nvPr>\n      <p:ph%s/>\n    </p:nvPr>\n  </p:nvSpPr>\n  <p:sp'
            'Pr/>\n%s</p:sp>\n' % (nsdecls('a', 'p'), '%d', '%s', '%s', '%s')
        )
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
        # exercise ---------------------
        for case_idx, argv in enumerate(test_cases):
            id_, name, ph_type, orient, sz, idx = argv
            sp = CT_Shape.new_placeholder_sp(id_, name, ph_type, orient, sz,
                                             idx)
            # verify ------------------
            expected_xml = expected_xml_tmpl % expected_values[case_idx]
            self.assertEqualLineByLine(expected_xml, sp)

    def test_new_textbox_sp_generates_correct_xml(self):
        """CT_Shape.new_textbox_sp() returns correct XML"""
        # setup ------------------------
        id_ = 9
        name = 'TextBox 8'
        left, top, width, height = 111, 222, 333, 444
        xml = (
            '<p:sp %s>\n  <p:nvSpPr>\n    <p:cNvPr id="%d" name="%s"/>\n    <'
            'p:cNvSpPr txBox="1"/>\n    <p:nvPr/>\n  </p:nvSpPr>\n  <p:spPr>'
            '\n    <a:xfrm>\n      <a:off x="%d" y="%d"/>\n      <a:ext cx="%'
            'd" cy="%d"/>\n    </a:xfrm>\n    <a:prstGeom prst="rect">\n     '
            ' <a:avLst/>\n    </a:prstGeom>\n    <a:noFill/>\n  </p:spPr>\n  '
            '<p:txBody>\n    <a:bodyPr wrap="none">\n      <a:spAutoFit/>\n  '
            '  </a:bodyPr>\n    <a:lstStyle/>\n    <a:p/>\n  </p:txBody>\n</p'
            ':sp>\n' %
            (nsdecls('a', 'p'), id_, name, left, top, width, height)
        )
        # exercise ---------------------
        sp = CT_Shape.new_textbox_sp(id_, name, left, top, width, height)
        # verify -----------------------
        self.assertEqualLineByLine(xml, sp)

    def test_prst_return_value(self):
        """CT_Shape.prst value is correct"""
        # setup ------------------------
        rounded_rect_sp = test_shape_elements.rounded_rectangle
        placeholder_sp = test_shape_elements.placeholder
        # verify -----------------------
        assert_that(rounded_rect_sp.prst, is_(equal_to('roundRect')))
        assert_that(placeholder_sp.prst, is_(equal_to(None)))

    def test_prstGeom_return_value(self):
        """CT_Shape.prstGeom value is correct"""
        # setup ------------------------
        rounded_rect_sp = test_shape_elements.rounded_rectangle
        placeholder_sp = test_shape_elements.placeholder
        # verify -----------------------
        assert_that(rounded_rect_sp.prstGeom,
                    instance_of(CT_PresetGeometry2D))
        assert_that(placeholder_sp.prstGeom, is_(none()))


class TestCT_Table(TestCase):
    """Test CT_Table"""
    boolprops = ('bandRow', 'firstRow', 'lastRow',
                 'bandCol', 'firstCol', 'lastCol')

    def test_new_tbl_generates_correct_xml(self):
        """CT_Table._new_tbl() returns correct XML"""
        # setup ------------------------
        rows, cols = 2, 3
        width, height = 334, 445
        xml = (
            '<a:tbl %s>\n  <a:tblPr firstRow="1" bandRow="1">\n    <a:tableSt'
            'yleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>\n '
            ' </a:tblPr>\n  <a:tblGrid>\n    <a:gridCol w="111"/>\n    <a:gri'
            'dCol w="111"/>\n    <a:gridCol w="112"/>\n  </a:tblGrid>\n  <a:t'
            'r h="222">\n    <a:tc>\n      <a:txBody>\n        <a:bodyPr/>\n '
            '       <a:lstStyle/>\n        <a:p/>\n      </a:txBody>\n      <'
            'a:tcPr/>\n    </a:tc>\n    <a:tc>\n      <a:txBody>\n        <a:'
            'bodyPr/>\n        <a:lstStyle/>\n        <a:p/>\n      </a:txBod'
            'y>\n      <a:tcPr/>\n    </a:tc>\n    <a:tc>\n      <a:txBody>\n'
            '        <a:bodyPr/>\n        <a:lstStyle/>\n        <a:p/>\n    '
            '  </a:txBody>\n      <a:tcPr/>\n    </a:tc>\n  </a:tr>\n  <a:tr '
            'h="223">\n    <a:tc>\n      <a:txBody>\n        <a:bodyPr/>\n   '
            '     <a:lstStyle/>\n        <a:p/>\n      </a:txBody>\n      <a:'
            'tcPr/>\n    </a:tc>\n    <a:tc>\n      <a:txBody>\n        <a:bo'
            'dyPr/>\n        <a:lstStyle/>\n        <a:p/>\n      </a:txBody>'
            '\n      <a:tcPr/>\n    </a:tc>\n    <a:tc>\n      <a:txBody>\n  '
            '      <a:bodyPr/>\n        <a:lstStyle/>\n        <a:p/>\n      '
            '</a:txBody>\n      <a:tcPr/>\n    </a:tc>\n  </a:tr>\n</a:tbl>\n'
            % nsdecls('a')
        )
        # exercise ---------------------
        tbl = CT_Table.new_tbl(rows, cols, width, height)
        # verify -----------------------
        self.assertEqualLineByLine(xml, tbl)

    def test_boolean_property_value_is_correct(self):
        """CT_Table boolean property value is correct"""
        def getter_cases(propname):
            """Test cases for boolean property getter tests"""
            return (
                # defaults to False if no tblPr element present
                (a_tbl(), False),
                # defaults to False if tblPr element is empty
                (a_tbl().with_tblPr, False),
                # returns True if firstCol is valid truthy value
                (a_tbl().with_prop(propname, '1'), True),
                (a_tbl().with_prop(propname, 'true'), True),
                # returns False if firstCol has valid False value
                (a_tbl().with_prop(propname, '0'), False),
                (a_tbl().with_prop(propname, 'false'), False),
                # returns False if firstCol is not valid xsd:boolean value
                (a_tbl().with_prop(propname, 'foobar'), False),
            )
        for propname in self.boolprops:
            cases = getter_cases(propname)
            for tbl_builder, expected_property_value in cases:
                reason = (
                    'tbl.%s did not return %s for this XML:\n\n%s' %
                    (propname, expected_property_value, tbl_builder.xml)
                )
                assert_that(
                    getattr(tbl_builder.element, propname),
                    is_(equal_to(expected_property_value)),
                    reason
                )

    def test_assignment_to_boolean_property_produces_correct_xml(self):
        """Assignment to boolean property of CT_Table produces correct XML"""
        def xml_check_cases(propname):
            return (
                # => True: tblPr and attribute should be added
                (a_tbl(), True, a_tbl().with_prop(propname, '1')),
                # => False: attribute should be removed if false
                (a_tbl().with_prop(propname, '1'), False, a_tbl().with_tblPr),
                # => False: attribute should not be added if false
                (a_tbl(), False, a_tbl()),
            )
        for propname in self.boolprops:
            cases = xml_check_cases(propname)
            for tc_builder, assigned_value, expected_tc_builder in cases:
                tc = tc_builder.element
                setattr(tc, propname, assigned_value)
                self.assertEqualLineByLine(expected_tc_builder.xml, tc)


class TestCT_TableCell(TestCase):
    """Test CT_TableCell"""
    def test_anchor_property_value_is_correct(self):
        """CT_TableCell.anchor property value is correct"""
        # setup ------------------------
        cases = (
            (test_table_elements.cell, None),
            (test_table_elements.top_aligned_cell, TANC.TOP)
        )
        # verify -----------------------
        for tc, expected_text_anchoring_type in cases:
            assert_that(tc.anchor,
                        is_(equal_to(expected_text_anchoring_type)))

    def test_assignment_to_anchor_sets_anchor_value(self):
        """Assignment to CT_TableCell.anchor sets anchor value"""
        # setup ------------------------
        cases = (
            # something => something else
            (test_table_elements.top_aligned_cell, TANC.MIDDLE),
            # something => None
            (test_table_elements.top_aligned_cell, None),
            # None => something
            (test_table_elements.cell, TANC.BOTTOM),
            # None => None
            (test_table_elements.cell, None)
        )
        # verify -----------------------
        for tc, anchor in cases:
            tc.anchor = anchor
            assert_that(tc.anchor, is_(equal_to(anchor)))

    def test_assignment_to_anchor_produces_correct_xml(self):
        """Assigning value to CT_TableCell.anchor produces correct XML"""
        # setup ------------------------
        cases = (
            # None => something
            (test_table_elements.cell, TANC.TOP,
             test_table_xml.top_aligned_cell),
            # something => None
            (test_table_elements.top_aligned_cell, None,
             test_table_xml.cell)
        )
        # verify -----------------------
        for tc, text_anchoring_type, expected_xml in cases:
            tc.anchor = text_anchoring_type
            self.assertEqualLineByLine(expected_xml, tc)

    def test_marX_property_values_are_correct(self):
        """CT_TableCell.marX property values are correct"""
        # setup ------------------------
        cases = (
            (test_table_elements.cell_with_margins, 12, 34, 56, 78),
            (test_table_elements.cell, 45720, 91440, 45720, 91440)
        )
        # verify -----------------------
        for tc, exp_marT, exp_marR, exp_marB, exp_marL in cases:
            assert_that(tc.marT, is_(equal_to(exp_marT)))
            assert_that(tc.marR, is_(equal_to(exp_marR)))
            assert_that(tc.marB, is_(equal_to(exp_marB)))
            assert_that(tc.marL, is_(equal_to(exp_marL)))

    def test_assignment_to_marX_sets_value(self):
        """Assignment to CT_TableCell.marX sets marX value"""
        # setup ------------------------
        cases = (
            # something => something else
            (
                test_table_elements.cell_with_margins,
                (98, 76, 54, 32),
                (98, 76, 54, 32)
            ),
            # something => None
            (
                test_table_elements.cell_with_margins,
                (None, None, None, None),
                (45720, 91440, 45720, 91440)
            ),
            # None => something
            (
                test_table_elements.cell,
                (98, 76, 54, 32),
                (98, 76, 54, 32)
            ),
            # None => None
            (
                test_table_elements.cell,
                (None, None, None, None),
                (45720, 91440, 45720, 91440)
            )
        )
        # verify -----------------------
        for tc, marX, expected_marX in cases:
            tc.marT, tc.marR, tc.marB, tc.marL = marX
            exp_marT, exp_marR, exp_marB, exp_marL = expected_marX
            assert_that(tc.marT, is_(equal_to(exp_marT)))
            assert_that(tc.marR, is_(equal_to(exp_marR)))
            assert_that(tc.marB, is_(equal_to(exp_marB)))
            assert_that(tc.marL, is_(equal_to(exp_marL)))

    def test_assignment_to_marX_produces_correct_xml(self):
        """Assigning value to CT_TableCell.marX produces correct XML"""
        # setup ------------------------
        cases = (
            # None => something
            (
                test_table_elements.cell,
                (12, 34, 56, 78),
                test_table_xml.cell_with_margins
            ),
            # something => None
            (
                test_table_elements.cell_with_margins,
                (None, None, None, None),
                test_table_xml.cell
            )
        )
        # verify -----------------------
        for tc, marX, expected_xml in cases:
            tc.marT, tc.marR, tc.marB, tc.marL = marX
            self.assertEqualLineByLine(expected_xml, tc)


class TestCT_TextParagraph(TestCase):
    """Test CT_TextParagraph"""
    def test_get_algn_returns_correct_value(self):
        """CT_TextParagraph.get_algn() returns correct value"""
        # setup ------------------------
        cases = (
            (test_text_elements.paragraph, None),
            (test_text_elements.centered_paragraph, TAT.CENTER)
        )
        # verify -----------------------
        for p, expected_algn in cases:
            assert_that(p.get_algn(), is_(equal_to(expected_algn)))

    def test_set_algn_sets_algn_value(self):
        """CT_TextParagraph.set_algn() sets algn value"""
        # setup ------------------------
        cases = (
            # something => something else
            (test_text_elements.centered_paragraph, TAT.JUSTIFY),
            # something => None
            (test_text_elements.centered_paragraph, None),
            # None => something
            (test_text_elements.paragraph, TAT.CENTER),
            # None => None
            (test_text_elements.paragraph, None)
        )
        # verify -----------------------
        for p, algn in cases:
            p.set_algn(algn)
            assert_that(p.get_algn(), is_(equal_to(algn)))

    def test_set_algn_produces_correct_xml(self):
        """Assigning value to CT_TextParagraph.algn produces correct XML"""
        # setup ------------------------
        cases = (
            # None => something
            (test_text_elements.paragraph, TAT.CENTER,
             test_text_xml.centered_paragraph),
            # something => None
            (test_text_elements.centered_paragraph, None,
             test_text_xml.paragraph)
        )
        # verify -----------------------
        for p, text_align_type, expected_xml in cases:
            p.set_algn(text_align_type)
            self.assertEqualLineByLine(expected_xml, p)
