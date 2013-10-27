# encoding: utf-8

"""Test suite for pptx.oxml.autoshape module."""

from __future__ import absolute_import

from hamcrest import assert_that, equal_to, instance_of, is_, none

from pptx.oxml.autoshape import CT_PresetGeometry2D, CT_Shape
from pptx.oxml.ns import nsdecls
from pptx.spec import (
    PH_ORIENT_HORZ, PH_ORIENT_VERT, PH_SZ_FULL, PH_SZ_HALF, PH_SZ_QUARTER,
    PH_TYPE_CTRTITLE, PH_TYPE_DT, PH_TYPE_FTR, PH_TYPE_OBJ, PH_TYPE_SLDNUM,
    PH_TYPE_SUBTITLE, PH_TYPE_TBL
)

from ..unitdata import a_prstGeom, test_shape_elements
from ..unitutil import TestCase


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
