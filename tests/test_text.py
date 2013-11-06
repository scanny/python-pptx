# encoding: utf-8

"""Test suite for pptx.text module."""

from __future__ import absolute_import

import pytest

from hamcrest import assert_that, equal_to, is_

from pptx.constants import MSO, PP
from pptx.oxml import parse_xml_bytes
from pptx.oxml.core import SubElement
from pptx.oxml.ns import namespaces, nsdecls
from pptx.text import _Font, _Paragraph, _Run, TextFrame
from pptx.util import Pt

from .oxml.unitdata.text import a_p, a_pPr, a_t, an_r, an_rPr
from .unitutil import (
    absjoin, actual_xml, parse_xml_file, serialize_xml, TestCase,
    test_file_dir
)


nsmap = namespaces('a', 'r', 'p')


class Describe_Font(object):

    def it_knows_the_bold_setting(self, font, bold_font, bold_off_font):
        assert font.bold is None
        assert bold_font.bold is True
        assert bold_off_font.bold is False

    def it_can_change_the_bold_setting(
            self, font, bold_rPr_xml, bold_off_rPr_xml, rPr_xml):
        assert actual_xml(font._rPr) == rPr_xml
        font.bold = True
        assert actual_xml(font._rPr) == bold_rPr_xml
        font.bold = False
        assert actual_xml(font._rPr) == bold_off_rPr_xml
        font.bold = None
        assert actual_xml(font._rPr) == rPr_xml

    def it_can_set_the_font_size(self, font):
        font.size = 2400
        expected_xml = an_rPr().with_nsdecls().with_sz(2400).xml()
        assert actual_xml(font._rPr) == expected_xml

    # fixtures ---------------------------------------------

    @pytest.fixture
    def bold_font(self, bold_rPr):
        return _Font(bold_rPr)

    @pytest.fixture
    def bold_off_font(self, bold_off_rPr):
        return _Font(bold_off_rPr)

    @pytest.fixture
    def bold_off_rPr(self, bold_off_rPr_bldr):
        return bold_off_rPr_bldr.element

    @pytest.fixture
    def bold_off_rPr_bldr(self):
        return an_rPr().with_nsdecls().with_b(0)

    @pytest.fixture
    def bold_off_rPr_xml(self, bold_off_rPr_bldr):
        return bold_off_rPr_bldr.xml()

    @pytest.fixture
    def bold_rPr(self, bold_rPr_bldr):
        return bold_rPr_bldr.element

    @pytest.fixture
    def bold_rPr_bldr(self):
        return an_rPr().with_nsdecls().with_b(1)

    @pytest.fixture
    def bold_rPr_xml(self, bold_rPr_bldr):
        return bold_rPr_bldr.xml()

    @pytest.fixture
    def rPr(self, rPr_bldr):
        return rPr_bldr.element

    @pytest.fixture
    def rPr_bldr(self):
        return an_rPr().with_nsdecls()

    @pytest.fixture
    def rPr_xml(self, rPr_bldr):
        return rPr_bldr.xml()

    @pytest.fixture
    def font(self, rPr):
        return _Font(rPr)


class Describe_Paragraph(object):

    def it_can_add_a_run(self, paragraph, p_with_r_xml):
        run = paragraph.add_run()
        assert actual_xml(paragraph._p) == p_with_r_xml
        assert isinstance(run, _Run)

    def it_knows_the_alignment_setting_of_the_paragraph(
            self, paragraph, paragraph_with_algn):
        assert paragraph.alignment is None
        assert paragraph_with_algn.alignment == PP.ALIGN_CENTER

    def it_can_change_its_alignment_setting(self, paragraph):
        paragraph.alignment = PP.ALIGN_LEFT
        assert paragraph._pPr.algn == 'l'
        paragraph.alignment = None
        assert paragraph._pPr.algn is None

    def test_clear_removes_all_runs(self, pList):
        """_Paragraph.clear() removes all runs from paragraph"""
        # setup ------------------------
        p = pList[2]
        SubElement(p, 'a:pPr')
        paragraph = _Paragraph(p)
        assert_that(len(paragraph.runs), is_(equal_to(2)))
        # exercise ---------------------
        paragraph.clear()
        # verify -----------------------
        assert len(paragraph.runs) == 0

    def test_clear_preserves_paragraph_properties(self, test_text):
        """_Paragraph.clear() preserves paragraph properties"""
        # setup ------------------------
        p_xml = ('<a:p %s><a:pPr lvl="1"/><a:r><a:t>%s</a:t></a:r></a:p>' %
                 (nsdecls('a'), test_text))
        p_elm = parse_xml_bytes(p_xml)
        paragraph = _Paragraph(p_elm)
        expected_p_xml = '<a:p %s><a:pPr lvl="1"/></a:p>' % nsdecls('a')
        # exercise ---------------------
        paragraph.clear()
        # verify -----------------------
        assert serialize_xml(paragraph._p) == expected_p_xml

    def test_set_font_size(self, paragraph_with_text):
        """Assignment to _Paragraph.font.size changes font size"""
        # setup ------------------------
        newfontsize = Pt(54.3)
        expected_xml = (
            '<a:p %s>\n  <a:pPr>\n    <a:defRPr sz="5430"/>\n  </a:pPr>\n  <a'
            ':r>\n    <a:t>test text</a:t>\n  </a:r>\n</a:p>\n' % nsdecls('a')
        )
        # exercise ---------------------
        paragraph_with_text.font.size = newfontsize
        # verify -----------------------
        assert actual_xml(paragraph_with_text._p) == expected_xml

    def test_level_setter_generates_correct_xml(self, paragraph_with_text):
        """_Paragraph.level setter generates correct XML"""
        # setup ------------------------
        expected_xml = (
            '<a:p %s>\n  <a:pPr lvl="2"/>\n  <a:r>\n    <a:t>test text</a:t>'
            '\n  </a:r>\n</a:p>\n' % nsdecls('a')
        )
        # exercise ---------------------
        paragraph_with_text.level = 2
        # verify -----------------------
        assert actual_xml(paragraph_with_text._p) == expected_xml

    def test_level_default_is_zero(self, paragraph_with_text):
        """_Paragraph.level defaults to zero on no lvl attribute"""
        # verify -----------------------
        assert paragraph_with_text.level == 0

    def test_level_roundtrips_intact(self, paragraph_with_text):
        """_Paragraph.level property round-trips intact"""
        # exercise ---------------------
        paragraph_with_text.level = 5
        # verify -----------------------
        assert paragraph_with_text.level == 5

    def test_level_raises_on_bad_value(self, paragraph_with_text):
        """_Paragraph.level raises on attempt to assign invalid value"""
        test_cases = ('0', -1, 9)
        for value in test_cases:
            with pytest.raises(ValueError):
                paragraph_with_text.level = value

    def test_runs_size(self, pList):
        """_Paragraph.runs is expected size"""
        # setup ------------------------
        actual_lengths = []
        for p in pList:
            paragraph = _Paragraph(p)
            # exercise ----------------
            actual_lengths.append(len(paragraph.runs))
        # verify ------------------
        expected = [0, 0, 2, 1, 1, 1]
        actual = actual_lengths
        msg = "expected run count %s, got %s" % (expected, actual)
        assert actual == expected, msg

    def test_text_setter_sets_single_run_text(self, pList):
        """assignment to _Paragraph.text creates single run containing value"""
        # setup ------------------------
        test_text = 'python-pptx was here!!'
        p_elm = pList[2]
        paragraph = _Paragraph(p_elm)
        # exercise ---------------------
        paragraph.text = test_text
        # verify -----------------------
        assert len(paragraph.runs) == 1
        assert paragraph.runs[0].text == test_text

    def test_text_accepts_non_ascii_strings(self, paragraph_with_text):
        """assignment of non-ASCII string to text does not raise"""
        # setup ------------------------
        _7bit_string = 'String containing only 7-bit (ASCII) characters'
        _8bit_string = '8-bit string: Hér er texti með íslenskum stöfum.'
        _utf8_literal = u'unicode literal: Hér er texti með íslenskum stöfum.'
        _utf8_from_8bit = unicode('utf-8 unicode: Hér er texti', 'utf-8')
        # verify -----------------------
        try:
            text = _7bit_string
            paragraph_with_text.text = text
            text = _8bit_string
            paragraph_with_text.text = text
            text = _utf8_literal
            paragraph_with_text.text = text
            text = _utf8_from_8bit
            paragraph_with_text.text = text
        except ValueError:
            msg = "_Paragraph.text rejects valid text string '%s'" % text
            pytest.fail(msg)

    # fixtures ---------------------------------------------

    @pytest.fixture
    def pList(self, sld, xpath):
        return sld.xpath(xpath, namespaces=nsmap)

    @pytest.fixture
    def p_bldr(self):
        return a_p().with_nsdecls()

    @pytest.fixture
    def p_with_r_xml(self):
        run_bldr = an_r().with_child(a_t())
        return a_p().with_nsdecls().with_child(run_bldr).xml()

    @pytest.fixture
    def p_with_text(self, p_with_text_xml):
        return parse_xml_bytes(p_with_text_xml)

    @pytest.fixture
    def p_with_text_xml(self, test_text):
        return ('<a:p %s><a:r><a:t>%s</a:t></a:r></a:p>' %
                (nsdecls('a'), test_text))

    @pytest.fixture
    def paragraph(self, p_bldr):
        return _Paragraph(p_bldr.element)

    @pytest.fixture
    def paragraph_with_algn(self):
        pPr_bldr = a_pPr().with_algn('ctr')
        p_bldr = a_p().with_nsdecls().with_child(pPr_bldr)
        return _Paragraph(p_bldr.element)

    @pytest.fixture
    def paragraph_with_text(self, p_with_text):
        return _Paragraph(p_with_text)

    @pytest.fixture
    def path(self):
        return absjoin(test_file_dir, 'slide1.xml')

    @pytest.fixture
    def sld(self, path):
        return parse_xml_file(path).getroot()

    @pytest.fixture
    def test_text(self):
        return 'test text'

    @pytest.fixture
    def xpath(self):
        return './p:cSld/p:spTree/p:sp/p:txBody/a:p'


class Describe_Run(object):

    def it_can_get_the_text_of_the_run(self, run, test_text):
        assert run.text == test_text

    def it_can_change_the_text_of_the_run(self, run):
        run.text = 'new text'
        assert run.text == 'new text'

    # fixtures ---------------------------------------------

    @pytest.fixture
    def test_text(self):
        return 'test text'

    @pytest.fixture
    def r_xml(self, test_text):
        return ('<a:r %s><a:t>%s</a:t></a:r>' %
                (nsdecls('a'), test_text))

    @pytest.fixture
    def r(self, r_xml):
        return parse_xml_bytes(r_xml)

    @pytest.fixture
    def run(self, r):
        return _Run(r)


class DescribeTextFrame(TestCase):

    def setUp(self):
        path = absjoin(test_file_dir, 'slide1.xml')
        self.sld = parse_xml_file(path).getroot()
        xpath = './p:cSld/p:spTree/p:sp/p:txBody'
        self.txBodyList = self.sld.xpath(xpath, namespaces=nsmap)

    def test_paragraphs_size(self):
        """TextFrame.paragraphs is expected size"""
        # setup ------------------------
        actual_lengths = []
        for txBody in self.txBodyList:
            textframe = TextFrame(txBody)
            # exercise ----------------
            actual_lengths.append(len(textframe.paragraphs))
        # verify -----------------------
        expected = [1, 1, 2, 1, 1]
        actual = actual_lengths
        msg = "expected paragraph count %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_add_paragraph_xml(self):
        """TextFrame.add_paragraph does what it says"""
        # setup ------------------------
        txBody_xml = (
            '<p:txBody %s><a:bodyPr/><a:p><a:r><a:t>Test text</a:t></a:r></a:'
            'p></p:txBody>' % nsdecls('p', 'a')
        )
        expected_xml = (
            '<p:txBody %s><a:bodyPr/><a:p><a:r><a:t>Test text</a:t></a:r></a:'
            'p><a:p/></p:txBody>' % nsdecls('p', 'a')
        )
        txBody = parse_xml_bytes(txBody_xml)
        textframe = TextFrame(txBody)
        # exercise ---------------------
        textframe.add_paragraph()
        # verify -----------------------
        assert_that(len(textframe.paragraphs), is_(equal_to(2)))
        textframe_xml = serialize_xml(textframe._txBody)
        expected = expected_xml
        actual = textframe_xml
        msg = "\nExpected: '%s'\n\n     Got: '%s'" % (expected, actual)
        if not expected == actual:
            raise AssertionError(msg)

    def test_text_setter_structure_and_value(self):
        """Assignment to TextFrame.text yields single run para set to value"""
        # setup ------------------------
        test_text = 'python-pptx was here!!'
        txBody = self.txBodyList[2]
        textframe = TextFrame(txBody)
        # exercise ---------------------
        textframe.text = test_text
        # verify paragraph count -------
        expected = 1
        actual = len(textframe.paragraphs)
        msg = "expected paragraph count %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)
        # verify value -----------------
        expected = test_text
        actual = textframe.paragraphs[0].runs[0].text
        msg = "expected text '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_vertical_anchor_works(self):
        """Assignment to TextFrame.vertical_anchor sets vert anchor"""
        # setup ------------------------
        txBody_xml = (
            '<p:txBody %s><a:bodyPr/><a:p><a:r><a:t>Test text</a:t></a:r></a:'
            'p></p:txBody>' % nsdecls('p', 'a')
        )
        expected_xml = (
            '<p:txBody %s>\n  <a:bodyPr anchor="ctr"/>\n  <a:p>\n    <a:r>\n '
            '     <a:t>Test text</a:t>\n    </a:r>\n  </a:p>\n</p:txBody>\n' %
            nsdecls('p', 'a')
        )
        txBody = parse_xml_bytes(txBody_xml)
        textframe = TextFrame(txBody)
        # exercise ---------------------
        textframe.vertical_anchor = MSO.ANCHOR_MIDDLE
        # verify -----------------------
        self.assertEqualLineByLine(expected_xml, textframe._txBody)

    def test_word_wrap_works(self):
        """Assignment to TextFrame.word_wrap sets word wrap value"""
        # setup ------------------------
        txBody_xml = (
            '<p:txBody %s><a:bodyPr/><a:p><a:r><a:t>Test text</a:t></a:r></a:'
            'p></p:txBody>' % nsdecls('p', 'a')
        )
        true_expected_xml = (
            '<p:txBody %s>\n  <a:bodyPr wrap="square"/>\n  <a:p>\n    <a:r>\n '
            '     <a:t>Test text</a:t>\n    </a:r>\n  </a:p>\n</p:txBody>\n' %
            nsdecls('p', 'a')
        )
        false_expected_xml = (
            '<p:txBody %s>\n  <a:bodyPr wrap="none"/>\n  <a:p>\n    <a:r>\n '
            '     <a:t>Test text</a:t>\n    </a:r>\n  </a:p>\n</p:txBody>\n' %
            nsdecls('p', 'a')
        )
        none_expected_xml = (
            '<p:txBody %s>\n  <a:bodyPr/>\n  <a:p>\n    <a:r>\n '
            '     <a:t>Test text</a:t>\n    </a:r>\n  </a:p>\n</p:txBody>\n' %
            nsdecls('p', 'a')
        )

        txBody = parse_xml_bytes(txBody_xml)
        textframe = TextFrame(txBody)

        self.assertEqual(textframe.word_wrap, None)

        # exercise ---------------------
        textframe.word_wrap = True
        # verify -----------------------
        self.assertEqualLineByLine(
            true_expected_xml, textframe._txBody)
        self.assertEqual(textframe.word_wrap, True)

        # exercise ---------------------
        textframe.word_wrap = False
        # verify -----------------------
        self.assertEqualLineByLine(
            false_expected_xml, textframe._txBody)
        self.assertEqual(textframe.word_wrap, False)

        # exercise ---------------------
        textframe.word_wrap = None
        # verify -----------------------
        self.assertEqualLineByLine(
            none_expected_xml, textframe._txBody)
        self.assertEqual(textframe.word_wrap, None)
