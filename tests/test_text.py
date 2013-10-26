# encoding: utf-8

"""Test suite for pptx.text module."""

from __future__ import absolute_import

from hamcrest import assert_that, equal_to, is_, same_instance
from mock import MagicMock, Mock, patch

from pptx.constants import MSO, PP
from pptx.oxml import (
    _SubElement, nsdecls, oxml_fromstring, oxml_parse, oxml_tostring
)
from pptx.spec import namespaces
from pptx.text import _Font, _Paragraph, _Run, _TextFrame
from pptx.util import Pt

from .unitdata import test_text_objects, test_text_xml
from .unitutil import absjoin, TestCase, test_file_dir


nsmap = namespaces('a', 'r', 'p')


class Test_Font(TestCase):
    """Test _Font class"""
    def setUp(self):
        self.rPr_xml = '<a:rPr %s/>' % nsdecls('a')
        self.rPr = oxml_fromstring(self.rPr_xml)
        self.font = _Font(self.rPr)

    def test_get_bold_setting(self):
        """_Font.bold returns True on bold font weight"""
        # setup ------------------------
        rPr_xml = '<a:rPr %s b="1"/>' % nsdecls('a')
        rPr = oxml_fromstring(rPr_xml)
        font = _Font(rPr)
        # verify -----------------------
        assert_that(self.font.bold, is_(False))
        assert_that(font.bold, is_(True))

    def test_set_bold(self):
        """Setting _Font.bold to True selects bold font weight"""
        # setup ------------------------
        expected_rPr_xml = (
            '<a:rPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006'
            '/main" b="1"/>')
        # exercise ---------------------
        self.font.bold = True
        # verify -----------------------
        rPr_xml = oxml_tostring(self.font._rPr)
        assert_that(rPr_xml, is_(equal_to(expected_rPr_xml)))

    def test_clear_bold(self):
        """Setting _Font.bold to None clears run-level bold setting"""
        # setup ------------------------
        rPr_xml = (
            '<a:rPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006'
            '/main" b="1"/>')
        rPr = oxml_fromstring(rPr_xml)
        font = _Font(rPr)
        expected_rPr_xml = (
            '<a:rPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006'
            '/main"/>')
        # exercise ---------------------
        font.bold = None
        # verify -----------------------
        rPr_xml = oxml_tostring(font._rPr)
        assert_that(rPr_xml, is_(equal_to(expected_rPr_xml)))

    def test_set_font_size(self):
        """Assignment to _Font.size changes font size"""
        # setup ------------------------
        newfontsize = 2400
        expected_xml = (
            '<a:rPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006'
            '/main" sz="%d"/>') % newfontsize
        # exercise ---------------------
        self.font.size = newfontsize
        # verify -----------------------
        actual_xml = oxml_tostring(self.font._rPr)
        assert_that(actual_xml, is_(equal_to(expected_xml)))


class Test_Paragraph(TestCase):
    """Test _Paragraph"""
    def setUp(self):
        path = absjoin(test_file_dir, 'slide1.xml')
        self.sld = oxml_parse(path).getroot()
        xpath = './p:cSld/p:spTree/p:sp/p:txBody/a:p'
        self.pList = self.sld.xpath(xpath, namespaces=nsmap)

        self.test_text = 'test text'
        self.p_xml = ('<a:p %s><a:r><a:t>%s</a:t></a:r></a:p>' %
                      (nsdecls('a'), self.test_text))
        self.p = oxml_fromstring(self.p_xml)
        self.paragraph = _Paragraph(self.p)

    def test_runs_size(self):
        """_Paragraph.runs is expected size"""
        # setup ------------------------
        actual_lengths = []
        for p in self.pList:
            paragraph = _Paragraph(p)
            # exercise ----------------
            actual_lengths.append(len(paragraph.runs))
        # verify ------------------
        expected = [0, 0, 2, 1, 1, 1]
        actual = actual_lengths
        msg = "expected run count %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_add_run_increments_run_count(self):
        """_Paragraph.add_run() increments run count"""
        # setup ------------------------
        p_elm = self.pList[0]
        paragraph = _Paragraph(p_elm)
        # exercise ---------------------
        paragraph.add_run()
        # verify -----------------------
        expected = 1
        actual = len(paragraph.runs)
        msg = "expected run count %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    @patch('pptx.text.ParagraphAlignment')
    def test_alignment_value(self, ParagraphAlignment):
        """_Paragraph.alignment value is calculated correctly"""
        # setup ------------------------
        paragraph = test_text_objects.paragraph
        paragraph._Paragraph__p = __p = MagicMock(name='__p')
        __p.get_algn = get_algn = Mock(name='get_algn')
        get_algn.return_value = algn_val = Mock(name='algn_val')
        alignment = Mock(name='alignment')
        from_text_align_type = ParagraphAlignment.from_text_align_type
        from_text_align_type.return_value = alignment
        # exercise ---------------------
        retval = paragraph.alignment
        # verify -----------------------
        get_algn.assert_called_once_with()
        from_text_align_type.assert_called_once_with(algn_val)
        assert_that(retval, is_(same_instance(alignment)))

    @patch('pptx.text.ParagraphAlignment')
    def test_alignment_assignment(self, ParagraphAlignment):
        """Assignment to _Paragraph.alignment assigns value"""
        # setup ------------------------
        paragraph = test_text_objects.paragraph
        paragraph._Paragraph__p = __p = MagicMock(name='__p')
        __p.set_algn = set_algn = Mock(name='set_algn')
        algn_val = Mock(name='algn_val')
        to_text_align_type = ParagraphAlignment.to_text_align_type
        to_text_align_type.return_value = algn_val
        alignment = PP.ALIGN_CENTER
        # exercise ---------------------
        paragraph.alignment = alignment
        # verify -----------------------
        to_text_align_type.assert_called_once_with(alignment)
        set_algn.assert_called_once_with(algn_val)

    def test_alignment_integrates_with_CT_TextParagraph(self):
        """_Paragraph.alignment integrates with CT_TextParagraph"""
        # setup ------------------------
        paragraph = test_text_objects.paragraph
        expected_xml = test_text_xml.centered_paragraph
        # exercise ---------------------
        paragraph.alignment = PP.ALIGN_CENTER
        # verify -----------------------
        self.assertEqualLineByLine(expected_xml, paragraph._Paragraph__p)

    def test_clear_removes_all_runs(self):
        """_Paragraph.clear() removes all runs from paragraph"""
        # setup ------------------------
        p = self.pList[2]
        _SubElement(p, 'a:pPr')
        paragraph = _Paragraph(p)
        assert_that(len(paragraph.runs), is_(equal_to(2)))
        # exercise ---------------------
        paragraph.clear()
        # verify -----------------------
        assert_that(len(paragraph.runs), is_(equal_to(0)))

    def test_clear_preserves_paragraph_properties(self):
        """_Paragraph.clear() preserves paragraph properties"""
        # setup ------------------------
        p_xml = ('<a:p %s><a:pPr lvl="1"/><a:r><a:t>%s</a:t></a:r></a:p>' %
                 (nsdecls('a'), self.test_text))
        p_elm = oxml_fromstring(p_xml)
        paragraph = _Paragraph(p_elm)
        expected_p_xml = '<a:p %s><a:pPr lvl="1"/></a:p>' % nsdecls('a')
        # exercise ---------------------
        paragraph.clear()
        # verify -----------------------
        p_xml = oxml_tostring(paragraph._Paragraph__p)
        assert_that(p_xml, is_(equal_to(expected_p_xml)))

    def test_level_setter_generates_correct_xml(self):
        """_Paragraph.level setter generates correct XML"""
        # setup ------------------------
        expected_xml = (
            '<a:p %s>\n  <a:pPr lvl="2"/>\n  <a:r>\n    <a:t>test text</a:t>'
            '\n  </a:r>\n</a:p>\n' % nsdecls('a')
        )
        # exercise ---------------------
        self.paragraph.level = 2
        # verify -----------------------
        self.assertEqualLineByLine(expected_xml, self.paragraph._Paragraph__p)

    def test_level_default_is_zero(self):
        """_Paragraph.level defaults to zero on no lvl attribute"""
        # verify -----------------------
        assert_that(self.paragraph.level, is_(equal_to(0)))

    def test_level_roundtrips_intact(self):
        """_Paragraph.level property round-trips intact"""
        # exercise ---------------------
        self.paragraph.level = 5
        # verify -----------------------
        assert_that(self.paragraph.level, is_(equal_to(5)))

    def test_level_raises_on_bad_value(self):
        """_Paragraph.level raises on attempt to assign invalid value"""
        test_cases = ('0', -1, 9)
        for value in test_cases:
            with self.assertRaises(ValueError):
                self.paragraph.level = value

    def test_set_font_size(self):
        """Assignment to _Paragraph.font.size changes font size"""
        # setup ------------------------
        newfontsize = Pt(54.3)
        expected_xml = (
            '<a:p %s>\n  <a:pPr>\n    <a:defRPr sz="5430"/>\n  </a:pPr>\n  <a'
            ':r>\n    <a:t>test text</a:t>\n  </a:r>\n</a:p>\n' % nsdecls('a')
        )
        # exercise ---------------------
        self.paragraph.font.size = newfontsize
        # verify -----------------------
        self.assertEqualLineByLine(expected_xml, self.paragraph._Paragraph__p)

    def test_text_setter_sets_single_run_text(self):
        """assignment to _Paragraph.text creates single run containing value"""
        # setup ------------------------
        test_text = 'python-pptx was here!!'
        p_elm = self.pList[2]
        paragraph = _Paragraph(p_elm)
        # exercise ---------------------
        paragraph.text = test_text
        # verify -----------------------
        assert_that(len(paragraph.runs), is_(equal_to(1)))
        assert_that(paragraph.runs[0].text, is_(equal_to(test_text)))

    def test_text_accepts_non_ascii_strings(self):
        """assignment of non-ASCII string to text does not raise"""
        # setup ------------------------
        _7bit_string = 'String containing only 7-bit (ASCII) characters'
        _8bit_string = '8-bit string: Hér er texti með íslenskum stöfum.'
        _utf8_literal = u'unicode literal: Hér er texti með íslenskum stöfum.'
        _utf8_from_8bit = unicode('utf-8 unicode: Hér er texti', 'utf-8')
        # verify -----------------------
        try:
            text = _7bit_string
            self.paragraph.text = text
            text = _8bit_string
            self.paragraph.text = text
            text = _utf8_literal
            self.paragraph.text = text
            text = _utf8_from_8bit
            self.paragraph.text = text
        except ValueError:
            msg = "_Paragraph.text rejects valid text string '%s'" % text
            self.fail(msg)


class Test_Run(TestCase):
    """Test _Run"""
    def setUp(self):
        self.test_text = 'test text'
        self.r_xml = ('<a:r %s><a:t>%s</a:t></a:r>' %
                      (nsdecls('a'), self.test_text))
        self.r = oxml_fromstring(self.r_xml)
        self.run = _Run(self.r)

    def test_set_font_size(self):
        """Assignment to _Run.font.size changes font size"""
        # setup ------------------------
        newfontsize = 2400
        expected_xml = (
            '<a:r %s>\n  <a:rPr sz="2400"/>\n  <a:t>test text</a:t>\n</a:r>\n'
            % nsdecls('a')
        )
        # exercise ---------------------
        self.run.font.size = newfontsize
        # verify -----------------------
        self.assertEqualLineByLine(expected_xml, self.run._Run__r)

    def test_text_value(self):
        """_Run.text value is correct"""
        # exercise ---------------------
        text = self.run.text
        # verify -----------------------
        expected = self.test_text
        actual = text
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_text_setter(self):
        """_Run.text setter stores passed value"""
        # setup ------------------------
        new_value = 'new string'
        # exercise ---------------------
        self.run.text = new_value
        # verify -----------------------
        expected = new_value
        actual = self.run.text
        msg = "expected '%s', got '%s'" % (expected, actual)
        self.assertEqual(expected, actual, msg)


class Test_TextFrame(TestCase):
    """Test _TextFrame"""
    def setUp(self):
        path = absjoin(test_file_dir, 'slide1.xml')
        self.sld = oxml_parse(path).getroot()
        xpath = './p:cSld/p:spTree/p:sp/p:txBody'
        self.txBodyList = self.sld.xpath(xpath, namespaces=nsmap)

    def test_paragraphs_size(self):
        """_TextFrame.paragraphs is expected size"""
        # setup ------------------------
        actual_lengths = []
        for txBody in self.txBodyList:
            textframe = _TextFrame(txBody)
            # exercise ----------------
            actual_lengths.append(len(textframe.paragraphs))
        # verify -----------------------
        expected = [1, 1, 2, 1, 1]
        actual = actual_lengths
        msg = "expected paragraph count %s, got %s" % (expected, actual)
        self.assertEqual(expected, actual, msg)

    def test_add_paragraph_xml(self):
        """_TextFrame.add_paragraph does what it says"""
        # setup ------------------------
        txBody_xml = (
            '<p:txBody %s><a:bodyPr/><a:p><a:r><a:t>Test text</a:t></a:r></a:'
            'p></p:txBody>' % nsdecls('p', 'a')
        )
        expected_xml = (
            '<p:txBody %s><a:bodyPr/><a:p><a:r><a:t>Test text</a:t></a:r></a:'
            'p><a:p/></p:txBody>' % nsdecls('p', 'a')
        )
        txBody = oxml_fromstring(txBody_xml)
        textframe = _TextFrame(txBody)
        # exercise ---------------------
        textframe.add_paragraph()
        # verify -----------------------
        assert_that(len(textframe.paragraphs), is_(equal_to(2)))
        textframe_xml = oxml_tostring(textframe._txBody)
        expected = expected_xml
        actual = textframe_xml
        msg = "\nExpected: '%s'\n\n     Got: '%s'" % (expected, actual)
        if not expected == actual:
            raise AssertionError(msg)

    def test_text_setter_structure_and_value(self):
        """Assignment to _TextFrame.text yields single run para set to value"""
        # setup ------------------------
        test_text = 'python-pptx was here!!'
        txBody = self.txBodyList[2]
        textframe = _TextFrame(txBody)
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
        """Assignment to _TextFrame.vertical_anchor sets vert anchor"""
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
        txBody = oxml_fromstring(txBody_xml)
        textframe = _TextFrame(txBody)
        # exercise ---------------------
        textframe.vertical_anchor = MSO.ANCHOR_MIDDLE
        # verify -----------------------
        self.assertEqualLineByLine(expected_xml, textframe._txBody)

    def test_word_wrap_works(self):
        """Assignment to _TextFrame.word_wrap sets word wrap value"""
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

        txBody = oxml_fromstring(txBody_xml)
        textframe = _TextFrame(txBody)

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
