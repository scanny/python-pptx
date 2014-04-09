# encoding: utf-8

"""
Test suite for tests.xmlbldr module
"""

from __future__ import absolute_import, print_function, unicode_literals

from pptx.enum.text import PP_ALIGN

from .oxml.unitdata.text import a_p, a_pPr, an_rPr


class DescribeCT_TextParagraphBuilder(object):

    def it_can_build_an_empty_p_element(self):
        p_bldr = a_p()
        assert p_bldr.xml() == '<a:p/>\n'

    def it_can_include_a_pPr_child_element(self):
        pPr_bldr = a_pPr()
        p_bldr = a_p().with_child(pPr_bldr)
        expected_xml = (
            '<a:p>\n'
            '  <a:pPr/>\n'
            '</a:p>\n'
        )
        assert p_bldr.xml() == expected_xml


class DescribeCT_TextParagraphPropertiesBuilder(object):

    def it_can_build_an_empty_rPr_element(self):
        pPr_bldr = a_pPr()
        assert pPr_bldr.xml() == '<a:pPr/>\n'

    def it_can_add_an_algn_attribute(self):
        pPr_bldr = a_pPr().with_algn(PP_ALIGN.to_xml(PP_ALIGN.CENTER))
        assert pPr_bldr.xml() == '<a:pPr algn="ctr"/>\n'

    def it_can_add_a_lvl_attribute(self):
        pPr_bldr = a_pPr().with_lvl(2)
        assert pPr_bldr.xml() == '<a:pPr lvl="2"/>\n'


class DescribeCT_TextCharacterPropertiesBuilder(object):

    def it_can_format_xml_text_for_an_empty_rPr_element(self):
        rPr_bldr = an_rPr()
        assert rPr_bldr.xml() == '<a:rPr/>\n'

    def it_can_format_xml_text_for_an_rPr_element_with_nsdecls(self):
        rPr_bldr = an_rPr().with_nsdecls()
        expected_xml = (
            '<a:rPr xmlns:a="http://schemas.openxmlformats.org/drawingml/200'
            '6/main"/>\n'
        )
        assert rPr_bldr.xml() == expected_xml

    def it_can_format_an_rPr_element_with_a_b_attribute(self):
        rPr_bldr = an_rPr().with_b(1)
        assert rPr_bldr.xml() == '<a:rPr b="1"/>\n'

    def it_can_format_an_rPr_element_with_a_i_attribute(self):
        rPr_bldr = an_rPr().with_i(1)
        assert rPr_bldr.xml() == '<a:rPr i="1"/>\n'

    def it_can_format_xml_text_for_an_rPr_element_with_everything(self):
        rPr_bldr = an_rPr().with_nsdecls().with_b(0).with_i(1)
        expected_xml = (
            '<a:rPr xmlns:a="http://schemas.openxmlformats.org/drawingml/200'
            '6/main" b="0" i="1"/>\n'
        )
        assert rPr_bldr.xml() == expected_xml
