# encoding: utf-8

"""Test suite for tests.xmlbldr module."""

from __future__ import absolute_import, print_function, unicode_literals

from .oxml.unitdata.text import an_rPr


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
