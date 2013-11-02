# encoding: utf-8

"""Test suite for pptx.oxml.presentation module."""

from __future__ import absolute_import, print_function, unicode_literals

from pptx.oxml import parse_xml_bytes
from pptx.oxml.ns import nsdecls
from pptx.oxml.presentation import CT_Presentation


class DescribeCT_Presentation(object):

    def it_is_used_by_the_parser_for_a_presentation_element(self):
        xml = ('<p:presentation %s/>' % nsdecls('p'))
        elm = parse_xml_bytes(xml)
        assert isinstance(elm, CT_Presentation)
