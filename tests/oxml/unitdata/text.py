# encoding: utf-8

"""Test data for unit tests"""

from pptx.oxml import parse_xml_bytes
from pptx.oxml.ns import nsdecls
from pptx.text import _Paragraph


class _TestTextXml(object):
    """XML snippets of text-related elements for use in unit tests"""
    @property
    def centered_paragraph(self):
        """
        XML for centered paragraph
        """
        return (
            '<a:p %s>\n'
            '  <a:pPr algn="ctr"/>\n'
            '</a:p>\n' % nsdecls('a')
        )

    @property
    def paragraph(self):
        """
        XML for a default, empty paragraph
        """
        return '<a:p %s/>\n' % nsdecls('a')


class _TestTextElements(object):
    """Text elements for use in unit tests"""
    @property
    def centered_paragraph(self):
        return parse_xml_bytes(test_text_xml.centered_paragraph)

    @property
    def paragraph(self):
        return parse_xml_bytes(test_text_xml.paragraph)


class _TestTextObjects(object):
    """Text object instances for use in unit tests"""
    @property
    def paragraph(self):
        return _Paragraph(test_text_elements.paragraph)


test_text_xml = _TestTextXml()
test_text_elements = _TestTextElements()
test_text_objects = _TestTextObjects()
