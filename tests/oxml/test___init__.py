# encoding: utf-8

"""
Test suite for pptx.oxml.__init__.py module, primarily XML parser-related.
"""

from __future__ import print_function, unicode_literals

import pytest

from lxml import etree, objectify

from pptx.oxml import oxml_parser, register_custom_element_class


class DescribeOxmlParser(object):

    def it_enables_objectified_xml_parsing(self, xml_bytes):
        foo = objectify.fromstring(xml_bytes, oxml_parser)
        assert foo.bar == 'foobar'

    def it_strips_whitespace_between_elements(self, foo, stripped_xml_bytes):
        xml_bytes = etree.tostring(foo)
        assert xml_bytes == stripped_xml_bytes


class DescribeRegisterCustomElementClass(object):

    def it_determines_cust_elm_class_constructed_for_specified_tag(
            self, xml_bytes):
        register_custom_element_class('a:foo', CustElmCls)
        foo = objectify.fromstring(xml_bytes, oxml_parser)
        assert type(foo) is CustElmCls
        assert type(foo.bar) is objectify.StringElement


# ===========================================================================
# fixtures
# ===========================================================================

class CustElmCls(objectify.ObjectifiedElement):
    pass


@pytest.fixture
def foo(xml_bytes):
    return objectify.fromstring(xml_bytes, oxml_parser)


@pytest.fixture
def stripped_xml_bytes():
    return (
        '<a:foo xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/ma'
        'in"><a:bar>foobar</a:bar></a:foo>'
    ).encode('utf-8')


@pytest.fixture
def xml_bytes():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<a:foo xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/ma'
        'in">\n'
        '  <a:bar>foobar</a:bar>\n'
        '</a:foo>\n'
    ).encode('utf-8')
