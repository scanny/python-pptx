# encoding: utf-8

"""
Test suite for pptx.oxml.core module.
"""

from __future__ import print_function, unicode_literals

import pytest

from lxml import objectify

from pptx.oxml import oxml_parser
from pptx.oxml.shared import (
    BaseOxmlElement, child, ChildTagnames, Element, get_or_add,
    serialize_part_xml, SubElement
)
from pptx.oxml.ns import nsdecls, qn
from pptx.oxml.text import CT_TextBody


class DescribeBaseOxmlElement(object):

    def it_knows_which_tagnames_follow_a_given_child_tagname(
            self, child_tagnames_fixture):
        ElementClass, tagname, tagnames_after = child_tagnames_fixture
        assert ElementClass.child_tagnames_after(tagname) == tagnames_after

    # fixtures -------------------------------------------------------

    @pytest.fixture(params=[
        (('foo', 'bar', 'baz'), 'foo', ('bar', 'baz')),
        ((('foo', 'bar'), 'baz'), 'foo', ('baz',)),
        (('foo', ('bar', 'baz')), 'bar', ()),
        (('1', '2', ('3', ('4', '5'), ('6', '7'))), '2',
         ('3', '4', '5', '6', '7')),
        (('1', '2', ('3', ('4', '5'), ('6', '7'))), '3', ()),
    ])
    def child_tagnames_fixture(self, request):
        nested_sequence, tagname, tagnames_after = request.param

        class ElementClass(BaseOxmlElement):
            child_tagnames = ChildTagnames.from_nested_sequence(
                *nested_sequence
            )

        return ElementClass, tagname, tagnames_after


class DescribeChild(object):

    def it_returns_a_matching_child_if_present(
            self, parent_elm, known_child_nsptag_str, known_child_elm):
        child_elm = child(parent_elm, known_child_nsptag_str)
        assert child_elm is known_child_elm

    def it_returns_none_if_no_matching_child_is_present(self, parent_elm):
        child_elm = child(parent_elm, 'p:baz')
        assert child_elm is None


class DescribeElement(object):

    def it_returns_an_element_with_the_specified_tag(self, nsptag_str):
        elm = Element(nsptag_str)
        assert elm.tag == qn(nsptag_str)

    def it_returns_custom_element_class_if_one_is_defined(self, nsptag_str):
        elm = Element(nsptag_str)
        assert type(elm) is CT_TextBody


class DescribeGetOrAddChild(object):

    def it_returns_a_matching_child_if_present(
            self, parent_elm, known_child_nsptag_str, known_child_elm):
        child_elm = get_or_add(parent_elm, known_child_nsptag_str)
        assert child_elm is known_child_elm

    def it_creates_a_new_child_if_one_is_not_present(self, parent_elm):
        child_elm = get_or_add(parent_elm, 'p:baz')
        assert child_elm.tag == qn('p:baz')
        assert child_elm.getparent() is parent_elm


class DescribeSerializePartXml(object):

    def it_produces_properly_formatted_xml_for_an_opc_part(
            self, part_elm, expected_part_xml):
        """
        Tested aspects:
        ---------------
        * [X] it generates an XML declaration
        * [X] it produces no whitespace between elements
        * [X] it removes any annotations
        * [X] it preserves unused namespaces
        * [X] it returns bytes ready to save to file (not unicode)
        """
        xml = serialize_part_xml(part_elm)
        assert xml == expected_part_xml
        # xml contains 188 chars, of which 3 are double-byte; it will have
        # len of 188 if it's unicode and 191 if it's bytes
        assert len(xml) == 191

    # fixtures -----------------------------------

    @pytest.fixture
    def part_elm(self, xml_bytes):
        return objectify.fromstring(xml_bytes)

    @pytest.fixture
    def expected_part_xml(self):
        return (
            '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>\n'
            '<f:foo xmlns:py="http://codespeak.net/lxml/objectify/pytype" xm'
            'lns:f="http://foo" xmlns:b="http://bar"><f:bar>fØØbÅr</f:bar></'
            'f:foo>'
        ).encode('utf-8')

    @pytest.fixture
    def xml_bytes(self):
        xml_unicode = (
            '<f:foo xmlns:py="http://codespeak.net/lxml/objectify/pytype" xm'
            'lns:f="http://foo" xmlns:b="http://bar">\n'
            '  <f:bar py:pytype="str">fØØbÅr</f:bar>\n'
            '</f:foo>\n'
        )
        xml_bytes = xml_unicode.encode('utf-8')
        return xml_bytes


class DescribeSubElement(object):

    def it_returns_a_child_of_the_passed_parent_elm(
            self, parent_elm, nsptag_str):
        elm = SubElement(parent_elm, nsptag_str)
        assert elm.getparent() is parent_elm

    def it_returns_an_element_with_the_specified_tag(
            self, parent_elm, nsptag_str):
        elm = SubElement(parent_elm, nsptag_str)
        assert elm.tag == qn(nsptag_str)

    def it_returns_custom_element_class_if_one_is_defined_for_tag(
            self, parent_elm, nsptag_str):
        # note this behavior depends on the parser of parent_elm being the
        # one on which the custom element class lookups are defined
        elm = SubElement(parent_elm, nsptag_str)
        assert type(elm) is CT_TextBody

    def it_can_set_element_attributes(self, parent_elm, nsptag_str):
        attr_dct = {'foo': 'f', 'bar': 'b'}
        elm = SubElement(parent_elm, nsptag_str, attrib=attr_dct, baz='1')
        assert elm.get('foo') == 'f'
        assert elm.get('bar') == 'b'
        assert elm.get('baz') == '1'


# ===========================================================================
# fixtures
# ===========================================================================

@pytest.fixture
def known_child_elm(parent_elm, known_child_nsptag_str):
    return parent_elm[qn(known_child_nsptag_str)]


@pytest.fixture
def known_child_nsptag_str():
    return 'a:bar'


@pytest.fixture
def nsptag_str():
    return 'p:txBody'


@pytest.fixture
def parent_elm():
    xml = '<p:foo %s><a:bar>foobar</a:bar></p:foo>' % nsdecls('p', 'a')
    return objectify.fromstring(xml, oxml_parser)
