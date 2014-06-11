# encoding: utf-8

"""
Test suite for pptx.oxml.core module.
"""

from __future__ import print_function, unicode_literals

import pytest

from pptx.oxml import parse_xml
from pptx.oxml.shared import serialize_part_xml


# class DescribeBaseOxmlElement(object):

#     def it_knows_which_tagnames_follow_a_given_child_tagname(
#             self, child_tagnames_fixture):
#         ElementClass, tagname, tagnames_after = child_tagnames_fixture
#         assert ElementClass.child_tagnames_after(tagname) == tagnames_after

#     # fixtures -------------------------------------------------------

#     @pytest.fixture(params=[
#         (('foo', 'bar', 'baz'), 'foo', ('bar', 'baz')),
#         ((('foo', 'bar'), 'baz'), 'foo', ('baz',)),
#         (('foo', ('bar', 'baz')), 'bar', ()),
#         (('1', '2', ('3', ('4', '5'), ('6', '7'))), '2',
#          ('3', '4', '5', '6', '7')),
#         (('1', '2', ('3', ('4', '5'), ('6', '7'))), '3', ()),
#     ])
#     def child_tagnames_fixture(self, request):
#         nested_sequence, tagname, tagnames_after = request.param
#         from pptx.oxml.shared import BaseOxmlElement, ChildTagnames

#         class ElementClass(BaseOxmlElement):
#             child_tagnames = ChildTagnames.from_nested_sequence(
#                 *nested_sequence
#             )

#         return ElementClass, tagname, tagnames_after


class DescribeSerializePartXml(object):

    def it_produces_properly_formatted_xml_for_an_opc_part(
            self, part_elm, expected_part_xml):
        """
        Tested aspects:
        ---------------
        * [X] it generates an XML declaration
        * [X] it produces no whitespace between elements
        * [X] it preserves unused namespaces
        * [X] it returns bytes ready to save to file (not unicode)
        """
        xml = serialize_part_xml(part_elm)
        assert xml == expected_part_xml
        # xml contains 134 chars, of which 3 are double-byte; it will have
        # len of 134 if it's unicode and 137 if it's bytes
        assert len(xml) == 137

    # fixtures -----------------------------------

    @pytest.fixture
    def part_elm(self):
        return parse_xml(
            '<f:foo xmlns:f="http://foo" xmlns:b="http://bar">\n  <f:bar>fØØ'
            'bÅr</f:bar>\n</f:foo>\n'
        )

    @pytest.fixture
    def expected_part_xml(self):
        unicode_xml = (
            '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>\n'
            '<f:foo xmlns:f="http://foo" xmlns:b="http://bar"><f:bar>fØØbÅr<'
            '/f:bar></f:foo>'
        )
        xml_bytes = unicode_xml.encode('utf-8')
        return xml_bytes
