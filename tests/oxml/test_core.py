# encoding: utf-8

"""
Test suite for pptx.oxml.core module.
"""

from __future__ import print_function, unicode_literals

import pytest

from lxml import objectify

from pptx.oxml.core import serialize_part_xml


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


# ===========================================================================
# fixtures
# ===========================================================================


@pytest.fixture
def part_elm(xml_bytes):
    return objectify.fromstring(xml_bytes)


@pytest.fixture
def expected_part_xml():
    return (
        '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\'?>\n'
        '<f:foo xmlns:py="http://codespeak.net/lxml/objectify/pytype" xmlns:'
        'f="http://foo" xmlns:b="http://bar"><f:bar>fØØbÅr</f:bar></f:foo>'
    ).encode('utf-8')


@pytest.fixture
def xml_bytes():
    xml_unicode = (
        '<f:foo xmlns:py="http://codespeak.net/lxml/objectify/pytype" xmlns:'
        'f="http://foo" xmlns:b="http://bar">\n'
        '  <f:bar py:pytype="str">fØØbÅr</f:bar>\n'
        '</f:foo>\n'
    )
    xml_bytes = xml_unicode.encode('utf-8')
    return xml_bytes
