"""Test suite for pptx.oxml.__init__.py module, primarily XML parser-related."""

from __future__ import annotations

import pytest
from lxml import etree

from pptx.oxml import oxml_parser, parse_xml, register_element_cls
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import BaseOxmlElement

from ..unitutil.mock import function_mock, loose_mock, var_mock


class DescribeOxmlParser(object):
    def it_strips_whitespace_between_elements(self, foo, stripped_xml_bytes):
        xml_bytes = etree.tostring(foo)
        assert xml_bytes == stripped_xml_bytes

    def it_does_not_expand_internal_entities(self):
        # --- "billion laughs" payload: if entity expansion were active the
        # --- expanded text would be 5**3 == 125 chars of "lol"; with
        # --- resolve_entities=False the text node is left empty.
        billion_laughs = (
            b'<?xml version="1.0"?>'
            b"<!DOCTYPE lolz ["
            b'  <!ENTITY lol "lol">'
            b'  <!ENTITY lol2 "&lol;&lol;&lol;&lol;&lol;">'
            b'  <!ENTITY lol3 "&lol2;&lol2;&lol2;&lol2;&lol2;">'
            b"]>"
            b"<lolz>&lol3;</lolz>"
        )
        root = parse_xml(billion_laughs)
        # --- no text content means the entity was *not* expanded
        assert root.text is None or "lol" not in (root.text or "")

    def it_does_not_resolve_external_entities(self, tmp_path):
        # --- write a secret file the parser must not read
        secret = tmp_path / "secret.txt"
        secret.write_text("top-secret")
        xxe = (
            '<?xml version="1.0"?>'
            "<!DOCTYPE foo ["
            f'  <!ENTITY xxe SYSTEM "file://{secret}">'
            "]>"
            "<foo>&xxe;</foo>"
        ).encode("utf-8")
        root = parse_xml(xxe)
        # --- external entity reference is left unresolved; the file contents
        # --- never appear in the element tree.
        assert root.text is None or "top-secret" not in (root.text or "")


class DescribeParseXml(object):
    def it_uses_oxml_configured_parser_to_parse_xml(
        self, mock_xml_bytes, fromstring, mock_oxml_parser
    ):
        element = parse_xml(mock_xml_bytes)
        fromstring.assert_called_once_with(mock_xml_bytes, mock_oxml_parser)
        assert element is fromstring.return_value

    def it_prefers_to_parse_bytes(self, xml_bytes):
        parse_xml(xml_bytes)

    def but_accepts_unicode_providing_there_is_no_encoding_declaration(self):
        non_enc_decl = '<?xml version="1.0" standalone="yes"?>'
        enc_decl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        xml_body = "<foo><bar>føøbår</bar></foo>"
        # unicode body by itself doesn't raise
        parse_xml(xml_body)
        # adding XML decl without encoding attr doesn't raise either
        xml_text = "%s\n%s" % (non_enc_decl, xml_body)
        parse_xml(xml_text)
        # but adding encoding in the declaration raises ValueError
        xml_text = "%s\n%s" % (enc_decl, xml_body)
        with pytest.raises(ValueError):
            parse_xml(xml_text)


class DescribeRegisterCustomElementClass(object):
    def it_determines_cust_elm_class_constructed_for_specified_tag(self, xml_bytes):
        register_element_cls("a:foo", CustElmCls)
        foo = etree.fromstring(xml_bytes, oxml_parser)
        assert type(foo) is CustElmCls
        assert type(foo.find(qn("a:bar"))) is etree._Element


# ===========================================================================
# fixtures
# ===========================================================================


class CustElmCls(BaseOxmlElement):
    pass


@pytest.fixture
def foo(xml_bytes):
    return etree.fromstring(xml_bytes, oxml_parser)


@pytest.fixture
def fromstring(request):
    return function_mock(request, "pptx.oxml.etree.fromstring")


@pytest.fixture
def mock_oxml_parser(request):
    return var_mock(request, "pptx.oxml.oxml_parser")


@pytest.fixture
def mock_xml_bytes(request):
    return loose_mock(request, "xml_bytes")


@pytest.fixture
def stripped_xml_bytes():
    return (
        '<a:foo xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/ma'
        'in"><a:bar>foobar</a:bar></a:foo>'
    ).encode("utf-8")


@pytest.fixture
def xml_bytes(xml_text):
    return xml_text.encode("utf-8")


@pytest.fixture
def xml_text():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<a:foo xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/ma'
        'in">\n'
        "  <a:bar>foobar</a:bar>\n"
        "</a:foo>\n"
    )
