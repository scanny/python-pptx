# encoding: utf-8

"""
Test suite for opc.oxml module
"""

from __future__ import absolute_import, print_function, unicode_literals

import pytest

from pptx.opc.constants import RELATIONSHIP_TARGET_MODE as RTM
from pptx.opc.oxml import (
    CT_Default,
    CT_Override,
    CT_Relationship,
    CT_Relationships,
    CT_Types,
    oxml_tostring,
    serialize_part_xml,
)
from pptx.oxml import parse_xml

from .unitdata.rels import (
    a_Default,
    an_Override,
    a_Relationship,
    a_Relationships,
    a_Types,
)


class DescribeCT_Default(object):
    def it_provides_read_access_to_xml_values(self):
        default = a_Default().element
        assert default.extension == "xml"
        assert default.contentType == "application/xml"


class DescribeCT_Override(object):
    def it_provides_read_access_to_xml_values(self):
        override = an_Override().element
        assert override.partName == "/part/name.xml"
        assert override.contentType == "app/vnd.type"


class DescribeCT_Relationship(object):
    def it_provides_read_access_to_xml_values(self):
        rel = a_Relationship().element
        assert rel.rId == "rId9"
        assert rel.reltype == "ReLtYpE"
        assert rel.target_ref == "docProps/core.xml"
        assert rel.targetMode == RTM.INTERNAL

    def it_can_construct_from_attribute_values(self):
        cases = (
            ("rId9", "ReLtYpE", "foo/bar.xml", None),
            ("rId9", "ReLtYpE", "bar/foo.xml", RTM.INTERNAL),
            ("rId9", "ReLtYpE", "http://some/link", RTM.EXTERNAL),
        )
        for rId, reltype, target, target_mode in cases:
            if target_mode is None:
                rel = CT_Relationship.new(rId, reltype, target)
            else:
                rel = CT_Relationship.new(rId, reltype, target, target_mode)
            builder = a_Relationship().with_target(target)
            if target_mode == RTM.EXTERNAL:
                builder = builder.with_target_mode(RTM.EXTERNAL)
            expected_rel_xml = builder.xml
            assert rel.xml == expected_rel_xml


class DescribeCT_Relationships(object):
    def it_can_construct_a_new_relationships_element(self):
        rels = CT_Relationships.new()
        expected_xml = (
            "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n"
            '<Relationships xmlns="http://schemas.openxmlformats.org/package'
            '/2006/relationships"/>'
        )
        assert rels.xml.decode("utf-8") == expected_xml

    def it_can_build_rels_element_incrementally(self):
        # setup ------------------------
        rels = CT_Relationships.new()
        # exercise ---------------------
        rels.add_rel("rId1", "http://reltype1", "docProps/core.xml")
        rels.add_rel("rId2", "http://linktype", "http://some/link", True)
        rels.add_rel("rId3", "http://reltype2", "../slides/slide1.xml")
        # verify -----------------------
        expected_rels_xml = a_Relationships().xml
        actual_xml = oxml_tostring(rels, encoding="unicode", pretty_print=True)
        assert actual_xml == expected_rels_xml

    def it_can_generate_rels_file_xml(self):
        expected_xml = (
            "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n"
            '<Relationships xmlns="http://schemas.openxmlformats.org/package'
            '/2006/relationships"/>'.encode("utf-8")
        )
        assert CT_Relationships.new().xml == expected_xml


class DescribeCT_Types(object):
    def it_provides_access_to_default_child_elements(self):
        types = a_Types().element
        assert len(types.default_lst) == 2
        for default in types.default_lst:
            assert isinstance(default, CT_Default)

    def it_provides_access_to_override_child_elements(self):
        types = a_Types().element
        assert len(types.override_lst) == 3
        for override in types.override_lst:
            assert isinstance(override, CT_Override)

    def it_should_have_empty_list_on_no_matching_elements(self):
        types = a_Types().empty().element
        assert types.default_lst == []
        assert types.override_lst == []

    def it_can_construct_a_new_types_element(self):
        types = CT_Types.new()
        expected_xml = a_Types().empty().xml
        assert types.xml == expected_xml

    def it_can_build_types_element_incrementally(self):
        types = CT_Types.new()
        types.add_default("xml", "application/xml")
        types.add_default("jpeg", "image/jpeg")
        types.add_override("/docProps/core.xml", "app/vnd.type1")
        types.add_override("/ppt/presentation.xml", "app/vnd.type2")
        types.add_override("/docProps/thumbnail.jpeg", "image/jpeg")
        expected_types_xml = a_Types().xml
        assert types.xml == expected_types_xml


class DescribeSerializePartXml(object):
    def it_produces_properly_formatted_xml_for_an_opc_part(
        self, part_elm, expected_part_xml
    ):
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
            "bÅr</f:bar>\n</f:foo>\n"
        )

    @pytest.fixture
    def expected_part_xml(self):
        unicode_xml = (
            "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n"
            '<f:foo xmlns:f="http://foo" xmlns:b="http://bar"><f:bar>fØØbÅr<'
            "/f:bar></f:foo>"
        )
        xml_bytes = unicode_xml.encode("utf-8")
        return xml_bytes
