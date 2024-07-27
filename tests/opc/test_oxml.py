"""Unit-test suite for `pptx.opc.oxml` module."""

from __future__ import annotations

from typing import cast

import pytest

from pptx.opc.constants import RELATIONSHIP_TARGET_MODE as RTM
from pptx.opc.oxml import (
    CT_Default,
    CT_Override,
    CT_Relationship,
    CT_Relationships,
    CT_Types,
    nsmap,
    oxml_tostring,
    serialize_part_xml,
)
from pptx.opc.packuri import PackURI
from pptx.oxml import parse_xml
from pptx.oxml.xmlchemy import BaseOxmlElement

from ..unitutil.cxml import element


class DescribeCT_Default:
    """Unit-test suite for `pptx.opc.oxml.CT_Default` objects."""

    def it_provides_read_access_to_xml_values(self):
        default = cast(CT_Default, element("ct:Default{Extension=xml,ContentType=application/xml}"))
        assert default.extension == "xml"
        assert default.contentType == "application/xml"


class DescribeCT_Override:
    """Unit-test suite for `pptx.opc.oxml.CT_Override` objects."""

    def it_provides_read_access_to_xml_values(self):
        override = cast(
            CT_Override, element("ct:Override{PartName=/part/name.xml,ContentType=text/plain}")
        )
        assert override.partName == "/part/name.xml"
        assert override.contentType == "text/plain"


class DescribeCT_Relationship:
    """Unit-test suite for `pptx.opc.oxml.CT_Relationship` objects."""

    def it_provides_read_access_to_xml_values(self):
        rel = cast(
            CT_Relationship,
            element("pr:Relationship{Id=rId9,Type=ReLtYpE,Target=docProps/core.xml}"),
        )
        assert rel.rId == "rId9"
        assert rel.reltype == "ReLtYpE"
        assert rel.target_ref == "docProps/core.xml"
        assert rel.targetMode == RTM.INTERNAL

    def it_constructs_an_internal_relationship_when_no_target_mode_is_provided(self):
        rel = CT_Relationship.new("rId9", "ReLtYpE", "foo/bar.xml")

        assert rel.rId == "rId9"
        assert rel.reltype == "ReLtYpE"
        assert rel.target_ref == "foo/bar.xml"
        assert rel.targetMode == RTM.INTERNAL
        assert rel.xml == (
            f'<Relationship xmlns="{nsmap["pr"]}" Id="rId9" Type="ReLtYpE" Target="foo/bar.xml"/>'
        )

    def and_it_constructs_an_internal_relationship_when_target_mode_INTERNAL_is_specified(self):
        rel = CT_Relationship.new("rId9", "ReLtYpE", "foo/bar.xml", RTM.INTERNAL)

        assert rel.rId == "rId9"
        assert rel.reltype == "ReLtYpE"
        assert rel.target_ref == "foo/bar.xml"
        assert rel.targetMode == RTM.INTERNAL
        assert rel.xml == (
            f'<Relationship xmlns="{nsmap["pr"]}" Id="rId9" Type="ReLtYpE" Target="foo/bar.xml"/>'
        )

    def and_it_constructs_an_external_relationship_when_target_mode_EXTERNAL_is_specified(self):
        rel = CT_Relationship.new("rId9", "ReLtYpE", "http://some/link", RTM.EXTERNAL)

        assert rel.rId == "rId9"
        assert rel.reltype == "ReLtYpE"
        assert rel.target_ref == "http://some/link"
        assert rel.targetMode == RTM.EXTERNAL
        assert rel.xml == (
            f'<Relationship xmlns="{nsmap["pr"]}" Id="rId9" Type="ReLtYpE"'
            f' Target="http://some/link" TargetMode="External"/>'
        )


class DescribeCT_Relationships:
    """Unit-test suite for `pptx.opc.oxml.CT_Relationships` objects."""

    def it_can_construct_a_new_relationships_element(self):
        rels = CT_Relationships.new()
        assert rels.xml == (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
        )

    def it_can_build_rels_element_incrementally(self):
        rels = CT_Relationships.new()

        rels.add_rel("rId1", "http://reltype1", "docProps/core.xml")
        rels.add_rel("rId2", "http://linktype", "http://some/link", True)
        rels.add_rel("rId3", "http://reltype2", "../slides/slide1.xml")

        assert oxml_tostring(rels, encoding="unicode", pretty_print=True) == (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
            '  <Relationship Id="rId1" Type="http://reltype1" Target="docProps/core.xml"/>\n'
            '  <Relationship Id="rId2" Type="http://linktype" Target="http://some/link"'
            ' TargetMode="External"/>\n'
            '  <Relationship Id="rId3" Type="http://reltype2" Target="../slides/slide1.xml"/>\n'
            "</Relationships>\n"
        )

    def it_can_generate_rels_file_xml(self):
        assert CT_Relationships.new().xml_file_bytes == (
            "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n"
            '<Relationships xmlns="http://schemas.openxmlformats.org/package'
            '/2006/relationships"/>'.encode("utf-8")
        )


class DescribeCT_Types:
    """Unit-test suite for `pptx.opc.oxml.CT_Types` objects."""

    def it_provides_access_to_default_child_elements(self, types: CT_Types):
        assert len(types.default_lst) == 2
        for default in types.default_lst:
            assert isinstance(default, CT_Default)

    def it_provides_access_to_override_child_elements(self, types: CT_Types):
        assert len(types.override_lst) == 3
        for override in types.override_lst:
            assert isinstance(override, CT_Override)

    def it_should_have_empty_list_on_no_matching_elements(self):
        types = cast(CT_Types, element("ct:Types"))
        assert types.default_lst == []
        assert types.override_lst == []

    def it_can_construct_a_new_types_element(self):
        types = CT_Types.new()
        assert types.xml == (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>\n'
        )

    def it_can_build_types_element_incrementally(self):
        types = CT_Types.new()
        types.add_default("xml", "application/xml")
        types.add_default("jpeg", "image/jpeg")
        types.add_override(PackURI("/docProps/core.xml"), "app/vnd.type1")
        types.add_override(PackURI("/ppt/presentation.xml"), "app/vnd.type2")
        types.add_override(PackURI("/docProps/thumbnail.jpeg"), "image/jpeg")
        assert types.xml == (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n'
            '  <Default Extension="xml" ContentType="application/xml"/>\n'
            '  <Default Extension="jpeg" ContentType="image/jpeg"/>\n'
            '  <Override PartName="/docProps/core.xml" ContentType="app/vnd.type1"/>\n'
            '  <Override PartName="/ppt/presentation.xml" ContentType="app/vnd.type2"/>\n'
            '  <Override PartName="/docProps/thumbnail.jpeg" ContentType="image/jpeg"/>\n'
            "</Types>\n"
        )

    # -- fixtures ----------------------------------------------------

    @pytest.fixture
    def types(self) -> CT_Types:
        return cast(
            CT_Types,
            element(
                "ct:Types/(ct:Default{Extension=xml,ContentType=application/xml}"
                ",ct:Default{Extension=jpeg,ContentType=image/jpeg}"
                ",ct:Override{PartName=/docProps/core.xml,ContentType=app/vnd.type1}"
                ",ct:Override{PartName=/ppt/presentation.xml,ContentType=app/vnd.type2}"
                ",ct:Override{PartName=/docProps/thunbnail.jpeg,ContentType=image/jpeg})"
            ),
        )


class Describe_serialize_part_xml:
    """Unit-test suite for `pptx.opc.oxml.serialize_part_xml` function."""

    def it_produces_properly_formatted_xml_for_an_opc_part(self):
        """
        Tested aspects:
        ---------------
        * [X] it generates an XML declaration
        * [X] it produces no whitespace between elements
        * [X] it preserves unused namespaces
        * [X] it returns bytes ready to save to file (not unicode)
        """
        part_elm = cast(
            BaseOxmlElement,
            parse_xml(
                '<f:foo xmlns:f="http://foo" xmlns:b="http://bar">\n  <f:bar>fØØ'
                "bÅr</f:bar>\n</f:foo>\n"
            ),
        )
        xml = serialize_part_xml(part_elm)
        # xml contains 134 chars, of which 3 are double-byte; it will have
        # len of 134 if it's unicode and 137 if it's bytes
        assert len(xml) == 137
        assert xml == (
            "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n"
            '<f:foo xmlns:f="http://foo" xmlns:b="http://bar"><f:bar>fØØbÅr</f:bar></f:foo>'
        ).encode("utf-8")
