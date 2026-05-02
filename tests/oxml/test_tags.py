# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.oxml.tags` module."""

from __future__ import annotations

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.oxml.tags import (
    CT_CustomerDataList,
    CT_StringTag,
    CT_TagList,
    CT_TagsData,
)


class DescribeCT_TagList:
    def it_constructs_a_new_empty_tagLst(self):
        tagLst = CT_TagList.new()

        assert isinstance(tagLst, CT_TagList)
        assert tagLst.tag == qn("p:tagLst")
        assert len(tagLst.tag_lst) == 0

    def it_appends_a_new_tag_via_set_tag(self):
        tagLst = CT_TagList.new()

        tag = tagLst.set_tag("PRIORITY", "HIGH")

        assert isinstance(tag, CT_StringTag)
        assert tag.name == "PRIORITY"
        assert tag.val == "HIGH"
        assert len(tagLst.tag_lst) == 1

    def it_overwrites_an_existing_tag_value_via_set_tag(self):
        tagLst = CT_TagList.new()
        tagLst.set_tag("OWNER", "alice")

        tagLst.set_tag("OWNER", "bob")

        assert len(tagLst.tag_lst) == 1
        assert tagLst.get_by_name("OWNER") is not None
        assert tagLst.get_by_name("OWNER").val == "bob"

    def it_looks_up_a_tag_by_name(self):
        tagLst = CT_TagList.new()
        tagLst.set_tag("A", "1")
        tagLst.set_tag("B", "2")

        found = tagLst.get_by_name("B")

        assert found is not None
        assert found.val == "2"

    def but_get_by_name_returns_None_when_absent(self):
        tagLst = CT_TagList.new()
        assert tagLst.get_by_name("NOPE") is None

    def it_removes_a_tag_by_name(self):
        tagLst = CT_TagList.new()
        tagLst.set_tag("A", "1")
        tagLst.set_tag("B", "2")

        removed = tagLst.remove_tag("A")

        assert removed is True
        assert len(tagLst.tag_lst) == 1
        assert tagLst.get_by_name("A") is None

    def but_remove_tag_returns_False_when_absent(self):
        tagLst = CT_TagList.new()

        assert tagLst.remove_tag("NOPE") is False


class DescribeCT_StringTag:
    def it_parses_the_required_attributes(self):
        xml = (
            '<p:tag xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            ' name="color" val="red"/>'
        )
        tag = parse_xml(xml)

        assert isinstance(tag, CT_StringTag)
        assert tag.name == "color"
        assert tag.val == "red"

    def it_allows_writing_new_values(self):
        tagLst = CT_TagList.new()
        tag = tagLst.set_tag("k", "v1")

        tag.val = "v2"
        tag.name = "K"

        assert tag.name == "K"
        assert tag.val == "v2"


class DescribeCT_TagsData:
    def it_reads_its_rId_attribute(self):
        xml = (
            '<p:tags xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
            ' r:id="rId7"/>'
        )
        tags = parse_xml(xml)

        assert isinstance(tags, CT_TagsData)
        assert tags.rId == "rId7"


class DescribeCT_CustomerDataList:
    def it_exposes_an_existing_tags_child(self):
        xml = (
            f"<p:custDataLst {nsdecls('p', 'r')}>"
            '<p:tags r:id="rId3"/>'
            "</p:custDataLst>"
        )
        custDataLst = parse_xml(xml)

        assert isinstance(custDataLst, CT_CustomerDataList)
        assert custDataLst.tags is not None
        assert custDataLst.tags.rId == "rId3"

    def it_reports_None_when_no_tags_child_is_present(self):
        xml = f"<p:custDataLst {nsdecls('p')}/>"
        custDataLst = parse_xml(xml)

        assert custDataLst.tags is None

    def it_appends_tags_after_existing_custData_children(self):
        xml = (
            f"<p:custDataLst {nsdecls('p', 'r')}>"
            '<p:custData r:id="rX"/>'
            "</p:custDataLst>"
        )
        custDataLst = parse_xml(xml)

        custDataLst.get_or_add_tags()

        children = list(custDataLst)
        assert children[0].tag == qn("p:custData")
        assert children[1].tag == qn("p:tags")
