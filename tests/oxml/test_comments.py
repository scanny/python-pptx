# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.oxml.comments` module."""

from __future__ import annotations

import datetime as dt

import pytest

from pptx.oxml import parse_xml
from pptx.oxml.comments import (
    CT_Comment,
    CT_CommentAuthor,
    CT_CommentAuthorList,
    CT_CommentList,
    _format_w3cdtf,
    _parse_w3cdtf,
)
from pptx.oxml.ns import qn


class DescribeCT_CommentAuthorList:
    def it_constructs_a_new_empty_cmAuthorLst(self):
        cmAuthorLst = CT_CommentAuthorList.new()

        assert isinstance(cmAuthorLst, CT_CommentAuthorList)
        assert cmAuthorLst.tag == qn("p:cmAuthorLst")
        assert len(cmAuthorLst.cmAuthor_lst) == 0

    def it_appends_new_authors_with_ascending_ids(self):
        cmAuthorLst = CT_CommentAuthorList.new()

        a0 = cmAuthorLst.add_author("Alice", "A.")
        a1 = cmAuthorLst.add_author("Bob", "B.")

        assert a0.id == 0
        assert a0.clrIdx == 0
        assert a0.lastIdx == 0
        assert a0.name == "Alice"
        assert a0.initials == "A."
        assert a1.id == 1
        assert a1.clrIdx == 1

    def it_can_look_up_an_author_by_id(self):
        cmAuthorLst = CT_CommentAuthorList.new()
        cmAuthorLst.add_author("Alice", "A.")
        cmAuthorLst.add_author("Bob", "B.")

        found = cmAuthorLst.get_by_id(1)

        assert found is not None
        assert found.name == "Bob"

    def but_get_by_id_returns_None_for_unknown_id(self):
        cmAuthorLst = CT_CommentAuthorList.new()

        assert cmAuthorLst.get_by_id(42) is None


class DescribeCT_CommentList:
    def it_constructs_a_new_empty_cmLst(self):
        cmLst = CT_CommentList.new()

        assert isinstance(cmLst, CT_CommentList)
        assert cmLst.tag == qn("p:cmLst")
        assert len(cmLst.cm_lst) == 0

    def it_appends_new_comments_with_correct_child_order(self):
        cmLst = CT_CommentList.new()

        cm = cmLst.add_comment(
            author_id=0,
            text="Hi",
            position=(10, 20),
            datetime_value=dt.datetime(2025, 1, 2, 3, 4, 5),
        )

        assert cm.authorId == 0
        assert cm.idx == 1
        assert cm.text_value == "Hi"
        assert cm.position == (10, 20)
        assert cm.datetime_value == dt.datetime(2025, 1, 2, 3, 4, 5)
        # -- `p:pos` must precede `p:text` per schema --
        children = list(cm)
        assert children[0].tag == qn("p:pos")
        assert children[1].tag == qn("p:text")

    def it_assigns_per_author_ascending_idx_values(self):
        cmLst = CT_CommentList.new()

        a = cmLst.add_comment(author_id=0, text="a1")
        b = cmLst.add_comment(author_id=0, text="a2")
        c = cmLst.add_comment(author_id=1, text="b1")

        assert a.idx == 1
        assert b.idx == 2
        assert c.idx == 1


class DescribeCT_Comment:
    def it_reads_and_writes_text(self):
        cmLst = CT_CommentList.new()
        cm = cmLst.add_comment(author_id=0, text="hello")

        cm.text_value = "world"

        assert cm.text_value == "world"

    def it_reads_and_writes_datetime(self):
        cmLst = CT_CommentList.new()
        cm = cmLst.add_comment(author_id=0, text="hi")

        cm.datetime_value = dt.datetime(2024, 12, 31, 23, 59, 59)

        assert cm.datetime_value == dt.datetime(2024, 12, 31, 23, 59, 59)

    def it_returns_None_for_datetime_when_dt_attribute_is_missing(self):
        cmLst = CT_CommentList.new()
        cm = cmLst.add_comment(author_id=0, text="hi")

        assert cm.datetime_value is None

    def it_clears_the_dt_attribute_when_assigned_None(self):
        cmLst = CT_CommentList.new()
        cm = cmLst.add_comment(author_id=0, text="hi", datetime_value=dt.datetime(2024, 1, 1))

        cm.datetime_value = None

        assert cm.datetime_value is None
        assert cm.get("dt") is None

    def it_reads_and_writes_position(self):
        cmLst = CT_CommentList.new()
        cm = cmLst.add_comment(author_id=0, text="hi", position=(5, 6))

        assert cm.position == (5, 6)

        cm.position = (100, 200)

        assert cm.position == (100, 200)

    def it_reads_empty_text_when_text_element_has_no_content(self):
        xml = (
            '<p:cm xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            ' authorId="0" idx="1"><p:pos x="0" y="0"/><p:text/></p:cm>'
        )
        cm = parse_xml(xml)
        assert isinstance(cm, CT_Comment)
        assert cm.text_value == ""


class DescribeCT_CommentAuthor:
    def it_parses_required_attribute_types(self):
        xml = (
            '<p:cmAuthor xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            ' id="3" name="Carol" initials="C." lastIdx="7" clrIdx="3"/>'
        )
        author = parse_xml(xml)

        assert isinstance(author, CT_CommentAuthor)
        assert author.id == 3
        assert author.name == "Carol"
        assert author.initials == "C."
        assert author.lastIdx == 7
        assert author.clrIdx == 3


class Describe_w3cdtf_helpers:
    @pytest.mark.parametrize(
        ("input_str", "expected"),
        [
            ("2025-01-02T03:04:05", dt.datetime(2025, 1, 2, 3, 4, 5)),
            ("2025-01-02", dt.datetime(2025, 1, 2)),
            ("2025-01", dt.datetime(2025, 1, 1)),
            ("2025", dt.datetime(2025, 1, 1)),
        ],
    )
    def it_parses_valid_strings(self, input_str: str, expected: dt.datetime):
        assert _parse_w3cdtf(input_str) == expected

    def it_applies_positive_timezone_offsets(self):
        # -- '+' means subtract from UTC to get local, so parse expects subtraction --
        result = _parse_w3cdtf("2025-01-02T03:04:05+02:00")
        assert result == dt.datetime(2025, 1, 2, 1, 4, 5)

    def it_applies_negative_timezone_offsets(self):
        result = _parse_w3cdtf("2025-01-02T03:04:05-05:00")
        assert result == dt.datetime(2025, 1, 2, 8, 4, 5)

    def it_returns_None_for_unparseable_input(self):
        assert _parse_w3cdtf("not-a-date") is None

    def it_formats_naive_datetimes_without_offset(self):
        assert _format_w3cdtf(dt.datetime(2025, 1, 2, 3, 4, 5)) == "2025-01-02T03:04:05"

    def it_normalizes_aware_datetimes_to_UTC(self):
        tz = dt.timezone(dt.timedelta(hours=5))
        aware = dt.datetime(2025, 1, 2, 10, 0, 0, tzinfo=tz)
        assert _format_w3cdtf(aware) == "2025-01-02T05:00:00"
