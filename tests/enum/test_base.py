"""Unit-test suite for `pptx.enum.base`."""

from __future__ import annotations

import pytest

from pptx.enum.action import PP_ACTION, PP_ACTION_TYPE
from pptx.enum.dml import MSO_LINE_DASH_STYLE


class DescribeBaseEnum:
    """Unit-test suite for `pptx.enum.base.BaseEnum`."""

    def it_produces_members_each_equivalent_to_an_integer_value(self):
        assert PP_ACTION_TYPE.END_SHOW == 6
        assert PP_ACTION_TYPE.NONE == 0

    def but_member_reprs_are_a_str_indicating_the_enum_and_member_name(self):
        assert repr(PP_ACTION_TYPE.END_SHOW) == "<PP_ACTION_TYPE.END_SHOW: 6>"
        assert repr(PP_ACTION_TYPE.RUN_MACRO) == "<PP_ACTION_TYPE.RUN_MACRO: 8>"

    def and_member_str_values_are_a_str_indicating_the_member_name(self):
        assert str(PP_ACTION_TYPE.FIRST_SLIDE) == "FIRST_SLIDE (3)"
        assert str(PP_ACTION_TYPE.HYPERLINK) == "HYPERLINK (7)"

    def it_provides_a_docstring_for_each_member(self):
        assert PP_ACTION_TYPE.LAST_SLIDE.__doc__ == "Moves to the last slide."
        assert PP_ACTION_TYPE.LAST_SLIDE_VIEWED.__doc__ == "Moves to the last slide viewed."

    def it_can_look_up_a_member_by_its_value(self):
        assert PP_ACTION_TYPE(10) == PP_ACTION_TYPE.NAMED_SLIDE_SHOW
        assert PP_ACTION_TYPE(101) == PP_ACTION_TYPE.NAMED_SLIDE

    def but_it_raises_when_no_member_has_that_value(self):
        with pytest.raises(ValueError, match="42 is not a valid PP_ACTION_TYPE"):
            PP_ACTION_TYPE(42)

    def it_knows_its_name(self):
        assert PP_ACTION_TYPE.NEXT_SLIDE.name == "NEXT_SLIDE"
        assert PP_ACTION_TYPE.NONE.name == "NONE"

    def it_can_be_referred_to_by_a_convenience_alias_if_defined(self):
        assert PP_ACTION_TYPE.OPEN_FILE is PP_ACTION.OPEN_FILE


class DescribeBaseXmlEnum:
    """Unit-test suite for `pptx.enum.base.BaseXmlEnum`."""

    def it_can_look_up_a_member_by_its_corresponding_XML_attribute_value(self):
        assert MSO_LINE_DASH_STYLE.from_xml("dash") == MSO_LINE_DASH_STYLE.DASH
        assert MSO_LINE_DASH_STYLE.from_xml("dashDot") == MSO_LINE_DASH_STYLE.DASH_DOT

    def but_it_raises_on_an_attribute_value_that_is_not_regitstered(self):
        with pytest.raises(ValueError, match="MSO_LINE_DASH_STYLE has no XML mapping for 'wavy'"):
            MSO_LINE_DASH_STYLE.from_xml("wavy")

    def and_the_empty_string_never_maps_to_a_member(self):
        with pytest.raises(ValueError, match="MSO_LINE_DASH_STYLE has no XML mapping for ''"):
            MSO_LINE_DASH_STYLE.from_xml("")

    def it_knows_the_XML_attribute_value_for_each_member_that_has_one(self):
        assert MSO_LINE_DASH_STYLE.to_xml(MSO_LINE_DASH_STYLE.SOLID) == "solid"

    def and_it_looks_up_the_member_by_int_value_before_mapping_when_provided_that_way(self):
        assert MSO_LINE_DASH_STYLE.to_xml(3) == "sysDot"

    def but_it_raises_when_no_member_has_the_provided_int_value(self):
        with pytest.raises(ValueError, match="42 is not a valid MSO_LINE_DASH_STYLE"):
            MSO_LINE_DASH_STYLE.to_xml(42)

    def and_it_raises_when_the_member_has_no_XML_value(self):
        with pytest.raises(ValueError, match="MSO_LINE_DASH_STYLE.DASH_STYLE_MIXED has no XML r"):
            MSO_LINE_DASH_STYLE.to_xml(-2)
