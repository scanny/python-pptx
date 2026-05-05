# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.oxml.custprops` module."""

from __future__ import annotations

import datetime as dt

import pytest
from lxml import etree

from pptx.oxml import parse_xml
from pptx.oxml.custprops import (
    FMTID_CUSTOM,
    CT_CustomProperties,
    _parse_iso_datetime,
    _python_value_of,
    _vt_tag_and_text_for,
)
from pptx.oxml.ns import qn


class DescribeCT_CustomProperties(object):
    """Unit-test suite for `pptx.oxml.custprops.CT_CustomProperties` objects."""

    def it_can_construct_a_new_empty_Properties_element(self):
        props = CT_CustomProperties.new_customProperties()

        assert isinstance(props, CT_CustomProperties)
        assert props.tag == (
            "{http://schemas.openxmlformats.org/officeDocument/2006/"
            "custom-properties}Properties"
        )
        assert len(props) == 0
        assert props.names() == []

    def it_assigns_pid_2_to_the_first_added_property(self):
        props = CT_CustomProperties.new_customProperties()

        props.set_value("Alpha", "one")

        prop = props.find(qn("cst:property"))
        assert prop is not None
        assert prop.get("pid") == "2"
        assert prop.get("fmtid") == FMTID_CUSTOM
        assert prop.get("name") == "Alpha"

    def it_increments_pid_for_subsequent_properties(self):
        props = CT_CustomProperties.new_customProperties()

        props.set_value("Alpha", "one")
        props.set_value("Beta", "two")
        props.set_value("Gamma", "three")

        pids = [prop.get("pid") for prop in props.findall(qn("cst:property"))]
        assert pids == ["2", "3", "4"]

    def it_serializes_string_as_vt_lpwstr(self):
        props = CT_CustomProperties.new_customProperties()
        props.set_value("MyStr", "hello")
        xml = etree.tostring(props).decode()
        assert "<vt:lpwstr>hello</vt:lpwstr>" in xml
        assert props.get_value("MyStr") == "hello"

    def it_serializes_int_as_vt_i4(self):
        props = CT_CustomProperties.new_customProperties()
        props.set_value("MyInt", 42)
        xml = etree.tostring(props).decode()
        assert "<vt:i4>42</vt:i4>" in xml
        assert props.get_value("MyInt") == 42

    def it_serializes_float_as_vt_r8(self):
        props = CT_CustomProperties.new_customProperties()
        props.set_value("MyFloat", 2.5)
        xml = etree.tostring(props).decode()
        assert "<vt:r8>2.5</vt:r8>" in xml
        assert props.get_value("MyFloat") == 2.5

    @pytest.mark.parametrize(
        "value, expected_text",
        [(True, "true"), (False, "false")],
    )
    def it_serializes_bool_as_vt_bool(self, value: bool, expected_text: str):
        props = CT_CustomProperties.new_customProperties()
        props.set_value("MyBool", value)
        xml = etree.tostring(props).decode()
        assert "<vt:bool>%s</vt:bool>" % expected_text in xml
        assert props.get_value("MyBool") is value

    def it_serializes_bool_separately_from_int(self):
        # -- bool is an int subclass; the setter must detect bool first to avoid
        # -- emitting `<vt:i4>1</vt:i4>` for True.
        props = CT_CustomProperties.new_customProperties()
        props.set_value("flag", True)
        prop = props.find(qn("cst:property"))
        assert prop is not None
        child = prop[0]
        assert child.tag == qn("vt:bool")

    def it_serializes_naive_datetime_as_vt_filetime(self):
        props = CT_CustomProperties.new_customProperties()
        props.set_value("created", dt.datetime(2024, 1, 2, 3, 4, 5))
        xml = etree.tostring(props).decode()
        assert "<vt:filetime>2024-01-02T03:04:05Z</vt:filetime>" in xml
        assert props.get_value("created") == dt.datetime(2024, 1, 2, 3, 4, 5)

    def it_converts_aware_datetime_to_utc_on_serialize(self):
        tz = dt.timezone(dt.timedelta(hours=5))
        props = CT_CustomProperties.new_customProperties()
        props.set_value("d", dt.datetime(2024, 1, 2, 3, 4, 5, tzinfo=tz))
        xml = etree.tostring(props).decode()
        # -- 03:04:05 +05:00 -> 22:04:05 UTC on the previous day --
        assert "<vt:filetime>2024-01-01T22:04:05Z</vt:filetime>" in xml

    def it_serializes_plain_date_as_vt_date(self):
        """`datetime.date` (not `datetime`) must route to `vt:date`, `YYYY-MM-DD`."""
        props = CT_CustomProperties.new_customProperties()
        props.set_value("review", dt.date(2024, 1, 2))
        xml = etree.tostring(props).decode()
        assert "<vt:date>2024-01-02</vt:date>" in xml
        assert "<vt:filetime>" not in xml

    def it_round_trips_a_vt_date_as_a_plain_date(self):
        """Round-tripping a `date` must yield a `date`, not a `datetime` —
        otherwise the on-disk distinction between `vt:date` and `vt:filetime`
        is lost on re-save."""
        props = CT_CustomProperties.new_customProperties()
        original = dt.date(2024, 1, 2)
        props.set_value("review", original)

        retrieved = props.get_value("review")

        assert retrieved == original
        assert isinstance(retrieved, dt.date)
        assert not isinstance(retrieved, dt.datetime)

    def it_dispatches_datetime_to_filetime_not_date(self):
        """`datetime` is a subclass of `date`; the encoder must keep them distinct."""
        props = CT_CustomProperties.new_customProperties()
        props.set_value("d", dt.datetime(2024, 1, 2, 3, 4, 5))
        xml = etree.tostring(props).decode()
        assert "<vt:filetime>" in xml
        assert "<vt:date>" not in xml

    def it_updates_existing_property_in_place_preserving_pid(self):
        props = CT_CustomProperties.new_customProperties()
        props.set_value("x", "first")
        props.set_value("y", "other")
        # -- change the value of `x` --
        props.set_value("x", "second")

        xs = [p for p in props.findall(qn("cst:property")) if p.get("name") == "x"]
        assert len(xs) == 1
        assert xs[0].get("pid") == "2"
        assert xs[0].find(qn("vt:lpwstr")).text == "second"

    def it_changes_the_typed_vt_child_on_type_change(self):
        props = CT_CustomProperties.new_customProperties()
        props.set_value("x", "string")
        props.set_value("x", 42)  # -- now an int --
        xml = etree.tostring(props).decode()
        assert "<vt:lpwstr>" not in xml
        assert "<vt:i4>42</vt:i4>" in xml
        assert props.get_value("x") == 42

    def it_can_delete_a_property(self):
        props = CT_CustomProperties.new_customProperties()
        props.set_value("x", "one")
        props.set_value("y", "two")

        assert props.delete_value("x") is True
        assert props.has_name("x") is False
        assert props.has_name("y") is True

    def it_returns_False_when_deleting_missing_property(self):
        props = CT_CustomProperties.new_customProperties()
        assert props.delete_value("nope") is False

    def it_returns_None_for_missing_get_value(self):
        props = CT_CustomProperties.new_customProperties()
        assert props.get_value("nope") is None

    def it_iterates_names_in_document_order(self):
        props = CT_CustomProperties.new_customProperties()
        props.set_value("zebra", 1)
        props.set_value("apple", 2)
        props.set_value("monkey", 3)
        assert list(props) == ["zebra", "apple", "monkey"]

    def it_appends_pid_above_max_even_when_holes_exist(self):
        # -- construct with holes: pids 2, 5 present; next should be 6, not 3 --
        xml = (
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/'
            '2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/'
            'officeDocument/2006/docPropsVTypes">'
            '<property fmtid="%s" pid="2" name="a"><vt:lpwstr>x</vt:lpwstr></property>'
            '<property fmtid="%s" pid="5" name="b"><vt:lpwstr>y</vt:lpwstr></property>'
            "</Properties>"
        ) % (FMTID_CUSTOM, FMTID_CUSTOM)
        props = parse_xml(xml)
        props.set_value("c", "new")
        new_prop = [p for p in props.findall(qn("cst:property")) if p.get("name") == "c"][0]
        assert new_prop.get("pid") == "6"

    @pytest.mark.parametrize("bad_name", ["", None, 123, True])
    def it_raises_on_a_bad_name(self, bad_name: object):
        props = CT_CustomProperties.new_customProperties()
        with pytest.raises(ValueError, match="name"):
            props.set_value(bad_name, "x")  # type: ignore[arg-type]

    @pytest.mark.parametrize("bad_value", [[1, 2], {"k": "v"}, object(), None])
    def it_raises_TypeError_for_unsupported_types(self, bad_value: object):
        props = CT_CustomProperties.new_customProperties()
        with pytest.raises(TypeError):
            props.set_value("x", bad_value)  # type: ignore[arg-type]

    def it_raises_OverflowError_for_int_beyond_int32_range(self):
        props = CT_CustomProperties.new_customProperties()
        with pytest.raises(OverflowError):
            props.set_value("big", 2**31)

    def it_accepts_int_at_the_int32_boundaries(self):
        props = CT_CustomProperties.new_customProperties()
        props.set_value("max", 2**31 - 1)
        props.set_value("min", -(2**31))
        assert props.get_value("max") == 2**31 - 1
        assert props.get_value("min") == -(2**31)


class Describe_vt_tag_and_text_for(object):
    @pytest.mark.parametrize(
        "value, expected",
        [
            ("hi", ("vt:lpwstr", "hi")),
            (7, ("vt:i4", "7")),
            (True, ("vt:bool", "true")),
            (False, ("vt:bool", "false")),
            (2.5, ("vt:r8", "2.5")),
        ],
    )
    def it_returns_the_right_tag_and_text(
        self, value: object, expected: tuple[str, str]
    ):
        assert _vt_tag_and_text_for(value) == expected  # type: ignore[arg-type]


class Describe_python_value_of(object):
    @pytest.mark.parametrize(
        "xml, expected",
        [
            ('<property><vt:lpwstr xmlns:vt="{VT}">text</vt:lpwstr></property>', "text"),
            ('<property><vt:lpstr xmlns:vt="{VT}">lp</vt:lpstr></property>', "lp"),
            ('<property><vt:bstr xmlns:vt="{VT}">bstr</vt:bstr></property>', "bstr"),
            ('<property><vt:i4 xmlns:vt="{VT}">42</vt:i4></property>', 42),
            ('<property><vt:int xmlns:vt="{VT}">5</vt:int></property>', 5),
            ('<property><vt:ui4 xmlns:vt="{VT}">9</vt:ui4></property>', 9),
            ('<property><vt:r8 xmlns:vt="{VT}">2.5</vt:r8></property>', 2.5),
            ('<property><vt:r4 xmlns:vt="{VT}">1.5</vt:r4></property>', 1.5),
            ('<property><vt:bool xmlns:vt="{VT}">true</vt:bool></property>', True),
            ('<property><vt:bool xmlns:vt="{VT}">1</vt:bool></property>', True),
            ('<property><vt:bool xmlns:vt="{VT}">false</vt:bool></property>', False),
            ('<property><vt:bool xmlns:vt="{VT}">0</vt:bool></property>', False),
        ],
    )
    def it_decodes_the_first_vt_child(self, xml: str, expected: object):
        VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
        prop = etree.fromstring(xml.replace("{VT}", VT))
        assert _python_value_of(prop) == expected

    def it_returns_None_when_no_vt_child(self):
        prop = etree.fromstring("<property/>")
        assert _python_value_of(prop) is None


class Describe_parse_iso_datetime(object):
    @pytest.mark.parametrize(
        "text, expected",
        [
            ("2024-01-02T03:04:05Z", dt.datetime(2024, 1, 2, 3, 4, 5)),
            (
                "2024-01-02T03:04:05+05:00",
                dt.datetime(2024, 1, 1, 22, 4, 5),
            ),
            ("2024-01-02T03:04:05", dt.datetime(2024, 1, 2, 3, 4, 5)),
        ],
    )
    def it_parses_iso8601(self, text: str, expected: dt.datetime):
        assert _parse_iso_datetime(text) == expected

    @pytest.mark.parametrize("bad", ["", "not-a-date", "2024-13-40"])
    def it_returns_None_for_bad_input(self, bad: str):
        assert _parse_iso_datetime(bad) is None
