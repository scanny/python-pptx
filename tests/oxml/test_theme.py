"""Unit-test suite for `pptx.oxml.theme` module."""

from __future__ import annotations

import pytest

from pptx.dml.color import RGBColor
from pptx.oxml import parse_xml
from pptx.oxml.theme import CT_OfficeStyleSheet

from ..unitutil.file import snippet_text


class DescribeCT_OfficeStyleSheet(object):
    def it_can_create_a_default_theme_element(self, new_fixture):
        expected_xml = new_fixture
        theme = CT_OfficeStyleSheet.new_default()
        assert theme.xml == expected_xml

    # -- hyperlink color-scheme accessors (issue #940) -----------------

    def it_returns_the_theme_hlink_color(self):
        theme = CT_OfficeStyleSheet.new_default()
        assert theme.hlink_color == RGBColor.from_string("0000FF")

    def it_returns_the_theme_folHlink_color(self):
        theme = CT_OfficeStyleSheet.new_default()
        assert theme.folHlink_color == RGBColor.from_string("800080")

    def it_returns_None_for_hlink_color_when_srgbClr_absent(self):
        xml = _wrap_clr_scheme("<a:hlink><a:sysClr val='windowText' lastClr='000000'/></a:hlink>")
        theme = parse_xml(xml)
        assert theme.hlink_color is None

    def it_returns_None_for_folHlink_color_when_srgbClr_absent(self):
        xml = _wrap_clr_scheme(
            "<a:folHlink><a:sysClr val='windowText' lastClr='000000'/></a:folHlink>"
        )
        theme = parse_xml(xml)
        assert theme.folHlink_color is None

    def it_can_set_the_theme_hlink_color(self):
        theme = CT_OfficeStyleSheet.new_default()
        theme.hlink_color = RGBColor(0x12, 0x34, 0x56)
        assert theme.hlink_color == RGBColor.from_string("123456")

    def it_can_set_the_theme_folHlink_color(self):
        theme = CT_OfficeStyleSheet.new_default()
        theme.folHlink_color = RGBColor(0xAB, 0xCD, 0xEF)
        assert theme.folHlink_color == RGBColor.from_string("ABCDEF")

    def it_replaces_a_sysClr_when_setting_hlink_color(self):
        xml = _wrap_clr_scheme("<a:hlink><a:sysClr val='windowText' lastClr='000000'/></a:hlink>")
        theme = parse_xml(xml)
        theme.hlink_color = RGBColor(0x01, 0x02, 0x03)
        assert theme.hlink_color == RGBColor.from_string("010203")

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def new_fixture(self):
        expected_xml = snippet_text("default-theme")
        return expected_xml


def _wrap_clr_scheme(inner: str) -> str:
    """Return a minimal `a:theme` XML string wrapping `inner` inside `a:clrScheme`."""
    return (
        '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="t">'
        '<a:themeElements><a:clrScheme name="x">' + inner + "</a:clrScheme></a:themeElements>"
        "</a:theme>"
    )
