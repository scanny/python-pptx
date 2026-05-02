# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.oxml.extprops` module."""

from __future__ import annotations

import pytest

from pptx.oxml import parse_xml
from pptx.oxml.extprops import CT_ExtendedProperties


class DescribeCT_ExtendedProperties(object):
    """Unit-test suite for `pptx.oxml.extprops.CT_ExtendedProperties` objects."""

    def it_can_construct_a_new_default_Properties_element(self):
        props = CT_ExtendedProperties.new_extendedProperties()

        # -- root is in the ep namespace with `Slides` = 0 --
        assert isinstance(props, CT_ExtendedProperties)
        assert props.tag == (
            "{http://schemas.openxmlformats.org/officeDocument/2006/"
            "extended-properties}Properties"
        )
        assert props.slide_count == 0

    def it_knows_its_slide_count_when_Slides_is_present(self):
        props = self._props_with_slides_xml("7")

        assert props.slide_count == 7

    def it_returns_zero_when_the_Slides_element_is_missing(self):
        props = parse_xml(
            '<Properties xmlns="http://schemas.openxmlformats.org/'
            'officeDocument/2006/extended-properties"/>'
        )

        assert props.slide_count == 0

    @pytest.mark.parametrize("bogus_text", ["", "not-a-number", "-3"])
    def it_returns_zero_when_the_Slides_text_is_invalid(self, bogus_text: str):
        props = self._props_with_slides_xml(bogus_text)

        assert props.slide_count == 0

    def it_can_set_the_slide_count_when_Slides_is_absent(self):
        props = CT_ExtendedProperties.new_extendedProperties()
        # -- remove the auto-added Slides child to exercise the "create" branch --
        slides = props.find(
            "{http://schemas.openxmlformats.org/officeDocument/2006/"
            "extended-properties}Slides"
        )
        props.remove(slides)

        props.slide_count = 12

        assert props.slide_count == 12

    def it_can_update_the_slide_count_when_Slides_is_present(self):
        props = self._props_with_slides_xml("2")

        props.slide_count = 9

        assert props.slide_count == 9

    @pytest.mark.parametrize("bad_value", [-1, "3", 3.0, None])
    def it_raises_on_a_bad_slide_count_value(self, bad_value: object):
        props = CT_ExtendedProperties.new_extendedProperties()

        with pytest.raises(ValueError, match="slide_count"):
            props.slide_count = bad_value  # type: ignore[assignment]

    # -- helpers ----------------------------------------------------------

    @staticmethod
    def _props_with_slides_xml(slides_text: str) -> CT_ExtendedProperties:
        xml = (
            '<Properties xmlns="http://schemas.openxmlformats.org/'
            'officeDocument/2006/extended-properties">'
            "<Slides>%s</Slides>"
            "</Properties>"
        ) % slides_text
        return parse_xml(xml)
