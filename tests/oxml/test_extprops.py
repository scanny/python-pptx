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

    # -- string-valued fields --------------------------------------------

    @pytest.mark.parametrize(
        ("prop_name", "tag_local", "value"),
        [
            ("application_text", "Application", "Microsoft Office PowerPoint"),
            ("app_version_text", "AppVersion", "16.0000"),
            ("company_text", "Company", "Acme Corp"),
            ("hyperlink_base_text", "HyperlinkBase", "https://example.com/docs/"),
            ("manager_text", "Manager", "Alice Jones"),
            ("presentation_format_text", "PresentationFormat", "Widescreen"),
            ("template_text", "Template", "Default Theme.potx"),
        ],
    )
    def it_reads_string_fields_when_present(
        self, prop_name: str, tag_local: str, value: str
    ):
        props = self._props_with_child_xml(tag_local, value)

        assert getattr(props, prop_name) == value

    @pytest.mark.parametrize(
        "prop_name",
        [
            "application_text",
            "app_version_text",
            "company_text",
            "hyperlink_base_text",
            "manager_text",
            "presentation_format_text",
            "template_text",
        ],
    )
    def it_returns_empty_string_when_field_is_absent(self, prop_name: str):
        props = CT_ExtendedProperties.new_extendedProperties()

        assert getattr(props, prop_name) == ""

    @pytest.mark.parametrize(
        ("prop_name", "tag_local"),
        [
            ("application_text", "Application"),
            ("app_version_text", "AppVersion"),
            ("company_text", "Company"),
            ("hyperlink_base_text", "HyperlinkBase"),
            ("manager_text", "Manager"),
            ("presentation_format_text", "PresentationFormat"),
            ("template_text", "Template"),
        ],
    )
    def it_creates_the_child_element_on_first_write(
        self, prop_name: str, tag_local: str
    ):
        props = CT_ExtendedProperties.new_extendedProperties()
        # -- element is absent to start with --
        qname = (
            "{http://schemas.openxmlformats.org/officeDocument/2006/"
            "extended-properties}%s" % tag_local
        )
        assert props.find(qname) is None

        setattr(props, prop_name, "hello")

        element = props.find(qname)
        assert element is not None
        assert element.text == "hello"

    def it_overwrites_an_existing_field_text_on_write(self):
        props = self._props_with_child_xml("Company", "Old Corp")

        props.company_text = "New Corp"

        assert props.company_text == "New Corp"

    def it_raises_TypeError_on_non_string_write(self):
        props = CT_ExtendedProperties.new_extendedProperties()

        with pytest.raises(TypeError, match="Company"):
            props.company_text = 42  # type: ignore[assignment]

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

    @staticmethod
    def _props_with_child_xml(tag_local: str, text: str) -> CT_ExtendedProperties:
        xml = (
            '<Properties xmlns="http://schemas.openxmlformats.org/'
            'officeDocument/2006/extended-properties">'
            "<%s>%s</%s>"
            "</Properties>"
        ) % (tag_local, text, tag_local)
        return parse_xml(xml)
