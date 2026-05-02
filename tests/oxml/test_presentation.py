# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.oxml.presentation` module."""

from __future__ import annotations

from typing import cast

import pytest

from pptx.oxml.presentation import (
    CT_EmbeddedFontList,
    CT_EmbeddedFontListEntry,
    CT_ExtensionList,
    CT_Presentation,
    CT_Section,
    CT_SectionList,
    CT_SectionSlideIdList,
    CT_SlideIdList,
)

from ..unitutil.cxml import element, xml


class DescribeCT_SlideIdList(object):
    """Unit-test suite for `pptx.oxml.presentation.CT_SlideIdLst` objects."""

    def it_can_add_a_sldId_element_as_a_child(self):
        sldIdLst = cast(CT_SlideIdList, element("p:sldIdLst/p:sldId{r:id=rId4,id=256}"))

        sldIdLst.add_sldId("rId1")

        assert sldIdLst.xml == xml(
            "p:sldIdLst/(p:sldId{r:id=rId4,id=256},p:sldId{r:id=rId1,id=257})"
        )

    @pytest.mark.parametrize(
        ("sldIdLst_cxml", "expected_value"),
        [
            ("p:sldIdLst", 256),
            ("p:sldIdLst/p:sldId{id=42}", 256),
            ("p:sldIdLst/p:sldId{id=256}", 257),
            ("p:sldIdLst/(p:sldId{id=256},p:sldId{id=712})", 713),
            ("p:sldIdLst/(p:sldId{id=280},p:sldId{id=257})", 281),
        ],
    )
    def it_knows_the_next_available_slide_id(self, sldIdLst_cxml: str, expected_value: int):
        sldIdLst = cast(CT_SlideIdList, element(sldIdLst_cxml))
        assert sldIdLst._next_id == expected_value

    @pytest.mark.parametrize(
        ("sldIdLst_cxml", "expected_value"),
        [
            ("p:sldIdLst/p:sldId{id=2147483646}", 2147483647),
            ("p:sldIdLst/p:sldId{id=2147483647}", 256),
            # -- 2147483648 is not a valid id but shouldn't stop us from finding a one that is --
            ("p:sldIdLst/p:sldId{id=2147483648}", 256),
            ("p:sldIdLst/(p:sldId{id=256},p:sldId{id=2147483647})", 257),
            ("p:sldIdLst/(p:sldId{id=256},p:sldId{id=2147483647},p:sldId{id=257})", 258),
            # -- 245 is also not a valid id but that shouldn't change the result either --
            ("p:sldIdLst/(p:sldId{id=245},p:sldId{id=2147483647},p:sldId{id=256})", 257),
        ],
    )
    def and_it_chooses_a_valid_slide_id_when_max_slide_id_is_used_for_a_slide(
        self, sldIdLst_cxml: str, expected_value: int
    ):
        sldIdLst = cast(CT_SlideIdList, element(sldIdLst_cxml))

        slide_id = sldIdLst._next_id

        assert 256 <= slide_id <= 2147483647
        assert slide_id == expected_value


class DescribeCT_EmbeddedFontList(object):
    """Unit-test suite for `pptx.oxml.presentation.CT_EmbeddedFontList`."""

    def it_can_add_an_embeddedFont_child_with_typeface(self):
        embeddedFontLst = cast(CT_EmbeddedFontList, element("p:embeddedFontLst"))

        embeddedFontLst.add_embeddedFont("Pacifico")

        assert embeddedFontLst.xml == xml(
            "p:embeddedFontLst/p:embeddedFont/p:font{typeface=Pacifico}"
        )

    def it_can_find_an_existing_entry_by_typeface(self):
        cxml = (
            "p:embeddedFontLst/(p:embeddedFont/p:font{typeface=A},"
            "p:embeddedFont/p:font{typeface=B})"
        )
        embeddedFontLst = cast(CT_EmbeddedFontList, element(cxml))

        entry = embeddedFontLst.entry_for_typeface("B")

        assert entry is not None
        assert entry.typeface == "B"

    def but_it_returns_None_when_no_entry_matches(self):
        cxml = "p:embeddedFontLst/p:embeddedFont/p:font{typeface=A}"
        embeddedFontLst = cast(CT_EmbeddedFontList, element(cxml))

        assert embeddedFontLst.entry_for_typeface("Missing") is None


class DescribeCT_EmbeddedFontListEntry(object):
    """Unit-test suite for `pptx.oxml.presentation.CT_EmbeddedFontListEntry`."""

    def it_knows_its_typeface(self):
        entry = cast(
            CT_EmbeddedFontListEntry,
            element("p:embeddedFont/p:font{typeface=Pacifico}"),
        )
        assert entry.typeface == "Pacifico"

    def it_can_change_its_typeface(self):
        entry = cast(
            CT_EmbeddedFontListEntry,
            element("p:embeddedFont/p:font{typeface=Pacifico}"),
        )

        entry.typeface = "Roboto"

        assert entry.typeface == "Roboto"

    def it_returns_the_rId_for_a_populated_style_slot(self):
        entry = cast(
            CT_EmbeddedFontListEntry,
            element("p:embeddedFont/(p:font{typeface=Pacifico},p:regular{r:id=rId7})"),
        )
        assert entry.rId_for_style("regular") == "rId7"

    def but_it_returns_None_for_an_empty_style_slot(self):
        entry = cast(
            CT_EmbeddedFontListEntry,
            element("p:embeddedFont/p:font{typeface=Pacifico}"),
        )
        assert entry.rId_for_style("bold") is None

    @pytest.mark.parametrize("style", ["regular", "bold", "italic", "boldItalic"])
    def it_can_set_the_rId_for_any_style(self, style: str):
        entry = cast(
            CT_EmbeddedFontListEntry,
            element("p:embeddedFont/p:font{typeface=Pacifico}"),
        )

        entry.set_rId_for_style(style, "rId3")

        assert entry.rId_for_style(style) == "rId3"

    def and_it_overwrites_an_existing_rId_on_the_same_style(self):
        entry = cast(
            CT_EmbeddedFontListEntry,
            element("p:embeddedFont/(p:font{typeface=Pacifico},p:regular{r:id=rId1})"),
        )

        entry.set_rId_for_style("regular", "rId9")

        assert entry.rId_for_style("regular") == "rId9"


class DescribeCT_ExtensionList(object):
    """Unit-test suite for `pptx.oxml.presentation.CT_ExtensionList`."""

    def it_can_find_an_ext_child_by_uri(self):
        extLst = cast(
            CT_ExtensionList,
            element("p:extLst/(p:ext{uri=FOO},p:ext{uri=BAR})"),
        )

        ext = extLst.get_ext_by_uri("BAR")

        assert ext is not None
        assert ext.uri == "BAR"

    def but_it_returns_None_when_uri_is_not_found(self):
        extLst = cast(CT_ExtensionList, element("p:extLst/p:ext{uri=FOO}"))
        assert extLst.get_ext_by_uri("NOT-THERE") is None

    def it_can_add_an_ext_child_with_a_given_uri(self):
        extLst = cast(CT_ExtensionList, element("p:extLst"))

        ext = extLst.get_or_add_ext_by_uri("FOO")

        assert ext.uri == "FOO"
        assert extLst.xml == xml("p:extLst/p:ext{uri=FOO}")

    def and_it_reuses_an_existing_ext_with_the_same_uri(self):
        extLst = cast(CT_ExtensionList, element("p:extLst/p:ext{uri=FOO}"))
        first = extLst.get_or_add_ext_by_uri("FOO")

        second = extLst.get_or_add_ext_by_uri("FOO")

        assert first is second
        assert len(extLst.ext_lst) == 1


class DescribeCT_Presentation_sections(object):
    """Unit-test suite for section-related helpers on `CT_Presentation`."""

    def it_returns_None_when_no_sectionLst_is_present(self):
        prs = cast(CT_Presentation, element("p:presentation"))
        assert prs.sectionLst is None

    def but_an_unrelated_ext_does_not_fool_it(self):
        prs = cast(
            CT_Presentation,
            element("p:presentation/p:extLst/p:ext{uri=SOMETHING-ELSE}"),
        )
        assert prs.sectionLst is None

    def it_finds_a_sectionLst_nested_under_the_correct_ext_uri(self):
        # -- The SectionList extension URI has curly braces in the real-world form
        # -- "{521415D9-…}"; cxml does not accept braces in attribute values, so we
        # -- assemble the XML directly rather than through cxml.
        from pptx.oxml import parse_xml

        prs_xml = (
            '<p:presentation xmlns:p="http://schemas.openxmlformats.org/'
            'presentationml/2006/main" xmlns:p14="http://schemas.microsoft.com/'
            'office/powerpoint/2010/main">'
            '<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">'
            "<p14:sectionLst/></p:ext></p:extLst></p:presentation>"
        )
        prs = cast(CT_Presentation, parse_xml(prs_xml))
        assert prs.sectionLst is not None

    def it_can_create_the_entire_extLst_ext_sectionLst_chain(self):
        prs = cast(CT_Presentation, element("p:presentation"))

        sectionLst = prs.get_or_add_sectionLst()

        assert sectionLst is not None
        ext = sectionLst.getparent()
        assert ext is not None
        assert ext.uri == "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"

    def and_it_reuses_existing_scaffolding_on_subsequent_calls(self):
        prs = cast(CT_Presentation, element("p:presentation"))

        first = prs.get_or_add_sectionLst()
        second = prs.get_or_add_sectionLst()

        assert first is second


class DescribeCT_SectionList(object):
    """Unit-test suite for `pptx.oxml.presentation.CT_SectionList`."""

    def it_can_add_a_section_with_a_given_name_and_id(self):
        sectionLst = cast(CT_SectionList, element("p14:sectionLst"))

        sectionLst.add_section("Intro", "SECTION-ID-1")

        assert sectionLst.xml == xml(
            "p14:sectionLst/p14:section{name=Intro,id=SECTION-ID-1}/p14:sldIdLst"
        )

    def it_exposes_sections_as_a_list(self):
        cxml = (
            "p14:sectionLst/(p14:section{name=A,id=SECTION-ID-A}/p14:sldIdLst,"
            "p14:section{name=B,id=SECTION-ID-B}/p14:sldIdLst)"
        )
        sectionLst = cast(CT_SectionList, element(cxml))
        assert [s.name for s in sectionLst.section_lst] == ["A", "B"]


class DescribeCT_Section(object):
    """Unit-test suite for `pptx.oxml.presentation.CT_Section`."""

    def it_knows_its_name_and_id(self):
        section = cast(
            CT_Section,
            element("p14:section{name=Intro,id=SECTION-ID-1}/p14:sldIdLst"),
        )

        assert section.name == "Intro"
        assert section.id == "SECTION-ID-1"

    def it_can_rename_itself(self):
        section = cast(
            CT_Section,
            element("p14:section{name=Intro,id=SECTION-ID-1}/p14:sldIdLst"),
        )

        section.name = "Summary"

        assert section.name == "Summary"


class DescribeCT_SectionSlideIdList(object):
    """Unit-test suite for `pptx.oxml.presentation.CT_SectionSlideIdList`."""

    def it_can_add_a_sldId_referencing_a_slide_by_id(self):
        sldIdLst = cast(CT_SectionSlideIdList, element("p14:sldIdLst"))

        sldIdLst.add_sldId(256)

        assert sldIdLst.xml == xml("p14:sldIdLst/p14:sldId{id=256}")

    def it_preserves_order_of_member_slide_ids(self):
        sldIdLst = cast(
            CT_SectionSlideIdList,
            element("p14:sldIdLst/(p14:sldId{id=256},p14:sldId{id=257},p14:sldId{id=258})"),
        )
        assert [e.id for e in sldIdLst.sldId_lst] == [256, 257, 258]
