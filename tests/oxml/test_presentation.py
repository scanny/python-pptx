# pyright: reportPrivateUsage=false

"""Unit-test suite for `pptx.oxml.presentation` module."""

from __future__ import annotations

from typing import cast

import pytest

from pptx.oxml.presentation import (
    CT_EmbeddedFontList,
    CT_EmbeddedFontListEntry,
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
