"""Unit-test suite for `pptx.oxml.slide` module."""

from __future__ import annotations

import pytest

from pptx.oxml.slide import CT_HeaderFooter, CT_NotesMaster, CT_NotesSlide

from ..unitutil.cxml import element, xml
from ..unitutil.file import snippet_text


class DescribeCT_HeaderFooter(object):
    """Unit-test suite for `pptx.oxml.slide.CT_HeaderFooter` objects."""

    @pytest.mark.parametrize(
        ("cxml", "expected"),
        [
            ("p:hf", (True, True, True, True)),
            ("p:hf{sldNum=0}", (False, True, True, True)),
            ("p:hf{hdr=0}", (True, False, True, True)),
            ("p:hf{ftr=0}", (True, True, False, True)),
            ("p:hf{dt=0}", (True, True, True, False)),
            ("p:hf{sldNum=0,hdr=0,ftr=0,dt=0}", (False, False, False, False)),
        ],
    )
    def it_reads_the_four_visibility_attributes(self, cxml: str, expected: tuple[bool, ...]):
        hf = element(cxml)
        assert (hf.sldNum, hf.hdr, hf.ftr, hf.dt) == expected

    def it_can_set_the_sldNum_attribute(self):
        hf = element("p:hf")
        hf.sldNum = False
        assert hf.xml == xml("p:hf{sldNum=0}")

    def it_defaults_every_visibility_flag_to_true_when_absent(self):
        hf = CT_HeaderFooter  # class-level default check via an empty element
        hf_element = element("p:hf")
        for attr_name in ("sldNum", "hdr", "ftr", "dt"):
            assert getattr(hf_element, attr_name) is True, (
                f"{hf.__name__}.{attr_name} should default to True"
            )


class DescribeCT_SlideMaster(object):
    """Unit-test suite for `pptx.oxml.slide.CT_SlideMaster` header-footer handling."""

    def it_adds_a_hf_child_on_get_or_add_hf(self):
        sldMaster = element("p:sldMaster/p:cSld/p:spTree")
        hf = sldMaster.get_or_add_hf()
        assert hf is sldMaster.hf
        assert hf.getparent() is sldMaster


class DescribeCT_SlideLayout(object):
    """Unit-test suite for `pptx.oxml.slide.CT_SlideLayout` header-footer handling."""

    def it_adds_a_hf_child_on_get_or_add_hf(self):
        sldLayout = element("p:sldLayout/p:cSld/p:spTree")
        hf = sldLayout.get_or_add_hf()
        assert hf is sldLayout.hf


class DescribeCT_NotesMaster(object):
    """Unit-test suite for `pptx.oxml.slide.CT_NotesMaster` objects."""

    def it_can_create_a_default_notesMaster_element(self):
        notesMaster = CT_NotesMaster.new_default()
        assert notesMaster.xml == snippet_text("default-notesMaster")

    def it_adds_a_hf_child_on_get_or_add_hf(self):
        notesMaster = element("p:notesMaster/p:cSld/p:spTree")
        hf = notesMaster.get_or_add_hf()
        assert hf is notesMaster.hf


class DescribeCT_NotesSlide(object):
    """Unit-test suite for `pptx.oxml.slide.CT_NotesSlide` objects."""

    def it_can_create_a_new_notes_element(self):
        notes = CT_NotesSlide.new()
        assert notes.xml == snippet_text("default-notes")
