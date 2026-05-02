"""Unit-test suite for `pptx.oxml.slide` module."""

from __future__ import annotations

import pytest

from pptx.oxml.ns import qn
from pptx.oxml.slide import (
    CT_HeaderFooter,
    CT_NotesMaster,
    CT_NotesSlide,
    CT_Slide,
    CT_SlideTransition,
    CT_TransitionMorph,
    CT_TransitionVariant,
)

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


class DescribeCT_Slide_transition(object):
    """Unit-test suite for the new `p:transition` descriptor on CT_Slide (F8)."""

    def it_has_a_None_transition_when_absent(self):
        sld = element("p:sld/p:cSld/p:spTree")
        assert sld.transition is None

    def it_exposes_the_transition_child_element_when_present(self):
        sld = element("p:sld/(p:cSld/p:spTree,p:transition)")
        assert isinstance(sld.transition, CT_SlideTransition)

    def it_can_add_a_transition_child_in_the_right_position(self):
        sld = element("p:sld/(p:cSld/p:spTree,p:timing)")
        transition = sld.get_or_add_transition()
        # -- p:transition must precede p:timing in p:sld's sequence --
        children = list(sld)
        transition_idx = children.index(transition)
        timing_idx = children.index(sld.find(qn("p:timing")))
        assert transition_idx < timing_idx


class DescribeCT_SlideTransition(object):
    """Unit-test suite for `pptx.oxml.slide.CT_SlideTransition`."""

    def it_is_used_for_the_p_transition_element(self):
        transition = element("p:transition")
        assert isinstance(transition, CT_SlideTransition)

    # -- @spd attribute -----------------------------------------------

    def it_defaults_spd_to_fast(self):
        transition = element("p:transition")
        assert transition.spd == "fast"

    def it_reads_explicit_spd(self):
        transition = element("p:transition{spd=med}")
        assert transition.spd == "med"

    def it_can_write_spd(self):
        transition = element("p:transition")
        transition.spd = "slow"
        assert transition.xml == xml("p:transition{spd=slow}")

    # -- @advClick attribute -----------------------------------------

    def it_defaults_advClick_to_True(self):
        transition = element("p:transition")
        assert transition.advClick is True

    def it_reads_explicit_advClick(self):
        transition = element("p:transition{advClick=0}")
        assert transition.advClick is False

    def it_can_write_advClick(self):
        transition = element("p:transition")
        transition.advClick = False
        assert transition.xml == xml("p:transition{advClick=0}")

    # -- @advTm attribute --------------------------------------------

    def it_has_None_advTm_when_absent(self):
        transition = element("p:transition")
        assert transition.advTm is None

    def it_reads_explicit_advTm(self):
        transition = element("p:transition{advTm=5000}")
        assert transition.advTm == 5000

    def it_can_write_advTm(self):
        transition = element("p:transition")
        transition.advTm = 2500
        assert transition.xml == xml("p:transition{advTm=2500}")

    # -- variant child helpers ---------------------------------------

    def it_returns_None_variant_tag_when_no_variant(self):
        transition = element("p:transition")
        assert transition.variant_tag is None

    def it_returns_p_fade_for_a_fade_child(self):
        transition = element("p:transition/p:fade")
        assert transition.variant_tag == "p:fade"

    def it_returns_p_push_for_a_push_child(self):
        transition = element("p:transition/p:push")
        assert transition.variant_tag == "p:push"

    def it_returns_p14_morph_for_a_morph_child(self):
        # -- p14:morph requires p14 nsdecl; element() accepts p14 prefix --
        from pptx.oxml.ns import nsdecls

        from pptx.oxml import parse_xml

        xml_str = (
            '<p:transition %s><p14:morph option="byObject"/></p:transition>'
            % nsdecls("p", "p14")
        )
        transition = parse_xml(xml_str)
        assert transition.variant_tag == "p14:morph"

    def it_can_set_a_p_variant(self):
        transition = element("p:transition")
        new_child = transition.set_variant("p:fade")
        assert new_child.tag == qn("p:fade")
        assert transition.variant_tag == "p:fade"

    def it_can_set_a_p14_morph_variant(self):
        transition = element("p:transition")
        new_child = transition.set_variant("p14:morph")
        assert new_child.tag == qn("p14:morph")
        assert transition.variant_tag == "p14:morph"
        assert isinstance(new_child, CT_TransitionMorph)

    def it_replaces_the_existing_variant_on_set(self):
        transition = element("p:transition/p:fade")
        transition.set_variant("p:push")
        assert transition.variant_tag == "p:push"
        # -- and no fade child lingers --
        assert transition.find(qn("p:fade")) is None

    def it_removes_the_variant_on_set_None(self):
        transition = element("p:transition/p:fade")
        transition.set_variant(None)
        assert transition.variant_tag is None

    # -- p14:dur attribute -------------------------------------------

    def it_returns_None_dur_when_absent(self):
        transition = element("p:transition")
        assert transition.dur is None

    def it_reads_a_written_dur(self):
        transition = element("p:transition")
        transition.dur = 800
        assert transition.dur == 800

    def it_removes_dur_when_set_to_None(self):
        transition = element("p:transition")
        transition.dur = 800
        transition.dur = None
        assert transition.dur is None
        assert qn("p14:dur") not in transition.attrib

    def it_rejects_negative_dur(self):
        transition = element("p:transition")
        with pytest.raises(ValueError, match="non-negative"):
            transition.dur = -1


class DescribeCT_TransitionVariant(object):
    """Unit-test suite for `pptx.oxml.slide.CT_TransitionVariant`."""

    @pytest.mark.parametrize(
        "cxml",
        [
            "p:fade",
            "p:push",
            "p:wipe",
            "p:cover",
            "p:circle",
            "p:dissolve",
            "p:wheel",
            "p:zoom",
            "p:plus",
        ],
    )
    def it_is_registered_for_every_variant_tag(self, cxml: str):
        variant = element(cxml)
        assert isinstance(variant, CT_TransitionVariant)


class DescribeCT_TransitionMorph(object):
    """Unit-test suite for `pptx.oxml.slide.CT_TransitionMorph`."""

    def it_is_the_class_used_for_p14_morph(self):
        from pptx.oxml.ns import nsdecls

        from pptx.oxml import parse_xml

        morph = parse_xml('<p14:morph %s option="byWord"/>' % nsdecls("p14"))
        assert isinstance(morph, CT_TransitionMorph)
        assert morph.option == "byWord"

    def it_defaults_option_to_byObject(self):
        from pptx.oxml.ns import nsdecls

        from pptx.oxml import parse_xml

        morph = parse_xml("<p14:morph %s/>" % nsdecls("p14"))
        assert morph.option == "byObject"
