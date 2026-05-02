"""Unit-test suite for `pptx.oxml.slide` module."""

from __future__ import annotations

import pytest

from pptx.oxml.ns import qn
from pptx.oxml.slide import (
    CT_HeaderFooter,
    CT_NotesMaster,
    CT_NotesSlide,
    CT_SideDirectionTransition,
    CT_Slide,
    CT_SlideTransition,
    CT_TransitionMorph,
    CT_TransitionVariant,
)

from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

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


class DescribeCT_Slide_showMasterSp(object):
    """Unit-test suite for the `showMasterSp` descriptor on CT_Slide (issue #845)."""

    @pytest.mark.parametrize(
        ("cxml", "expected"),
        [
            # -- attribute absent => schema default True --
            ("p:sld/p:cSld", True),
            ('p:sld{showMasterSp=1}/p:cSld', True),
            ('p:sld{showMasterSp=true}/p:cSld', True),
            ('p:sld{showMasterSp=0}/p:cSld', False),
            ('p:sld{showMasterSp=false}/p:cSld', False),
        ],
    )
    def it_reads_the_showMasterSp_attribute(self, cxml: str, expected: bool):
        sld = element(cxml)
        assert sld.showMasterSp is expected

    def it_can_write_showMasterSp_false(self):
        sld = element("p:sld/p:cSld")
        sld.showMasterSp = False
        assert sld.xml == xml('p:sld{showMasterSp=0}/p:cSld')

    def it_removes_showMasterSp_when_set_to_default_true(self):
        sld = element('p:sld{showMasterSp=0}/p:cSld')
        sld.showMasterSp = True
        assert sld.xml == xml("p:sld/p:cSld")


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

    @pytest.mark.parametrize("option", ["byObject", "byWord", "byChar"])
    def it_accepts_any_of_the_three_valid_options(self, option: str):
        from pptx.oxml.ns import nsdecls

        from pptx.oxml import parse_xml

        morph = parse_xml("<p14:morph %s/>" % nsdecls("p14"))
        morph.option = option
        assert morph.option == option

    def it_rejects_an_invalid_option(self):
        from pptx.oxml.ns import nsdecls

        from pptx.oxml import parse_xml

        morph = parse_xml("<p14:morph %s/>" % nsdecls("p14"))
        with pytest.raises(ValueError):
            morph.option = "byParagraph"


class DescribeCT_SideDirectionTransition(object):
    """Unit-test suite for `pptx.oxml.slide.CT_SideDirectionTransition`."""

    @pytest.mark.parametrize("cxml", ["p:wipe", "p:push"])
    def it_is_the_class_used_for_wipe_and_push(self, cxml: str):
        variant = element(cxml)
        assert isinstance(variant, CT_SideDirectionTransition)

    def it_defaults_dir_to_l(self):
        wipe = element("p:wipe")
        assert wipe.dir == "l"

    def it_reads_an_explicit_dir(self):
        wipe = element("p:wipe{dir=r}")
        assert wipe.dir == "r"

    def it_can_write_dir(self):
        wipe = element("p:wipe")
        wipe.dir = "u"
        assert wipe.xml == xml("p:wipe{dir=u}")


class DescribeCT_Slide_transition_alt_content(object):
    """Unit-test suite for the mc:AlternateContent helpers (issue #942).

    See ``CT_Slide.transition_effective`` /
    ``wrap_transition_in_alt_content`` / ``unwrap_transition_from_alt_content``.
    """

    def it_returns_the_direct_transition_as_effective(self):
        sld = element("p:sld/(p:cSld/p:spTree,p:transition/p:fade)")
        assert sld.transition_effective is not None
        assert sld.transition_effective.variant_tag == "p:fade"
        assert sld.transition_is_alt_content_wrapped is False

    def it_returns_None_effective_for_a_slide_with_no_transition(self):
        sld = element("p:sld/p:cSld/p:spTree")
        assert sld.transition_effective is None
        assert sld.transition_is_alt_content_wrapped is False

    def it_resolves_a_wrapped_transition(self):
        from pptx.oxml.ns import nsdecls

        from pptx.oxml import parse_xml

        xml_str = (
            "<p:sld %s>\n"
            "  <p:cSld><p:spTree/></p:cSld>\n"
            '  <mc:AlternateContent>\n'
            '    <mc:Choice Requires="p14">\n'
            '      <p:transition spd="slow">\n'
            '        <p14:morph option="byWord"/>\n'
            "      </p:transition>\n"
            "    </mc:Choice>\n"
            "    <mc:Fallback>\n"
            "      <p:transition><p:fade/></p:transition>\n"
            "    </mc:Fallback>\n"
            "  </mc:AlternateContent>\n"
            "</p:sld>\n"
        ) % nsdecls("p", "p14", "mc")
        sld = parse_xml(xml_str)

        assert sld.transition is None
        assert sld.transition_effective is not None
        assert sld.transition_effective.variant_tag == "p14:morph"
        assert sld.transition_is_alt_content_wrapped is True

    def it_can_wrap_a_direct_transition_preserving_attrs(self):
        sld = element(
            "p:sld/(p:cSld/p:spTree,p:transition{spd=med,advClick=0}/p:fade)"
        )
        inner = sld.wrap_transition_in_alt_content()
        assert sld.transition is None
        assert sld.transition_is_alt_content_wrapped is True
        # -- inner is the mc:Choice/p:transition; attributes preserved --
        assert inner.spd == "med"
        assert inner.advClick is False
        # -- mc:Fallback carries a p:transition/p:fade that mirrors attrs --
        ac = sld._transition_alt_content_lst[0]
        fallback = ac.find(qn("mc:Fallback"))
        assert fallback is not None
        fb_transition = fallback.find(qn("p:transition"))
        assert fb_transition is not None
        assert fb_transition.find(qn("p:fade")) is not None
        assert fb_transition.get("spd") == "med"
        assert fb_transition.get("advClick") == "0"

    def it_is_idempotent_on_wrap(self):
        sld = element("p:sld/(p:cSld/p:spTree,p:transition/p:fade)")
        first = sld.wrap_transition_in_alt_content()
        second = sld.wrap_transition_in_alt_content()
        assert first is second
        # -- only one mc:AlternateContent wrapper in the transition slot --
        assert len(sld._transition_alt_content_lst) == 1

    def it_can_unwrap_a_wrapped_transition(self):
        sld = element("p:sld/(p:cSld/p:spTree,p:transition/p:fade)")
        sld.wrap_transition_in_alt_content()
        assert sld.transition_is_alt_content_wrapped is True

        sld.unwrap_transition_from_alt_content()
        assert sld.transition_is_alt_content_wrapped is False
        assert sld.transition is not None
        assert sld.transition.variant_tag == "p:fade"
        # -- and no mc:AlternateContent remains in transition slot --
        assert sld._transition_alt_content_lst == []

    def it_is_a_noop_to_unwrap_when_not_wrapped(self):
        sld = element("p:sld/(p:cSld/p:spTree,p:transition/p:fade)")
        result = sld.unwrap_transition_from_alt_content()
        assert result is sld.transition
        assert sld.transition_is_alt_content_wrapped is False

    def it_removes_both_forms_on_remove_transition_effective(self):
        sld = element("p:sld/(p:cSld/p:spTree,p:transition/p:fade)")
        sld.wrap_transition_in_alt_content()
        sld.remove_transition_effective()
        assert sld.transition_effective is None
        assert sld.transition_is_alt_content_wrapped is False


class DescribeCT_Slide_childTnLst(object):
    """Regression suite for `pptx.oxml.slide.CT_Slide.get_or_add_childTnLst` (issue #954).

    The helper must locate an existing `p:timing` element whether it
    sits as a plain direct child of `p:sld` or is wrapped inside an
    ``mc:AlternateContent``/``mc:Choice`` block (the form PowerPoint
    emits when timing content references a 2010+ extension such as a
    morph trigger). Before the fix, the wrapped case was not detected
    and :meth:`SlideShapes.add_movie` appended a second, orphan
    ``p:timing`` element alongside the wrapper.
    """

    def it_reuses_an_unwrapped_existing_timing_childTnLst(self):
        sld = parse_xml(
            "<p:sld %s>\n"
            "  <p:cSld><p:spTree/></p:cSld>\n"
            "  <p:timing>\n"
            "    <p:tnLst>\n"
            "      <p:par>\n"
            '        <p:cTn id="1" nodeType="tmRoot">\n'
            "          <p:childTnLst/>\n"
            "        </p:cTn>\n"
            "      </p:par>\n"
            "    </p:tnLst>\n"
            "  </p:timing>\n"
            "</p:sld>" % nsdecls("p")
        )
        existing_childTnLst = sld.xpath(".//p:childTnLst")[0]

        childTnLst = sld.get_or_add_childTnLst()

        assert childTnLst is existing_childTnLst
        assert len(sld.xpath("./p:timing")) == 1

    def it_finds_childTnLst_inside_mc_AlternateContent_wrapping(self):
        """Regression: #954 — mc:AlternateContent-wrapped timing was hidden."""
        sld = parse_xml(
            "<p:sld %s>\n"
            "  <p:cSld><p:spTree/></p:cSld>\n"
            "  <mc:AlternateContent>\n"
            '    <mc:Choice xmlns:p14="http://schemas.microsoft.com/office/po'
            'werpoint/2010/main" Requires="p14">\n'
            "      <p:timing>\n"
            "        <p:tnLst>\n"
            "          <p:par>\n"
            '            <p:cTn id="1" nodeType="tmRoot">\n'
            "              <p:childTnLst/>\n"
            "            </p:cTn>\n"
            "          </p:par>\n"
            "        </p:tnLst>\n"
            "      </p:timing>\n"
            "    </mc:Choice>\n"
            "    <mc:Fallback>\n"
            "      <p:timing/>\n"
            "    </mc:Fallback>\n"
            "  </mc:AlternateContent>\n"
            "</p:sld>" % nsdecls("p", "mc")
        )
        # -- the pre-existing childTnLst lives inside mc:Choice --
        wrapped_childTnLst = sld.xpath(".//mc:Choice/p:timing//p:childTnLst")[0]

        childTnLst = sld.get_or_add_childTnLst()

        # -- the same existing element is returned (no fresh timing created) --
        assert childTnLst is wrapped_childTnLst
        # -- and there is NO orphan p:timing added as a direct child of p:sld --
        assert sld.find(qn("p:timing")) is None
        # -- the mc:AlternateContent wrapper is left in place --
        assert sld.find(qn("mc:AlternateContent")) is not None
        # -- and the whole slide still has exactly one `p:timing` anywhere --
        assert len(sld.xpath(".//p:timing")) == 2  # one in Choice, one in Fallback
        # -- counting only Choice's timing gives exactly one --
        assert len(sld.xpath("./mc:AlternateContent/mc:Choice/p:timing")) == 1

    def it_replaces_a_wrapped_timing_inside_the_wrapper_when_structure_mismatches(
        self,
    ):
        """When wrapped timing lacks the p:tnLst/p:par/p:cTn/p:childTnLst path,
        the freshly-built replacement stays inside the mc:Choice wrapper."""
        sld = parse_xml(
            "<p:sld %s>\n"
            "  <p:cSld><p:spTree/></p:cSld>\n"
            "  <mc:AlternateContent>\n"
            '    <mc:Choice xmlns:p14="http://schemas.microsoft.com/office/po'
            'werpoint/2010/main" Requires="p14">\n'
            "      <p:timing>\n"
            "        <p:bldLst/>\n"
            "      </p:timing>\n"
            "    </mc:Choice>\n"
            "  </mc:AlternateContent>\n"
            "</p:sld>" % nsdecls("p", "mc")
        )

        childTnLst = sld.get_or_add_childTnLst()

        # -- no bare p:timing was added to p:sld as a sibling of the wrapper --
        assert sld.find(qn("p:timing")) is None
        # -- the wrapper still exists, and its mc:Choice now owns a
        # --- fresh p:timing with the required descendant ---
        wrapped_timings = sld.xpath("./mc:AlternateContent/mc:Choice/p:timing")
        assert len(wrapped_timings) == 1
        assert wrapped_timings[0].find(qn("p:tnLst")) is not None
        # -- returned childTnLst is the one inside the new timing ---
        assert childTnLst.getparent().tag == qn("p:cTn")

    def it_creates_a_fresh_timing_when_none_present(self):
        sld = element("p:sld/p:cSld/p:spTree")

        childTnLst = sld.get_or_add_childTnLst()

        assert sld.find(qn("p:timing")) is not None
        # -- newly created `p:timing` sits at the schema-correct position
        #    (after p:cSld / p:clrMapOvr / p:transition) ---
        timing = sld.find(qn("p:timing"))
        assert timing.getparent() is sld
        assert childTnLst.getparent().tag == qn("p:cTn")
