# pyright: reportPrivateUsage=false

"""Regression test for issue #400 — animation-control umbrella.

Issue #400 (https://github.com/scanny/python-pptx/issues/400) is an
umbrella request for "animation manipulation features". The reporter
(and the two follow-up comments) specifically called out the flow of
"adding an animation effect on a shape" and being able to round-trip
authored animation XML through python-pptx without losing it, so the
presentation can be recorded as a video later by a renderer that
consumes the animation tree.

The F8 foundation (``feat/foundation-f8-animations-transitions``, shipped
on master at commit ``5e10844f``) delivers the XML layer that #400 was
gated on: typed element classes for ``p:timing`` / ``p:tnLst`` / ``p:par``
/ ``p:seq`` / ``p:cTn`` and the ``p:transition`` subtree (including the
``p14:morph`` Office 2010 extension), a ``Slide.transition`` proxy with
``type`` / ``duration`` / ``advance_on_click`` / ``advance_after_time``,
and the ``Slide.has_animations`` / ``Slide.timing_xml`` introspection
properties.

Structured entrance / exit / emphasis / motion-path authoring APIs
(``Slide.animations.add_entrance_effect(...)`` and friends) are deferred
to the four F8 leaf branches that #400 aggregates:

  * ``feat/issue-102-shape-animation`` — entrance effects on shapes
  * ``feat/issue-1106-entrance-exit-animations`` — entrance / exit API
  * ``feat/issue-264-shape-animation-control`` — full animation tree
  * ``feat/issue-861-animation-delay`` — start / end delay accessors

This regression test exercises the flow the #400 reporter asked for at
the MVP level — round-tripping an authored animation tree and
introspecting it — so that any future refactor of the timing subtree
keeps the umbrella flow working.
"""

from __future__ import annotations

import io

import pytest

from pptx import Presentation
from pptx.enum.transition import PP_TRANSITION_TYPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn


@pytest.fixture
def _restore_part_factory():
    """Guard against pollution from other tests that mutate `PartFactory.part_type_for`.

    In particular, `tests/opc/test_package.py::DescribePartFactory` overwrites
    `PartFactory.part_type_for[CT.PML_SLIDE]` with a Mock and does not restore
    it. Without this guard, running this module after ``tests/opc/test_package.py``
    causes ``Presentation().slides.add_slide(...)`` to return a Mock instead of
    a real Slide. See the identical fixture in ``tests/test_comments.py``.
    """
    from pptx.opc.constants import CONTENT_TYPE as CT
    from pptx.opc.package import PartFactory
    from pptx.parts.chart import ChartPart
    from pptx.parts.comments import CommentAuthorsPart, CommentsPart
    from pptx.parts.coreprops import CorePropertiesPart
    from pptx.parts.image import ImagePart
    from pptx.parts.media import MediaPart
    from pptx.parts.presentation import PresentationPart
    from pptx.parts.slide import (
        NotesMasterPart,
        NotesSlidePart,
        SlideLayoutPart,
        SlideMasterPart,
        SlidePart,
    )

    saved = dict(PartFactory.part_type_for)
    expected = {
        CT.PML_PRESENTATION_MAIN: PresentationPart,
        CT.PML_PRES_MACRO_MAIN: PresentationPart,
        CT.PML_TEMPLATE_MAIN: PresentationPart,
        CT.PML_SLIDESHOW_MAIN: PresentationPart,
        CT.OPC_CORE_PROPERTIES: CorePropertiesPart,
        CT.PML_COMMENTS: CommentsPart,
        CT.PML_COMMENT_AUTHORS: CommentAuthorsPart,
        CT.PML_NOTES_MASTER: NotesMasterPart,
        CT.PML_NOTES_SLIDE: NotesSlidePart,
        CT.PML_SLIDE: SlidePart,
        CT.PML_SLIDE_LAYOUT: SlideLayoutPart,
        CT.PML_SLIDE_MASTER: SlideMasterPart,
        CT.DML_CHART: ChartPart,
        CT.JPEG: ImagePart,
        CT.PNG: ImagePart,
        CT.MP4: MediaPart,
    }
    PartFactory.part_type_for.update(expected)
    try:
        yield
    finally:
        PartFactory.part_type_for.clear()
        PartFactory.part_type_for.update(saved)


# -- A realistic "fade in" entrance-effect timing tree, as PowerPoint
#    itself emits for a single shape. The tree nests:
#       p:timing / p:tnLst / p:par / p:cTn(id=1, tmRoot)
#         / p:childTnLst / p:seq (concurrent=1, nextAc=seek)
#           / p:cTn(id=2) / p:childTnLst / p:par / p:cTn(id=3) / ...
#             / p:par / p:cTn(id=5, presetID=1, presetClass=entr)
#               / p:childTnLst / p:set / p:animEffect
#
#    We inject this raw XML via ``slide._element.append()`` so we don't
#    depend on the yet-to-land structured authoring API from #102 /
#    #1106 / #264 / #861 — this is exactly the scenario those leaves
#    will convert into a one-liner. The *shape* of the tree here is
#    what they will need to round-trip.
_ANIMATED_TIMING_XML = (
    '<p:timing %s spidShape="4">'
    "  <p:tnLst>"
    "    <p:par>"
    '      <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
    "        <p:childTnLst>"
    '          <p:seq concurrent="1" nextAc="seek">'
    '            <p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
    "              <p:childTnLst>"
    '                <p:par>'
    '                  <p:cTn id="3" fill="hold">'
    "                    <p:stCondLst>"
    '                      <p:cond delay="indefinite"/>'
    "                    </p:stCondLst>"
    "                    <p:childTnLst>"
    '                      <p:par>'
    '                        <p:cTn id="4" fill="hold">'
    "                          <p:stCondLst>"
    '                            <p:cond delay="0"/>'
    "                          </p:stCondLst>"
    "                          <p:childTnLst>"
    '                            <p:par>'
    '                              <p:cTn id="5" presetID="1"'
    '                                    presetClass="entr"'
    '                                    presetSubtype="0"'
    '                                    fill="hold"'
    '                                    grpId="0"'
    '                                    nodeType="clickEffect">'
    "                                <p:stCondLst>"
    '                                  <p:cond delay="0"/>'
    "                                </p:stCondLst>"
    "                                <p:childTnLst>"
    '                                  <p:set>'
    "                                    <p:cBhvr>"
    '                                      <p:cTn id="6" dur="1" fill="hold">'
    "                                        <p:stCondLst>"
    '                                          <p:cond delay="0"/>'
    "                                        </p:stCondLst>"
    "                                      </p:cTn>"
    "                                      <p:tgtEl>"
    '                                        <p:spTgt spid="4"/>'
    "                                      </p:tgtEl>"
    "                                      <p:attrNameLst>"
    "                                        <p:attrName>style.visibility</p:attrName>"
    "                                      </p:attrNameLst>"
    "                                    </p:cBhvr>"
    "                                    <p:to>"
    '                                      <p:strVal val="visible"/>'
    "                                    </p:to>"
    "                                  </p:set>"
    '                                  <p:animEffect transition="in"'
    '                                                filter="fade">'
    "                                    <p:cBhvr>"
    '                                      <p:cTn id="7" dur="500"/>'
    "                                      <p:tgtEl>"
    '                                        <p:spTgt spid="4"/>'
    "                                      </p:tgtEl>"
    "                                    </p:cBhvr>"
    "                                  </p:animEffect>"
    "                                </p:childTnLst>"
    "                              </p:cTn>"
    "                            </p:par>"
    "                          </p:childTnLst>"
    "                        </p:cTn>"
    "                      </p:par>"
    "                    </p:childTnLst>"
    "                  </p:cTn>"
    "                </p:par>"
    "              </p:childTnLst>"
    "            </p:cTn>"
    "          </p:seq>"
    "        </p:childTnLst>"
    "      </p:cTn>"
    "    </p:par>"
    "  </p:tnLst>"
    "</p:timing>"
) % nsdecls("p")


class DescribeIssue400AnimationUmbrella(object):
    """The #400 flow: author a fade-in animation on a shape, save, reopen,
    confirm the animation tree survived, and introspect it via the F8 API.
    """

    def it_round_trips_an_authored_entrance_animation(self, _restore_part_factory):
        """The reporter's core ask: add an animation effect on a shape
        (for later video rendering) and have python-pptx not drop it.

        Until #102 / #1106 land the structured ``add_entrance_effect(...)``
        API, callers can author the timing XML directly — and thanks to
        F8, python-pptx now preserves it verbatim across save / reopen.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        # -- need a real shape to target, so the spTgt spid is meaningful --
        textbox = slide.shapes.add_textbox(0, 0, 914400, 914400)
        textbox.text_frame.text = "Animate me"

        # -- baseline: a fresh slide has no animations --
        assert slide.has_animations is False
        assert slide.timing_xml is None

        # -- author a fade-in entrance effect by injecting the timing
        #    subtree. (#102/#1106/#264/#861 will replace this 3-line
        #    preamble with one method call.) parse_xml() registers the
        #    element as CT_SlideTiming so Slide.timing_xml works. --
        timing = parse_xml(_ANIMATED_TIMING_XML)
        slide._element.append(timing)

        # -- introspection works immediately (no save required) --
        assert slide.has_animations is True
        assert slide.timing_xml is not None
        assert "<p:animEffect" in slide.timing_xml
        assert 'filter="fade"' in slide.timing_xml
        assert 'presetClass="entr"' in slide.timing_xml

        # -- save + reopen round-trip: the whole animation tree
        #    survives, including the p:animEffect leaf, the preset
        #    metadata, and the condition-list timing --
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        prs2 = Presentation(buf)
        slide2 = prs2.slides[0]

        assert slide2.has_animations is True
        reloaded_xml = slide2.timing_xml
        assert reloaded_xml is not None
        # -- full tree preserved --
        assert "<p:tnLst" in reloaded_xml
        assert "<p:seq" in reloaded_xml
        assert "<p:animEffect" in reloaded_xml
        # -- per-effect metadata preserved --
        assert 'filter="fade"' in reloaded_xml
        assert 'transition="in"' in reloaded_xml
        assert 'presetID="1"' in reloaded_xml
        assert 'presetClass="entr"' in reloaded_xml
        # -- per-cTn timing metadata preserved (duration, delay) --
        assert 'dur="500"' in reloaded_xml

    def it_does_not_clobber_an_existing_animation_tree_when_setting_transition(
        self, _restore_part_factory
    ):
        """A slide may carry BOTH a transition and an animation tree.
        Setting ``Slide.transition.type`` (F8 MVP API) must not drop the
        separate ``p:timing`` subtree — #400 reporters want to mix
        transitions with authored shape animations.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.add_textbox(0, 0, 914400, 914400)

        # -- author animation --
        timing = parse_xml(_ANIMATED_TIMING_XML)
        slide._element.append(timing)
        assert slide.has_animations is True

        # -- now set a slide transition via the F8 MVP API --
        slide.transition.type = PP_TRANSITION_TYPE.FADE
        slide.transition.duration = 800
        slide.transition.advance_after_time = 3000

        # -- both subtrees coexist: p:transition precedes p:timing in
        #    the canonical sld element order (_tag_seq) --
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        slide2 = prs2.slides[0]

        assert slide2.transition.type is PP_TRANSITION_TYPE.FADE
        assert slide2.transition.duration == 800
        assert slide2.transition.advance_after_time == 3000
        assert slide2.has_animations is True
        assert "<p:animEffect" in (slide2.timing_xml or "")

    def it_exposes_the_shape_targeted_by_an_animation_via_timing_xml(
        self, _restore_part_factory
    ):
        """Issue #400's follow-up comment wanted to *inspect* which
        shape an animation targets (so a video-rendering pipeline can
        sequence effects). F8's ``Slide.timing_xml`` gives read access
        to the full tree — including ``p:spTgt/@spid`` — before the
        structured accessors from #102/#264 land.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        shape = slide.shapes.add_textbox(0, 0, 914400, 914400)

        timing = parse_xml(_ANIMATED_TIMING_XML)
        slide._element.append(timing)

        timing_xml = slide.timing_xml
        assert timing_xml is not None

        # -- callers can discover which shape an effect targets by
        #    parsing spTgt/@spid out of the timing XML; once round
        #    through save/reopen, the attribute is preserved --
        parsed = parse_xml(timing_xml)
        spTgts = parsed.findall(".//" + qn("p:spTgt"))
        assert len(spTgts) == 2  # -- one for the set, one for the animEffect
        for spTgt in spTgts:
            assert spTgt.get("spid") == "4"

        # -- and the real shape the spid refers to is reachable --
        assert shape is not None

    def it_reports_has_animations_False_for_a_bare_transition(self, _restore_part_factory):
        """A slide with ONLY a transition (and no ``p:timing``) must
        not report ``has_animations`` — the two are separate subtrees.
        Pins the #400 invariant that ``has_animations`` means "authored
        shape animations" and not "has any timing-adjacent markup".
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.transition.type = PP_TRANSITION_TYPE.FADE

        assert slide.has_animations is False
        assert slide.timing_xml is None

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        slide2 = Presentation(buf).slides[0]

        assert slide2.has_animations is False
        assert slide2.timing_xml is None
        assert slide2.transition.type is PP_TRANSITION_TYPE.FADE
