# pyright: reportPrivateUsage=false

"""Regression test for issue #954 — ``SlideShapes.add_movie()`` corrupts
a ``.pptx`` when the slide already carries a pre-existing ``p:timing``
wrapped inside ``mc:AlternateContent``/``mc:Choice``.

Issue #954 (https://github.com/scanny/python-pptx/issues/954) reports
that calling ``add_movie()`` on a slide whose timing contains a
non-OpenXML (i.e. 2010+ extension) feature — most commonly the
``p14:morph`` MORPH trigger — produces a file PowerPoint refuses to
open. The root cause is that PowerPoint wraps such ``p:timing`` elements
inside an ``mc:AlternateContent``/``mc:Choice`` block so pre-2010
viewers can fall back gracefully. The pre-fix ``get_or_add_childTnLst``
helper only looked at the direct ``./p:timing`` path, missed the
wrapped timing, and appended a second, orphan ``p:timing`` as a direct
child of ``p:sld``. A slide with two sibling ``p:timing`` elements
violates the schema and PowerPoint rejects the whole file.

Wave 5's commit 95030bb7 ("fix(movie): merge p:video into wrapped
p:timing (#954)") resolved this by teaching ``CT_Slide`` to locate the
existing ``p:timing`` whether it sits as a plain direct child or as a
descendant of ``mc:AlternateContent``/``mc:Choice``, and to merge the
new ``p:video`` into its ``p:childTnLst`` in place. When the wrapped
timing does not match the expected inner shape, the fix replaces it
*inside* the wrapper rather than alongside it, so only one ``p:timing``
ever ends up on the slide.

This verify suite pins the end-to-end resolution from the user's
perspective. Unlike the oxml-level unit tests in
``tests/oxml/test_slide.py`` (``DescribeCT_Slide_childTnLst``), which
cover the helper in isolation, and the behave scenario in
``features/shp-shapes.feature`` (``SlideShapes.add_movie() on a slide
with mc:AlternateContent-wrapped timing``), which covers the SlideShapes
call once through, this test:

  1. Authors a slide with a PowerPoint-realistic ``mc:AlternateContent``
     wrapper carrying a ``p14:morph``-triggered timing.
  2. Calls ``add_movie()`` on that slide.
  3. **Saves the presentation and reloads it** — the critical part of
     the original bug report, since the corruption only surfaced when
     PowerPoint tried to open the saved file.
  4. Asserts that the reloaded slide still has exactly one ``p:timing``
     (wrapped), that the ``p14:morph`` trigger is preserved, and that
     the new movie's ``p:video`` entry has been merged into the wrapped
     timing's ``p:childTnLst`` — not appended as a duplicate sibling.
"""

from __future__ import annotations

import io
from typing import TYPE_CHECKING

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.shapes.picture import Movie
from pptx.util import Inches

if TYPE_CHECKING:
    from pptx.presentation import Presentation as _PresentationT
    from pptx.slide import Slide


def _blank_slide() -> tuple[_PresentationT, Slide]:
    """Return ``(presentation, slide)`` with a fresh blank slide."""
    prs = Presentation()
    return prs, prs.slides.add_slide(prs.slide_layouts[6])  # -- blank layout --


def _inject_mc_wrapped_morph_timing(slide: Slide) -> None:
    """Inject an ``mc:AlternateContent``-wrapped ``p:timing`` onto ``slide``.

    Shapes PowerPoint's actual output for a slide whose timing
    references the 2010 ``p14:morph`` extension: a
    ``mc:Choice Requires="p14"`` holding the real timing (with an
    in-line ``p14:morph`` trigger inside a ``p:cond``) plus a
    ``mc:Fallback`` carrying a plain ``p:timing`` for pre-2010 viewers.
    """
    wrapper_xml = (
        "<mc:AlternateContent %s>\n"
        '  <mc:Choice xmlns:p14="http://schemas.microsoft.com/office/po'
        'werpoint/2010/main" Requires="p14">\n'
        "    <p:timing>\n"
        "      <p:tnLst>\n"
        "        <p:par>\n"
        '          <p:cTn id="1" dur="indefinite" restart="never" '
        'nodeType="tmRoot">\n'
        "            <p:childTnLst>\n"
        '              <p:seq concurrent="1" nextAc="seek">\n'
        '                <p:cTn id="2" dur="indefinite" nodeType="'
        'mainSeq">\n'
        "                  <p:childTnLst/>\n"
        "                </p:cTn>\n"
        "                <p:prevCondLst>\n"
        '                  <p:cond evt="onPrev" delay="0">\n'
        "                    <p:tgtEl>\n"
        "                      <p:sldTgt/>\n"
        "                    </p:tgtEl>\n"
        "                  </p:cond>\n"
        "                </p:prevCondLst>\n"
        "              </p:seq>\n"
        "            </p:childTnLst>\n"
        "          </p:cTn>\n"
        "        </p:par>\n"
        "      </p:tnLst>\n"
        "      <p:extLst>\n"
        '        <p:ext uri="{{DEADBEEF-BEEF-BEEF-BEEF-DEADBEEFBEEF}}">\n'
        '          <p14:morph option="byObject"/>\n'
        "        </p:ext>\n"
        "      </p:extLst>\n"
        "    </p:timing>\n"
        "  </mc:Choice>\n"
        "  <mc:Fallback>\n"
        "    <p:timing>\n"
        "      <p:tnLst>\n"
        "        <p:par>\n"
        '          <p:cTn id="1" dur="indefinite" restart="never" '
        'nodeType="tmRoot"/>\n'
        "        </p:par>\n"
        "      </p:tnLst>\n"
        "    </p:timing>\n"
        "  </mc:Fallback>\n"
        "</mc:AlternateContent>" % nsdecls("p", "mc")
    )
    slide._element.append(parse_xml(wrapper_xml))


def _movie_shape(slide: Slide) -> Movie:
    """Return the single MEDIA shape on ``slide``."""
    media = next(sh for sh in slide.shapes if sh.shape_type == MSO_SHAPE_TYPE.MEDIA)
    assert isinstance(media, Movie)
    return media


class DescribeIssue954MovieCorruptionVerify:
    """#954 verify-and-close: add_movie() on a slide with wrapped timing.

    Exercises the end-to-end scenario Wave 5 resolved: authoring a
    movie on a slide whose pre-existing ``p:timing`` is wrapped in
    ``mc:AlternateContent`` (the form PowerPoint emits for 2010+
    extensions) no longer produces a duplicate ``p:timing`` sibling,
    and the saved ``.pptx`` round-trips cleanly.
    """

    # -- 1: the core regression scenario --------------------------------

    def it_does_not_duplicate_p_timing_when_wrapped_timing_present(self):
        """``add_movie`` must not emit a second ``p:timing`` sibling
        when one already exists inside ``mc:AlternateContent``."""
        _prs, slide = _blank_slide()
        _inject_mc_wrapped_morph_timing(slide)
        video = io.BytesIO(b"\x00" * 128)

        slide.shapes.add_movie(
            video,
            left=Inches(1),
            top=Inches(1),
            width=Inches(4),
            height=Inches(3),
        )

        sld = slide._element
        # -- no direct `./p:timing` child — it stays wrapped --
        assert sld.find(qn("p:timing")) is None
        # -- exactly one wrapper, with exactly one timing inside its Choice --
        wrappers = sld.findall(qn("mc:AlternateContent"))
        assert len(wrappers) == 1
        choice_timings = sld.xpath("./mc:AlternateContent/mc:Choice/p:timing")
        assert len(choice_timings) == 1

    def it_merges_the_new_p_video_into_the_wrapped_childTnLst(self):
        """The new ``p:video`` must land inside the wrapped timing's
        ``p:childTnLst``, not alongside the wrapper as an orphan."""
        _prs, slide = _blank_slide()
        _inject_mc_wrapped_morph_timing(slide)
        video = io.BytesIO(b"\x00" * 128)

        slide.shapes.add_movie(
            video,
            left=Inches(1),
            top=Inches(1),
            width=Inches(4),
            height=Inches(3),
        )

        sld = slide._element
        # -- the new p:video is inside the wrapped p:timing subtree --
        wrapped_videos = sld.xpath("./mc:AlternateContent/mc:Choice/p:timing//p:video")
        assert len(wrapped_videos) == 1
        # -- and the p:video carries the shape id of the movie --
        movie = _movie_shape(slide)
        sp_id_els = wrapped_videos[0].xpath(".//p:spTgt/@spid")
        assert sp_id_els == [str(movie.shape_id)]

    def it_preserves_the_p14_morph_trigger_of_the_wrapped_timing(self):
        """The fix must leave the non-OpenXML ``p14:morph`` extension
        that motivated the wrap in the first place untouched."""
        _prs, slide = _blank_slide()
        _inject_mc_wrapped_morph_timing(slide)
        video = io.BytesIO(b"\x00" * 128)

        slide.shapes.add_movie(
            video,
            left=Inches(1),
            top=Inches(1),
            width=Inches(4),
            height=Inches(3),
        )

        morphs = slide._element.xpath("./mc:AlternateContent/mc:Choice/p:timing//p14:morph")
        assert len(morphs) == 1
        # -- and mc:Fallback is still present (pre-2010 viewer safety net) --
        fallbacks = slide._element.xpath("./mc:AlternateContent/mc:Fallback")
        assert len(fallbacks) == 1

    # -- 2: the original bug signature — round-trip corruption ---------

    def it_produces_a_pptx_that_reopens_without_error(self):
        """Save + reload: the scenario the original issue flagged as
        "corruption" must now round-trip cleanly."""
        prs, slide = _blank_slide()
        _inject_mc_wrapped_morph_timing(slide)
        video = io.BytesIO(b"\x00" * 128)
        slide.shapes.add_movie(
            video,
            left=Inches(1),
            top=Inches(1),
            width=Inches(4),
            height=Inches(3),
        )

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)  # -- must not raise --

        reloaded_sld = prs2.slides[0]._element
        # -- after round-trip the slide still has exactly one p:timing
        # -- (inside the mc:Choice) and the movie's p:video is nested
        # -- inside that timing's childTnLst tree --
        assert reloaded_sld.find(qn("p:timing")) is None
        choice_timings = reloaded_sld.xpath("./mc:AlternateContent/mc:Choice/p:timing")
        assert len(choice_timings) == 1
        wrapped_videos = reloaded_sld.xpath("./mc:AlternateContent/mc:Choice/p:timing//p:video")
        assert len(wrapped_videos) == 1

    # -- 3: the pre-existing unwrapped-timing path still works ---------

    def it_still_merges_into_a_plain_direct_p_timing(self):
        """A slide with a plain (unwrapped) ``p:timing`` must continue
        to merge cleanly — no regression on the non-#954 code path."""
        _prs, slide = _blank_slide()
        # -- inject a plain, direct-child `p:timing` with one existing
        # -- childTnLst entry so the new p:video is expected to join it --
        plain_timing_xml = (
            "<p:timing %s>\n"
            "  <p:tnLst>\n"
            "    <p:par>\n"
            '      <p:cTn id="1" dur="indefinite" restart="never" '
            'nodeType="tmRoot">\n'
            "        <p:childTnLst/>\n"
            "      </p:cTn>\n"
            "    </p:par>\n"
            "  </p:tnLst>\n"
            "</p:timing>" % nsdecls("p")
        )
        slide._element.append(parse_xml(plain_timing_xml))
        video = io.BytesIO(b"\x00" * 128)

        slide.shapes.add_movie(
            video,
            left=Inches(1),
            top=Inches(1),
            width=Inches(4),
            height=Inches(3),
        )

        sld = slide._element
        # -- exactly one plain `p:timing`, no mc:AlternateContent --
        assert len(sld.findall(qn("p:timing"))) == 1
        assert sld.find(qn("mc:AlternateContent")) is None
        # -- the video sits inside that timing's childTnLst --
        videos = sld.xpath("./p:timing//p:video")
        assert len(videos) == 1

    # -- 4: multi-movie + wrapped timing round-trip --------------------

    def it_handles_multiple_add_movie_calls_on_a_wrapped_slide(self):
        """Calling ``add_movie`` twice on the same wrapped-timing slide
        keeps both videos inside the wrapped timing — still no duplicate
        ``p:timing`` emitted and both shape ids round-trip."""
        prs, slide = _blank_slide()
        _inject_mc_wrapped_morph_timing(slide)

        for _ in range(2):
            slide.shapes.add_movie(
                io.BytesIO(b"\x00" * 128),
                left=Inches(1),
                top=Inches(1),
                width=Inches(4),
                height=Inches(3),
            )

        # -- save + reopen mirrors how the bug originally surfaced --
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        reloaded_sld = prs2.slides[0]._element

        # -- still one timing, still wrapped --
        assert reloaded_sld.find(qn("p:timing")) is None
        choice_timings = reloaded_sld.xpath("./mc:AlternateContent/mc:Choice/p:timing")
        assert len(choice_timings) == 1
        # -- both p:video entries merged into the same childTnLst --
        wrapped_videos = reloaded_sld.xpath("./mc:AlternateContent/mc:Choice/p:timing//p:video")
        assert len(wrapped_videos) == 2
        # -- the childTnLst parent of each p:cTn that owns a p:video is
        # -- the single shared one inside the wrapped timing (i.e. the
        # -- two videos share a timing subtree, they don't live in
        # -- parallel ones) --
        parent_childTnLsts = {v.xpath("ancestor::p:childTnLst[1]")[0] for v in wrapped_videos}
        assert len(parent_childTnLsts) == 1
