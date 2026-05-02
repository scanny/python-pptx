# pyright: reportPrivateUsage=false

"""Regression test for issue #765 — "effective font" discussion.

Issue #765 (https://github.com/scanny/python-pptx/issues/765) collected a
long-running discussion about the library exposing the *effective* font
of a run — the size / color / bold / italic / typeface PowerPoint would
actually render, accounting for the full OOXML inheritance chain:

    run ``a:rPr`` → paragraph ``a:pPr/a:defRPr`` → text-body
    ``a:lstStyle/a:lvl{N}pPr/a:defRPr`` → slide-master
    ``p:txStyles/p:{body,other,title}Style/a:lvl{N}pPr/a:defRPr`` →
    presentation ``p:defaultTextStyle``

Before the fork's Wave 6 / Wave 12 work, :attr:`Font.size`,
:attr:`Font.bold`, :attr:`Font.italic`, :attr:`Font.name`, and
:attr:`Font.color.rgb` only reflected the run's own ``a:rPr``. On a
placeholder whose size/color/typeface were inherited from the master —
which is the normal authoring pattern — every one of those properties
read back as |None| (or raised on ``color.rgb`` when the fill was a
theme-scheme reference). Callers who needed the rendered values were
forced to walk the ``a:rPr`` → paragraph → lstStyle → master chain by
hand and resolve scheme colors against the theme themselves.

This fork closes #765 via two fork-era shipments:

  * **Wave 6, #938** (``feat/issue-938-effective-font-color``,
    ``cec3edc1``) added :attr:`Font.effective_color` — an |RGBColor|
    resolver that walks the full chain above, resolves
    ``a:schemeClr`` references against the slide-master's theme, and
    applies any sibling ``a:lumMod`` / ``a:lumOff`` transforms.

  * **Wave 12, #378** (``feat/issue-378-effective-font-size``,
    ``ec6e3c8e``) added :attr:`Font.effective_size`,
    :attr:`Font.effective_bold`, :attr:`Font.effective_italic`, and
    :attr:`Font.effective_name` — the size / bold / italic / Latin
    typeface siblings of ``effective_color``, walking the same chain
    and returning the value PowerPoint would render (or |None| when
    no ancestor in the chain declares an explicit value).

Together, the five ``effective_*`` properties cover the exact
"effective font" surface the #765 discussion asked for. A caller who
previously had to reach into ``a:rPr`` XML can now write::

    size   = run.font.effective_size     # Length (EMU) or None
    color  = run.font.effective_color    # RGBColor or None
    bold   = run.font.effective_bold     # True / False / None
    italic = run.font.effective_italic   # True / False / None
    name   = run.font.effective_name     # str or None ('+mn-lt' etc.)

and get the rendered values back without any manual inheritance walk.

#765 is therefore verified-and-closed by #938 + #378. This suite pins
the complete ``effective_*`` surface end-to-end so any regression —
dropping a step in the chain, breaking scheme-color resolution in
Wave 6, losing the ``lvl``-matching in Wave 12 — fails loudly against
the reporter's framing rather than silently reverting #765.
"""

from __future__ import annotations

import io

import pytest

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt


@pytest.fixture
def _restore_part_factory():
    """Guard against test-module pollution of ``PartFactory.part_type_for``.

    Earlier test modules in the suite install mocks on
    :class:`~pptx.opc.package.PartFactory` and do not restore them;
    ``Presentation(stream)`` reopen dispatches through that factory
    so this fixture pins the mapping back to the production classes
    around any round-trip test.
    """
    from pptx.opc.constants import CONTENT_TYPE as CT
    from pptx.opc.package import PartFactory
    from pptx.parts.chart import ChartPart
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


def _retint_body_style_lvl1(prs, scheme_val):
    """Retarget the master's ``p:bodyStyle/a:lvl1pPr/a:defRPr/a:solidFill``
    at a different theme scheme color.

    The default template's ``bodyStyle`` ``schemeClr val="tx1"`` resolves
    to black — useful for the inherited-default case, but not
    distinctive on its own. Swapping it to, say, ``"accent1"`` lets a
    test pin the full "scheme color walked from master, resolved
    against theme" path with an unambiguously non-black RGB.
    """
    master_elt = prs.slide_masters[0]._element
    txStyles = master_elt.find(qn("p:txStyles"))
    assert txStyles is not None
    bodyStyle = txStyles.find(qn("p:bodyStyle"))
    assert bodyStyle is not None
    lvl1 = bodyStyle.find(qn("a:lvl1pPr"))
    assert lvl1 is not None
    defRPr = lvl1.find(qn("a:defRPr"))
    assert defRPr is not None
    solidFill = defRPr.find(qn("a:solidFill"))
    assert solidFill is not None
    schemeClr = solidFill.find(qn("a:schemeClr"))
    assert schemeClr is not None
    schemeClr.set("val", scheme_val)


class DescribeIssue765EffectiveFont:
    """#765 "effective font" verify-and-close via #938 + #378.

    Pins the five ``Font.effective_*`` properties end-to-end against
    the reporter's framing: a run whose formatting is inherited from
    the placeholder / paragraph / master chain reports the values
    PowerPoint would actually render, not the |None| the legacy
    ``Font`` surface returns.

    Cross-reference:

    * :attr:`.Font.effective_color` — Wave 6 / issue #938
    * :attr:`.Font.effective_size`, :attr:`.Font.effective_bold`,
      :attr:`.Font.effective_italic`, :attr:`.Font.effective_name` —
      Wave 12 / issue #378
    """

    # -- effective_size: placeholder → master body-style inheritance ------

    def it_resolves_effective_size_on_a_placeholder_via_master(self):
        """A body placeholder whose run carries no explicit size resolves
        through the master's ``p:bodyStyle/a:lvl1pPr/a:defRPr/@sz``.

        The bundled default master sets ``sz="3200"`` on ``bodyStyle``
        lvl1 — the #765 reporter's "what would PowerPoint render?"
        question answered by the #378 chain walk.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        body = slide.placeholders[1]
        body.text_frame.text = "Hello body"
        run = body.text_frame.paragraphs[0].runs[0]

        # -- legacy Font.size still reports "not set on run" --
        assert run.font.size is None
        # -- #378 resolves the chain all the way to the master --
        assert run.font.effective_size == Pt(32)

    # -- effective_color: master/theme walk through a placeholder --------

    def it_resolves_effective_color_through_theme_master_placeholder(self):
        """A body placeholder's inherited color walks the theme-scheme
        chain: ``master bodyStyle/lvl1pPr/defRPr/solidFill/schemeClr`` is
        resolved against the slide-master's theme.

        Retargeting the master's scheme entry from ``tx1`` to
        ``accent1`` produces a non-black RGB, pinning the scheme-color
        hop the #765 discussion flagged as the hardest manual step.
        """
        prs = Presentation()
        # -- default theme's accent1 is (79, 129, 189) --
        _retint_body_style_lvl1(prs, "accent1")
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        body = slide.placeholders[1]
        body.text_frame.text = "Accent1 text"
        run = body.text_frame.paragraphs[0].runs[0]

        # -- legacy .color is "not set" on an inherited-color run --
        assert run.font.color.type is None
        # -- #938 resolves the theme-scheme reference to a concrete RGB --
        assert run.font.effective_color == RGBColor(79, 129, 189)

    # -- effective_bold / effective_italic: paragraph-level defRPr -------

    def it_resolves_effective_bold_and_italic_from_paragraph_defRPr(self):
        """Setting ``paragraph.font.bold`` / ``.italic`` writes onto
        ``a:pPr/a:defRPr`` — not onto any run's ``a:rPr``. The run's
        own ``.bold`` / ``.italic`` stay |None|; ``effective_bold`` and
        ``effective_italic`` walk the one step up and return the
        paragraph-style default.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        body = slide.placeholders[1]
        body.text_frame.text = "Bold italic"
        paragraph = body.text_frame.paragraphs[0]
        paragraph.font.bold = True
        paragraph.font.italic = True

        run = paragraph.runs[0]
        # -- the run's own rPr is untouched --
        assert run.font.bold is None
        assert run.font.italic is None
        # -- #378 walks to the paragraph defRPr --
        assert run.font.effective_bold is True
        assert run.font.effective_italic is True

    # -- effective_name: master minor-Latin typeface ---------------------

    def it_resolves_effective_name_to_master_minor_latin(self):
        """A placeholder run with no explicit ``a:latin`` resolves to
        the master's minor Latin typeface placeholder — ``"+mn-lt"`` in
        the default template. PowerPoint dereferences that against the
        theme's ``a:minorFont``; the library surfaces the placeholder
        string directly so callers can distinguish "inherited theme
        typeface" from a hard-coded font.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        body = slide.placeholders[1]
        body.text_frame.text = "Typeface"
        run = body.text_frame.paragraphs[0].runs[0]

        # -- legacy Font.name reports "not set on run" --
        assert run.font.name is None
        # -- #378 walks to the master's bodyStyle/lvl1pPr/defRPr/latin --
        assert run.font.effective_name == "+mn-lt"

    # -- explicit overrides beat inherited values -------------------------

    def it_prefers_explicit_run_overrides_over_inherited_values(self):
        """The inheritance chain is most-specific-wins. A run that sets
        its own size / color / bold / italic / name overrides anything
        inherited from the paragraph or the master — matching how
        PowerPoint resolves the chain, and keeping the #765 API usable
        for templating callers that selectively override per-run.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        body = slide.placeholders[1]
        body.text_frame.text = "Override"
        paragraph = body.text_frame.paragraphs[0]

        # -- paragraph-level inherited defaults (would otherwise win) --
        paragraph.font.bold = True
        paragraph.font.italic = True
        paragraph.font.size = Pt(12)
        paragraph.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)  # red

        # -- run-level explicit overrides --
        run = paragraph.runs[0]
        run.font.bold = False
        run.font.italic = False
        run.font.size = Pt(48)
        run.font.name = "Georgia"
        run.font.color.rgb = RGBColor(0x00, 0x00, 0xFF)  # blue

        assert run.font.effective_size == Pt(48)
        assert run.font.effective_bold is False
        assert run.font.effective_italic is False
        assert run.font.effective_name == "Georgia"
        assert run.font.effective_color == RGBColor(0x00, 0x00, 0xFF)

    # -- every effective_* returns None on a bare Font --------------------

    def it_returns_None_when_nothing_in_the_chain_declares_a_value(self):
        """When nothing along the chain declares a value for a given
        attribute, every ``effective_*`` property returns |None|. This
        is the signal callers use to distinguish "PowerPoint inherits
        from its own default" from "the file explicitly says bold/…".

        Exercised with a bare ``Font(a:rPr)`` — no parent part, no
        enclosing paragraph, no master — which is the simplest chain
        that reaches exhaustion without any ancestor declaring a
        value.
        """
        from pptx.oxml import parse_xml
        from pptx.text.text import Font

        rPr_xml = '<a:rPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'
        font = Font(parse_xml(rPr_xml))

        assert font.effective_size is None
        assert font.effective_color is None
        assert font.effective_bold is None
        assert font.effective_italic is None
        assert font.effective_name is None

    # -- round-trip: effective_* survives save + reopen -------------------

    def it_round_trips_effective_values_through_save_and_reload(self, _restore_part_factory):
        """The #765 reporter framed the ask around reading already-saved
        decks. All five ``effective_*`` must survive a ``Presentation
        .save`` + reopen cycle so a reloaded run that relies on
        placeholder / master / theme inheritance still resolves the
        same values it did before the save.
        """
        prs = Presentation()
        _retint_body_style_lvl1(prs, "accent1")
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        body = slide.placeholders[1]
        body.text_frame.text = "Round trip"
        paragraph = body.text_frame.paragraphs[0]
        # -- bold / italic land on the paragraph's pPr/defRPr --
        paragraph.font.bold = True
        paragraph.font.italic = True

        # -- confirm pre-save state for the test's own sanity --
        run = paragraph.runs[0]
        assert run.font.effective_size == Pt(32)
        assert run.font.effective_color == RGBColor(79, 129, 189)
        assert run.font.effective_bold is True
        assert run.font.effective_italic is True
        assert run.font.effective_name == "+mn-lt"

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)
        body2 = prs2.slides[0].placeholders[1]
        run2 = body2.text_frame.paragraphs[0].runs[0]

        # -- text content survives save/reopen (sanity) --
        assert run2.text == "Round trip"
        # -- and every effective_* resolves the same value --
        assert run2.font.effective_size == Pt(32)
        assert run2.font.effective_color == RGBColor(79, 129, 189)
        assert run2.font.effective_bold is True
        assert run2.font.effective_italic is True
        assert run2.font.effective_name == "+mn-lt"

    # -- effective_size walks past paragraph+lstStyle to the textbox ------

    def it_resolves_effective_size_on_a_textbox_through_the_master(self):
        """A library-created textbox (no placeholder inheritance) still
        resolves through the master's ``p:bodyStyle`` when the slide
        layout carries no competing ``lstStyle`` / ``pPr/defRPr``.

        Pins the "no paragraph-level defaults, no lstStyle, walk to
        master" branch of the chain — the fall-through case the #378
        implementation adds on top of the #938 chain-walk helper.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])  # "Title Only"
        box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
        tf = box.text_frame
        tf.text = "plain textbox"
        run = tf.paragraphs[0].runs[0]

        # -- nothing on the run, nothing on the paragraph, nothing on
        # -- the textbox's lstStyle; walk falls through to the master --
        assert run.font.size is None
        assert run.font.effective_size == Pt(32)
        assert run.font.effective_name == "+mn-lt"
