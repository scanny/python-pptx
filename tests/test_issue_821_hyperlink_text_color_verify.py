# pyright: reportPrivateUsage=false

"""Regression test for issue #821 — "How to change hyperlink text color".

Issue #821 (https://github.com/scanny/python-pptx/issues/821) asked how a
caller can recolor the text of a hyperlinked run. The ask was long-standing
because PowerPoint renders hyperlinked text in the theme color-scheme's
``a:hlink`` entry regardless of any explicit ``a:solidFill`` set on the
run's ``a:rPr`` — so the intuitive ``run.font.color.rgb = ...`` silently
had no effect in PowerPoint once the run carried an ``a:hlinkClick``.

The related issue #940 shipped the public API that resolves #821:
:attr:`.Font.use_theme_hyperlink_color` — a tri-state property that, when
set to |False|, writes a python-pptx-owned ``a:extLst/a:ext`` marker under
the run's ``a:hlinkClick`` element to record the caller's intent that the
run's explicit color be preferred over the theme. The marker round-trips
through python-pptx; PowerPoint itself still honors the run's own
``a:solidFill`` color only when the theme's ``a:hlink`` entry is
additionally replaced at the theme layer, but the caller now has a
first-class, lossless surface to express "this hyperlink should be red"
from python-pptx alone.

The canonical authoring pattern is therefore::

    run = paragraph.add_run()
    run.text = "click me"
    run.hyperlink.address = "https://example.com"
    run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
    run.font.use_theme_hyperlink_color = False

#821 is verified-and-closed by #940. This suite pins the six scenarios the
#821 thread and its workarounds collectively emphasize:

  * explicit run-level color + ``use_theme_hyperlink_color = False`` reads
    back as the assigned RGB (the "red hyperlink" recipe);
  * the default (no ``use_theme_hyperlink_color`` override) leaves the run
    inheriting the theme hyperlink color — the theme's ``a:hlink`` entry
    is the source of truth;
  * toggling ``use_theme_hyperlink_color`` back and forth is idempotent
    and restores the initial state each time;
  * the full "red hyperlink" recipe survives ``Presentation.save`` +
    reopen (the #821 reporter's deck is saved and redistributed);
  * underline is preserved across a color change (a hyperlink's underline
    is the visual cue PowerPoint users rely on — accidentally dropping it
    would defeat the point of #821 even if color read correctly);
  * sibling runs with independent color settings stay independent — a
    hyperlinked run whose color was overridden does not leak its settings
    into a neighboring non-hyperlinked run in the same paragraph.

Any future regression that silently drops the ``{PY-PPTX-940}`` marker,
fails to preserve the run's explicit ``a:solidFill`` across save+reopen
on a hyperlinked run, or couples the color override to the underline
setting will reproduce a #821 symptom and be caught here.
"""

from __future__ import annotations

import io

import pytest

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_UNDERLINE
from pptx.oxml.ns import qn
from pptx.util import Inches


@pytest.fixture
def _restore_part_factory():
    """Guard against test-module pollution of ``PartFactory.part_type_for``.

    Mirrors the guard used by ``tests/test_issue_1013_placeholder_color_verify.py``
    and friends — ``Presentation(stream)`` reopen dispatches through
    ``PartFactory.part_type_for``, which earlier test modules overwrite with
    mocks and do not restore.
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


def _add_hyperlinked_textbox_run(prs, address="https://example.com", text="click me"):
    """Return a fresh ``(textbox, run)`` with ``run.hyperlink.address`` set.

    Keeps the per-scenario fixture boilerplate to a single line. The blank
    layout (``slide_layouts[6]``) is used so no inherited placeholder style
    interferes with the run-level color surface under test.
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    textbox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
    run = textbox.text_frame.paragraphs[0].add_run()
    run.text = text
    run.hyperlink.address = address
    return textbox, run


class DescribeIssue821HyperlinkTextColor:
    """#821 hyperlink text-color verify-and-close via #940.

    Pins that :attr:`.Font.use_theme_hyperlink_color` together with
    :attr:`.Font.color` provides the first-class surface the #821 thread
    asked for — a run-level color override on a hyperlinked run that
    round-trips through save + reopen.
    """

    # -- API surface ------------------------------------------------------

    def it_exposes_use_theme_hyperlink_color_on_Font(self):
        """Guard: the public API that resolves #821 is present.

        A caller reaching for a run-level hyperlink-color override needs
        :attr:`.Font.use_theme_hyperlink_color` — a refactor that removed
        or renamed it would silently reintroduce #821.
        """
        from pptx.text.text import Font

        assert hasattr(Font, "use_theme_hyperlink_color")

    # -- explicit red hyperlink (the canonical #821 recipe) --------------

    def it_recolors_a_hyperlinked_run_to_red_via_the_override_flag(self):
        """The canonical "how do I make hyperlink text red?" recipe.

        The #821 thread's lead question. Assigning
        ``font.color.rgb = RGBColor(255, 0, 0)`` together with
        ``font.use_theme_hyperlink_color = False`` expresses the caller's
        intent fully; reading back returns the explicit RGB, and the
        marker is present on the ``a:hlinkClick`` element.
        """
        prs = Presentation()
        _textbox, run = _add_hyperlinked_textbox_run(prs)

        run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
        run.font.use_theme_hyperlink_color = False

        # -- run-level color reads back as red --
        assert run.font.color.rgb == RGBColor(0xFF, 0x00, 0x00)
        # -- override flag reads back as False --
        assert run.font.use_theme_hyperlink_color is False
        # -- and the marker is on the hlinkClick element --
        rPr = run._r.get_or_add_rPr()
        hlinkClick = rPr.find(qn("a:hlinkClick"))
        assert hlinkClick is not None
        extLst = hlinkClick.find(qn("a:extLst"))
        assert extLst is not None
        exts = extLst.findall(qn("a:ext"))
        assert len(exts) == 1
        assert exts[0].get("uri") == "{PY-PPTX-940}"

    # -- default (no override) ------------------------------------------

    def it_reads_None_for_a_hyperlinked_run_when_no_override_is_set(self):
        """Default = theme hyperlink color applies.

        A hyperlinked run without any override must leave
        :attr:`.Font.use_theme_hyperlink_color` at |None| — that is the
        read-side signal that no marker is present and PowerPoint will
        apply the theme's ``a:hlink`` color. The theme color is
        ``(0, 0, 255)`` in the default template — assert that the theme
        still defines a color so the "whatever the theme defines" promise
        is intact, without pinning the run's rendered color (which
        PowerPoint, not python-pptx, resolves).
        """
        prs = Presentation()
        _textbox, run = _add_hyperlinked_textbox_run(prs)

        # -- no override marker present --
        assert run.font.use_theme_hyperlink_color is None
        # -- run-level color is "not set" (inherits from theme/style) --
        assert run.font.color.type is None
        # -- the theme's hlink entry is the source of truth PowerPoint uses --
        theme_colors = prs.slide_masters[0].theme_colors
        assert "hlink" in theme_colors
        # -- default template's theme hlink is (0, 0, 255); assert the
        # -- theme defines *some* hlink colour rather than hard-coding
        # -- blue (the test reads "whatever the theme defines"). --
        assert theme_colors["hlink"] is not None

    # -- toggle back and forth -------------------------------------------

    def it_toggles_use_theme_hyperlink_color_back_and_forth(self):
        """Idempotent on/off.

        A caller changing their mind mid-authoring must get a clean round
        trip — setting to ``False``, then ``True`` / ``None`` restores the
        initial state and leaves no stale marker XML behind. Any failure
        here would manifest as either duplicate markers accumulating or a
        stuck "off" state that doesn't re-enable theme color.
        """
        prs = Presentation()
        _textbox, run = _add_hyperlinked_textbox_run(prs)

        # -- starts out at None (theme applies) --
        assert run.font.use_theme_hyperlink_color is None

        # -- opt out --
        run.font.use_theme_hyperlink_color = False
        assert run.font.use_theme_hyperlink_color is False

        # -- opt back in via True (symmetric form) --
        run.font.use_theme_hyperlink_color = True
        assert run.font.use_theme_hyperlink_color is None
        # -- marker physically gone --
        rPr = run._r.get_or_add_rPr()
        hlinkClick = rPr.find(qn("a:hlinkClick"))
        assert hlinkClick is not None
        assert hlinkClick.find(qn("a:extLst")) is None

        # -- opt out again and back in via None (alternate form) --
        run.font.use_theme_hyperlink_color = False
        assert run.font.use_theme_hyperlink_color is False
        run.font.use_theme_hyperlink_color = None
        assert run.font.use_theme_hyperlink_color is None
        hlinkClick = rPr.find(qn("a:hlinkClick"))
        assert hlinkClick is not None
        assert hlinkClick.find(qn("a:extLst")) is None

    # -- round-trip through save + reopen --------------------------------

    def it_round_trips_the_red_hyperlink_recipe_through_save_and_reload(
        self, _restore_part_factory
    ):
        """The #821 reporter's deck is saved and redistributed.

        The full "red hyperlink" recipe — explicit RGB + override flag —
        must survive ``Presentation.save`` + reopen without dropping
        either the color or the marker. A regression that, say, strips
        the ``{PY-PPTX-940}`` marker during serialization would
        reintroduce #821 on every saved file.
        """
        prs = Presentation()
        _textbox, run = _add_hyperlinked_textbox_run(prs)
        run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
        run.font.use_theme_hyperlink_color = False

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        prs2 = Presentation(buf)

        run2 = prs2.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]
        assert run2.text == "click me"
        assert run2.hyperlink.address == "https://example.com"
        assert run2.font.color.rgb == RGBColor(0xFF, 0x00, 0x00)
        assert run2.font.use_theme_hyperlink_color is False

    # -- underline preserved across color change -------------------------

    def it_preserves_underline_across_a_hyperlink_color_change(self):
        """A hyperlink's underline is its visual cue.

        Toggling ``use_theme_hyperlink_color`` or assigning ``color.rgb``
        must not touch the ``a:rPr/@u`` attribute — the two surfaces are
        independent. A regression that coupled color to underline (e.g.
        by rebuilding the ``a:rPr`` element) would drop the underline on
        every #821 authoring call.
        """
        prs = Presentation()
        _textbox, run = _add_hyperlinked_textbox_run(prs)

        run.font.underline = MSO_UNDERLINE.SINGLE_LINE
        assert run.font.underline is True  # -- SINGLE_LINE reports as True --

        run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
        run.font.use_theme_hyperlink_color = False

        # -- underline still set --
        assert run.font.underline is True
        # -- color still set --
        assert run.font.color.rgb == RGBColor(0xFF, 0x00, 0x00)
        # -- override still set --
        assert run.font.use_theme_hyperlink_color is False

        # -- and a reverse toggle back to theme doesn't disturb underline --
        run.font.use_theme_hyperlink_color = True
        assert run.font.underline is True

    # -- hyperlinked + non-hyperlinked sibling runs ----------------------

    def it_keeps_hyperlinked_and_plain_sibling_runs_independent(self):
        """Two runs in the same paragraph don't bleed into each other.

        A paragraph with a hyperlinked red run and a plain sibling run
        must keep their colour settings independent. Setting the override
        flag on the hyperlinked run must not propagate up to the
        paragraph or sideways to the plain run, and a colour assigned to
        the plain sibling must remain its own.
        """
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        textbox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
        paragraph = textbox.text_frame.paragraphs[0]

        # -- run 0: hyperlinked, recolored red, opted out of theme --
        run_link = paragraph.add_run()
        run_link.text = "red link"
        run_link.hyperlink.address = "https://example.com"
        run_link.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
        run_link.font.use_theme_hyperlink_color = False

        # -- run 1: plain sibling, green, no hyperlink --
        run_plain = paragraph.add_run()
        run_plain.text = " and plain green"
        run_plain.font.color.rgb = RGBColor(0x00, 0x80, 0x00)

        # -- each run carries its own color --
        assert run_link.font.color.rgb == RGBColor(0xFF, 0x00, 0x00)
        assert run_plain.font.color.rgb == RGBColor(0x00, 0x80, 0x00)

        # -- only the hyperlinked run has the override flag ---
        assert run_link.font.use_theme_hyperlink_color is False
        # -- plain sibling has no hyperlink so flag reads None regardless --
        assert run_plain.font.use_theme_hyperlink_color is None

        # -- and only the hyperlinked run carries an a:hlinkClick --
        assert run_link._r.get_or_add_rPr().find(qn("a:hlinkClick")) is not None
        assert run_plain._r.get_or_add_rPr().find(qn("a:hlinkClick")) is None
